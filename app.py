import sys
import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
os.environ["PYTHONIOENCODING"] = "utf-8"
if hasattr(sys.stdout, 'buffer'):
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

open("startup_test.txt", "w").write("app.py loaded OK")
"""
建具図面チェッカー — ローカル動作版
対応フォーマット:
  建具・床図面: WD① WD② WD③ ... 形式（丸囲み数字）
  木工事図面:   WD101 WD102 ... 形式（3桁数字）※CIDフォントはClaude APIビジョンで読み取り
"""
import re
import io
import os
import sys
import base64
import unicodedata
import pdfplumber
import fitz  # PyMuPDF（PDF→画像変換用）
from flask import Flask, request, jsonify, render_template


# Kangxi Radicals (U+2F00..U+2FD5) と CJK Radicals Supplement (U+2E80..U+2EF3)
# はそれぞれ通常のCJK統合漢字に compatibility decomposition を持つ。
# CIDフォント由来で 色→⾊・衣→⾐・面→⾯・手→⼿ 等に化けるので、
# これらの範囲の文字だけ NFKC で正規化する（丸囲み数字 ①→1 は壊さない）。
def _is_radical_char(ch):
    c = ord(ch)
    return (0x2F00 <= c <= 0x2FD5) or (0x2E80 <= c <= 0x2EF3)


# 部首文字ではない（NFKCで吸収されない）異体字／旧字体の対応表。
# 建具図面で観測された/観測されそうな差異のみ追加する。
_VARIANT_MAP = {
    "戶": "戸",  # U+6236 (旧字体・繁体字) → U+6238 (日本標準)
    "户": "戸",  # U+6237 (簡体字)         → U+6238
}


def _normalize_text(s):
    """CIDフォント由来の部首文字や旧字体／異体字を通常字へ戻す。
    例: 敷居⾊∕種類 → 敷居色／種類、片引戶 → 片引戸。
    NFKC全体は丸囲み数字（①→1）を壊すので、部首文字＋既知の異体字だけ選択的に正規化。"""
    if not s:
        return s
    if any(_is_radical_char(ch) for ch in s) or "∕" in s or any(k in s for k in _VARIANT_MAP):
        out = []
        for ch in s:
            if _is_radical_char(ch):
                out.append(unicodedata.normalize("NFKC", ch))
            elif ch == "∕":
                out.append("／")
            elif ch in _VARIANT_MAP:
                out.append(_VARIANT_MAP[ch])
            else:
                out.append(ch)
        s = "".join(out)
    return s

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

# ─── APIキー自動読み込み ──────────────────────────────────────────────────────
def _load_api_key():
    """環境変数 ANTHROPIC_API_KEY からキーを読み込む"""
    return os.environ.get("ANTHROPIC_API_KEY", "")

# ─── 丸囲み数字ユーティリティ ────────────────────────────────────────────────

CIRCLE_CHARS = "".join(chr(c) for c in
    list(range(0x2460, 0x2474)) +   # ①-⑳
    list(range(0x3251, 0x3260)) +   # ㉑-㉟
    list(range(0x32B1, 0x32C0))     # ㊱-㊿
)
WD_CIRCLE_PAT = re.compile(r"WD([" + CIRCLE_CHARS + r"])")
WD_DIGIT_PAT  = re.compile(r"WD(\d{1,3})\b")

# 把手デザインのキーワード（長いものを先に並べて部分マッチを防ぐ）
HANDLE_KEYWORDS = sorted([
    "スクエアL", "スクエアM", "スクエアS",
    "引手", "手掛け", "レバーハンドル", "プッシュプル", "ツマミ",
], key=len, reverse=True)
_HANDLE_KW_PAT = re.compile("|".join(re.escape(k) for k in HANDLE_KEYWORDS))


def circle_to_int(ch):
    c = ord(ch)
    if 0x2460 <= c <= 0x2473: return c - 0x245F        # ①(1) - ⑳(20)
    if 0x3251 <= c <= 0x325F: return c - 0x3250 + 21   # ㉑(21) - ㉟(35)
    if 0x32B1 <= c <= 0x32BF: return c - 0x32B0 + 37   # ㊱(37) - ㊿(50)
    return None


def normalize_mokuko_wd(raw):
    """WD101 → ('WD1','1F'), WD202 → ('WD2','2F'), WD1 → ('WD1', None)"""
    raw = raw.strip().upper()
    m3 = re.match(r"WD(\d)(\d{2})$", raw)
    if m3:
        floor = m3.group(1)
        door  = str(int(m3.group(2)))
        return f"WD{door}", f"{floor}F"
    m1 = re.match(r"WD(\d+)$", raw)
    if m1:
        return f"WD{int(m1.group(1))}", None
    return raw, None


# ─── テキスト抽出 ────────────────────────────────────────────────────────────

def read_pdf_text(pdf_bytes):
    """ページ番号付きで全テキストを返す。CIDフォントは空文字。"""
    pages = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                pages.append({"page": i + 1, "text": _normalize_text(text)})
    except Exception as e:
        pages.append({"page": 1, "text": "", "error": str(e)})
    return pages


# ─── 建具・床図面パーサー ────────────────────────────────────────────────────

def split_by_keyword(line, keyword):
    """'種類 A 種類 B 種類 C' → ['A','B','C']"""
    parts = re.split(re.escape(keyword) + r"\s*", line)
    return [p.strip() for p in parts[1:] if p.strip()]


def find_floor_label(text):
    # 半角・全角数字どちらにも対応
    m = re.search(r"([\d０-９])\s*階平面図", text)
    if not m:
        return None
    d = m.group(1)
    # 全角数字を半角化
    if "０" <= d <= "９":
        d = chr(ord(d) - ord("０") + ord("0"))
    return f"{d}F"


def _cluster_words_by_row(words, y_tolerance=10):
    """座標付き単語リストを行ごとにグループ化して返す（y座標が近いものを同一行とみなす）"""
    if not words:
        return []
    sorted_words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    rows = []
    current_row = [sorted_words[0]]
    for w in sorted_words[1:]:
        if abs(w["top"] - current_row[0]["top"]) <= y_tolerance:
            current_row.append(w)
        else:
            rows.append(current_row)
            current_row = [w]
    rows.append(current_row)
    return rows


def _assign_sill_with_coords(sill_row_words, group, n_wd, page_words):
    """座標ベースで「有/無」を各WD列に割り当てる。成功した場合Trueを返す。"""
    if not sill_row_words or not page_words:
        return False
    hdrs = sorted(sill_row_words, key=lambda w: w["x0"])[:n_wd]
    col_x_centers = [(w["x0"] + w["x1"]) / 2 for w in hdrs]
    hdr_y_bottom = max(w["bottom"] for w in hdrs)
    sill_val_words = [
        w for w in page_words
        if w["text"] in ("有", "無")
        and hdr_y_bottom < w["top"] < hdr_y_bottom + 150
    ]
    if not sill_val_words:
        return False
    for w in sill_val_words:
        mid_x = (w["x0"] + w["x1"]) / 2
        nearest = min(range(len(col_x_centers)), key=lambda ci: abs(col_x_centers[ci] - mid_x))
        if nearest < n_wd and not group[nearest]["sill"]:
            group[nearest]["sill"] = w["text"]
    return True


def _assign_sill_color_heuristic(val, group, n_wd):
    """
    「枠色は建具仕様をご確認ください」の出現回数で列インデックスを算出し、
    PL色を対応するWD列に割り当てる。
    例: "枠色... 枠色... PL ペール 枠色..." → PL ペールはインデックス1（WD②）
    """
    ANCHOR = "枠色は建具仕様をご確認ください"
    pl_matches = list(re.finditer(r"PL\s+\S+", val))
    if pl_matches:
        for pm in pl_matches:
            prefix = val[:pm.start()]
            col_idx = prefix.count(ANCHOR) - 1
            if col_idx < 0:
                col_idx = 0  # アンカーなしの場合は先頭列
            if col_idx < n_wd:
                group[col_idx]["sill_color"] = pm.group()
    else:
        color_cands = re.findall(r"[^\s]{2,20}", val)
        for ci, cc in enumerate(color_cands[:n_wd]):
            group[ci]["sill_color"] = cc


def parse_tategu_pdf(pdf_bytes):
    """建具・床図面 → WDエントリ辞書 (key: 'WD1_1F' など)"""
    entries = {}

    try:
        pdf_obj = pdfplumber.open(io.BytesIO(pdf_bytes))
    except Exception:
        pdf_obj = None

    # ページ単位の座標付き単語データを事前取得
    page_words_map = {}
    sill_rows_map  = {}
    if pdf_obj:
        try:
            for pi, pg in enumerate(pdf_obj.pages):
                try:
                    pw = pg.extract_words(x_tolerance=5, y_tolerance=5) or []
                except Exception:
                    pw = []
                # CIDフォント由来の代用文字を正規化
                for w in pw:
                    if "text" in w:
                        w["text"] = _normalize_text(w["text"])
                page_words_map[pi + 1] = pw
                sill_hdrs = [w for w in pw if w["text"] == "敷居有無"]
                sill_rows_map[pi + 1] = _cluster_words_by_row(sill_hdrs)
        except Exception:
            pass
        pdf_obj.close()

    pages = read_pdf_text(pdf_bytes)

    for page_data in pages:
        pnum  = page_data["page"]
        text  = page_data["text"]
        lines = text.splitlines()

        floor = find_floor_label(text) or f"p{pnum}"

        page_words = page_words_map.get(pnum, [])
        sill_rows  = sill_rows_map.get(pnum, [])
        block_sill_idx = 0  # このページで何番目のWDブロックか

        i = 0
        while i < len(lines):
            line = lines[i]

            # WD定義行を検索（丸囲み数字）
            wd_matches = list(WD_CIRCLE_PAT.finditer(line))
            if not wd_matches:
                i += 1
                continue

            # ─ グループのWD番号と部屋名を抽出 ─
            group = []
            for mi, m in enumerate(wd_matches):
                n = circle_to_int(m.group(1))
                if not n:
                    continue
                wd_key = f"WD{n}"
                pos_start = m.end()
                pos_end   = wd_matches[mi + 1].start() if mi + 1 < len(wd_matches) else len(line)
                room = line[pos_start:pos_end].strip()
                group.append({
                    "key": wd_key,
                    "full_key": f"{wd_key}_{floor}",
                    "raw": m.group(0),
                    "room": room,
                    "floor": floor,
                    "page": pnum,
                    "type": "", "w": "", "h": "",
                    "handle": "", "sill": "", "sill_color": "", "hinban": "",
                })

            if not group:
                i += 1
                continue

            # ─ 次のWD定義行までをブロックとして収集 ─
            i += 1
            block = []
            while i < len(lines):
                if WD_CIRCLE_PAT.search(lines[i]):
                    break
                block.append(lines[i])
                i += 1

            n_wd = len(group)

            # 「大きさ」行: "大きさ W654 H2035 大きさ W778 H2035 大きさ H1961"
            # 前の行にstandalone W###があるケース (WD③のようにW/Hが別行の場合)
            pending_w = {}   # {index: W値}

            for bi, bl in enumerate(block):

                # 種類（「敷居色／種類」ヘッダー行は除外）
                if "種類" in bl and "敷居色" not in bl:
                    types = split_by_keyword(bl, "種類")
                    for ti, tv in enumerate(types):
                        # 空・"敷居色/"・ヘッダーっぽい値は除外
                        if ti < n_wd and not group[ti]["type"] and tv and "敷居色" not in tv:
                            group[ti]["type"] = tv[:40]

                # 品番
                if "品番" in bl:
                    hinbans = split_by_keyword(bl, "品番")
                    for hi, hv in enumerate(hinbans):
                        if hi < n_wd and not group[hi]["hinban"]:
                            group[hi]["hinban"] = hv[:50]

                # 大きさ
                if "大きさ" in bl:
                    size_sections = re.split(r"大きさ\s*", bl)
                    for si, sp in enumerate(size_sections[1:]):
                        if si >= n_wd:
                            break
                        wm = re.search(r"W\s*(\d{2,4})", sp)
                        hm = re.search(r"H\s*(\d{2,4}|特注)", sp)
                        if wm:
                            group[si]["w"] = wm.group(1)
                        elif si in pending_w:
                            group[si]["w"] = pending_w[si]
                        if hm:
                            group[si]["h"] = hm.group(1)

                # standalone W### (次の大きさ行と対応するケース)
                sw = re.match(r"^\s*(W\s*\d{2,4})\s*$", bl)
                if sw:
                    w_val = re.search(r"(\d{2,4})", sw.group(1)).group(1)
                    pending_w[n_wd - 1] = w_val

            # 把手デザイン
            sill_hdr_idx = None
            sill_col_hdr_idx = None
            handle_hdr_idx = None

            for bi, bl in enumerate(block):
                if "敷居有無" in bl and sill_hdr_idx is None:
                    sill_hdr_idx = bi
                if ("敷居色" in bl or "敷居色／種類" in bl) and sill_col_hdr_idx is None:
                    sill_col_hdr_idx = bi
                if "把手デザイン" in bl and handle_hdr_idx is None:
                    handle_hdr_idx = bi

            # 敷居有無（座標ベースで列判定、失敗時はテキストベースにフォールバック）
            if sill_hdr_idx is not None:
                sill_row_for_block = sill_rows[block_sill_idx] if block_sill_idx < len(sill_rows) else None
                coord_ok = _assign_sill_with_coords(sill_row_for_block, group, n_wd, page_words)
                if not coord_ok:
                    for bi in range(sill_hdr_idx + 1, min(sill_hdr_idx + 4, len(block))):
                        val = block[bi].strip()
                        if val and "敷居" not in val and "種類" not in val:
                            vals = re.findall(r"有|無", val)
                            for si, sv in enumerate(vals):
                                if si < n_wd:
                                    group[si]["sill"] = sv
                            break

            # 敷居色（「枠色は建具仕様をご確認ください」出現回数で列インデックスを特定）
            if sill_col_hdr_idx is not None:
                for bi in range(sill_col_hdr_idx + 1, min(sill_col_hdr_idx + 4, len(block))):
                    val = block[bi].strip()
                    if val and "敷居" not in val:
                        _assign_sill_color_heuristic(val, group, n_wd)
                        break

            # 把手（列ごとに「把手デザイン→値→STOP」が繰り返す構造にも対応）
            if handle_hdr_idx is not None:
                all_handles = []
                STOP_WORDS = {"敷居", "種類", "大きさ", "品番", "備考", "建具仕様"}
                collecting = True  # 把手デザインヘッダー直後はTrue
                for bi in range(handle_hdr_idx + 1, len(block)):
                    if len(all_handles) >= n_wd:
                        break
                    val = block[bi].strip()
                    if not val:
                        continue
                    # 次のWD列の「把手デザイン」ヘッダー → 収集再開
                    if "把手デザイン" in val:
                        collecting = True
                        continue
                    if not collecting:
                        continue
                    if any(sw in val for sw in STOP_WORDS):
                        collecting = False  # このWD列の把手セクション終了
                        continue
                    handles = re.findall(r"[" + CIRCLE_CHARS + r"][^\s" + CIRCLE_CHARS + r"]*", val)
                    if not handles:
                        handles = _HANDLE_KW_PAT.findall(val)
                    if handles:
                        all_handles.extend(handles)
                    elif not all_handles:
                        all_handles.append(val[:30])
                for hi, hv in enumerate(all_handles[:n_wd]):
                    if not group[hi]["handle"]:
                        group[hi]["handle"] = hv

            # エントリ登録（種類・W・Hがすべて空 → 斜線欄とみなしてスキップ）
            for wd in group:
                if not any([wd["type"], wd["w"], wd["h"]]):
                    continue
                entries[wd["full_key"]] = wd

            block_sill_idx += 1

    return entries


# ─── OCR（CIDフォントPDF対応） ──────────────────────────────────────────────

_ocr_reader = None  # easyocr.Reader のシングルトン


def get_ocr_reader():
    global _ocr_reader
    if _ocr_reader is None:
        import easyocr
        _ocr_reader = easyocr.Reader(["ja", "en"], gpu=False)
    return _ocr_reader


def is_cid_font_pdf(text):
    """テキスト抽出結果がCIDフォント由来かチェック"""
    if not text:
        return True
    cid_count = text.count("(cid:")
    total     = max(len(text), 1)
    # CIDが多い、またはほぼ全文が制御文字
    readable = sum(1 for c in text if c.isprintable() and ord(c) > 31)
    return cid_count > 5 or readable / total < 0.3


def pdf_to_images(pdf_bytes, dpi=300):
    """PyMuPDF で PDF各ページをPNG画像(bytes)に変換（高解像度）"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    zoom = dpi / 72
    mat  = fitz.Matrix(zoom, zoom)
    for page in doc:
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
        images.append(pix.tobytes("png"))
    return images


def ocr_pdf(pdf_bytes):
    """CIDフォントPDFをOCRで読み取り、全テキストを返す"""
    import sys, io as _io
    # Windows CP932端末でのUnicode出力エラーを回避
    old_stdout = sys.stdout
    sys.stdout = _io.TextIOWrapper(_io.BytesIO(), encoding="utf-8")
    try:
        reader = get_ocr_reader()
    finally:
        sys.stdout = old_stdout

    images = pdf_to_images(pdf_bytes)
    all_text = []
    for i, img_bytes in enumerate(images):
        results = reader.readtext(img_bytes, detail=0, paragraph=False)
        # Unicode文字を安全に処理
        cleaned = []
        for r in results:
            if isinstance(r, str):
                cleaned.append(r)
        page_text = "\n".join(cleaned)
        all_text.append({"page": i + 1, "text": page_text})
    return all_text


# ─── Claude API ビジョンOCR ──────────────────────────────────────────────────

def ocr_images_with_claude(image_list, api_key):
    """画像バイト列のリストをClaude APIビジョンで読み取り、WDテーブルテキストを返す"""
    try:
        import anthropic
    except ImportError:
        raise RuntimeError("anthropic パッケージが見つかりません。`pip install anthropic` を実行してください。")

    client = anthropic.Anthropic(api_key=api_key)
    prompt_text = (
        "この画像は木工事図面（建具スケジュール表）です。\n"
        "以下の手順で、表に記載されたWD番号・種類・W寸法・H寸法を正確に読み取ってください。\n\n"
        "【読み取り手順】\n"
        "手順1: 表の一番上の行にあるWD番号（WD101, WD102, WD201など）を左から右へ順番に確認する。\n"
        "手順2: 各WD番号の「列」を画像上で垂直に追う。その列の中にある種類・W寸法・H寸法を読む。\n"
        "手順3: 隣の列と混同しないよう、WD番号と寸法が同じ列に属していることを確認してから出力する。\n\n"
        "【出力形式】1つのWD番号につき1行、カンマ区切りで出力（実際の値を画像から読み取ること）:\n"
        "WDxxx,種類名,W寸法,H寸法\n\n"
        "【厳守ルール】\n"
        "1. カンマ区切り形式のみ（スペース区切り不可）\n"
        "2. 1つのWD番号につき必ず1行のみ\n"
        "3. WD番号は「WD」+3桁数字（例: WD101, WD201）\n"
        "4. 種類の表記:\n"
        "   - 蝶番で開く扉（弧を描く軌跡） → 「片開きドア」\n"
        "   - 横にスライドする扉（直線の軌跡） → 「片引き戸」\n"
        "   - アウトセット・壁外スライド → 「アウトセット引き戸」\n"
        "   - 上吊り・ハンガー式スライド → 「上吊り引き戸」\n"
        "   - 蛇腹状・2枚以上が折れる扉 → 「クローゼットドア」または「折戸」\n"
        "   ※「片引き戸」と「折戸（クローゼット）」は別物。混同禁止。\n"
        "   ※トイレ・洗面所のスライド扉も「片引き戸」または「アウトセット引き戸」。\n"
        "5. W寸法・H寸法は必ず【そのWD番号と同じ列】の値を読む。\n"
        "   絶対に隣の列の値を読まないこと。列がずれたまま出力することは最大の禁止事項。\n"
        "   W寸法は「W数字」形式（例: W654, W778, W1155, W1644）\n"
        "   H寸法は「H数字」または「H特注」形式（例: H2035）\n"
        "6. 読み取れない項目は空欄（例: WD103,片引き戸,,）\n"
        "7. WDスケジュール表以外（床材・備考・サッシ表・ヘッダー等）は出力しない\n"
        "8. 解説・説明文は一切書かない。データ行のみ出力する\n"
        "9. 「WD」が付いていない数字（101, 102等）はサッシ番号なので無視する\n"
        "10. WD1xx（1階）とWD2xx（2階）のみ対象。WD3xx以上は出力しない\n"
        "11. 2階建具スケジュール表の列構成（参考）:\n"
        "    1列目=WD201(片引戸,W1644,H2035)  2列目=WD202(片開きドア,W654,H2035)\n"
        "    3列目=WD203(片開きドア,W778,H2035)  4列目=WD204(片開きドア,W778,H2035)\n"
        "    ※WD201のみ片引き戸。WD202〜WD204は片開きドア。\n"
        "12. 1階建具スケジュール表は画像をそのまま読んで正確に出力すること。\n"
        "    列の対応を慎重に確認し、W寸法・H寸法を正確に読み取ること。"
    )

    all_text = []
    for img_bytes in image_list:
        # PNG/JPGどちらにも対応
        mime = "image/png"
        if img_bytes[:3] == b'\xff\xd8\xff':
            mime = "image/jpeg"
        img_b64 = base64.b64encode(img_bytes).decode()
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4096,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": mime, "data": img_b64}},
                    {"type": "text", "text": prompt_text}
                ]
            }]
        )
        all_text.append(response.content[0].text)
    return "\n".join(all_text)


def ocr_with_claude(pdf_bytes, api_key):
    """CIDフォントPDFをClaude APIビジョンで読み取り、WDテーブルテキストを返す"""
    try:
        import anthropic
    except ImportError:
        raise RuntimeError("anthropic パッケージが見つかりません。`pip install anthropic` を実行してください。")

    images = pdf_to_images(pdf_bytes, dpi=200)
    return ocr_images_with_claude(images, api_key)


# ─── 木工事図面パーサー ──────────────────────────────────────────────────────

def parse_mokuko_from_text(text):
    """テキストからWD10x エントリを抽出。
    カンマ区切り形式（Claude API返答）: WD101,片開きドア,W654,H2035
    スペース区切り形式: WD101 片開きドア W654 H2035
    多行テーブル形式（WD番号行 + 種類行 + W行 + H行）
    の3形式に対応。
    """
    entries = {}
    lines = text.splitlines()

    type_kws = ["片開きドア", "開き戸", "片引戸", "引戸", "引き戸",
                "折戸", "クローゼット", "上吊", "アウトセット", "ドア", "サッシ"]

    def extract_type_only(seg):
        """セグメントから種類キーワードのみ抽出（W/H寸法を除去）"""
        seg = re.sub(r"W\s*\d{3,4}", "", seg)
        seg = re.sub(r"H\s*(\d{3,4}|特注)", "", seg)
        seg = seg.strip()
        # 先頭の "W" を除去（例: "W片開き扉" → "片開き扉"）
        seg = re.sub(r"^W([^\d])", r"\1", seg)
        return seg.strip()

    # ──────────────────────────────────────────────────────────
    # パス1: カンマ区切り形式（Claude API推奨形式）
    #   WD101,片開きドア,W654,H2035
    # ──────────────────────────────────────────────────────────
    csv_found = False
    for line in lines:
        # "WD101,..." パターン
        m = re.match(r"(WD\d{3})\s*,\s*([^,]*),\s*(W[\d]+|)\s*,?\s*(H[\d]+|H特注|)", line.strip())
        if not m:
            # より緩いマッチ: WD番号とカンマが1つでもあれば
            m2 = re.match(r"(WD\d{3})\s*,(.+)", line.strip())
            if not m2:
                continue
            raw = m2.group(1)
            rest_csv = m2.group(2)
            norm, floor = normalize_mokuko_wd(raw)
            entry = {"raw": raw, "key": norm, "floor": floor, "type": "", "w": "", "h": ""}
            cols = [c.strip() for c in rest_csv.split(",")]
            for col in cols:
                if any(kw in col for kw in type_kws) and not entry["type"]:
                    entry["type"] = extract_type_only(col)
                w_m = re.search(r"W(\d{3,4})", col)
                if w_m and not entry["w"]:
                    entry["w"] = w_m.group(1)
                h_m = re.search(r"H(\d{3,4}|特注)", col)
                if h_m and not entry["h"]:
                    entry["h"] = h_m.group(1)
            key = f"{norm}_{floor}" if floor else norm
            entries[key] = entry
            csv_found = True
            continue
        raw = m.group(1)
        norm, floor = normalize_mokuko_wd(raw)
        type_str = extract_type_only(m.group(2).strip())
        w_str = re.sub(r"[^\d]", "", m.group(3))
        h_str = m.group(4).replace("H", "").strip()
        key = f"{norm}_{floor}" if floor else norm
        entries[key] = {"raw": raw, "key": norm, "floor": floor,
                        "type": type_str, "w": w_str, "h": h_str}
        csv_found = True

    if csv_found:
        return entries

    # ──────────────────────────────────────────────────────────
    # パス2: スペース区切り・多行形式
    # ──────────────────────────────────────────────────────────
    i = 0
    while i < len(lines):
        line = lines[i]
        wds = list(re.finditer(r"WD(\d{3})", line))
        if not wds:
            i += 1
            continue

        group = []
        for m in wds:
            raw  = m.group(0)
            norm, floor = normalize_mokuko_wd(raw)
            group.append({"raw": raw, "key": norm, "floor": floor,
                          "type": "", "w": "", "h": ""})

        # ケース1: 同じ行にデータ（"WD101 片開きドア W654 H2035"）
        rest = re.sub(r"WD\d{3}", "", line).strip()
        if rest:
            has_type = any(kw in rest for kw in type_kws)
            if has_type:
                parts = re.split(r"\s{2,}|\t", rest)
                if len(parts) == 1:
                    extracted = extract_type_only(parts[0])
                    if extracted and group:
                        group[0]["type"] = extracted
                else:
                    type_idx = 0
                    for tp in parts:
                        if type_idx < len(group) and any(k in tp for k in type_kws):
                            group[type_idx]["type"] = extract_type_only(tp)
                            type_idx += 1
            for wi, wv in enumerate(re.findall(r"W\s*(\d{3,4})", rest)):
                if wi < len(group) and not group[wi]["w"]:
                    group[wi]["w"] = wv
            for hi, hv in enumerate(re.findall(r"H\s*(\d{3,4}|特注)", rest)):
                if hi < len(group) and not group[hi]["h"]:
                    group[hi]["h"] = hv

        # ケース2: 次の行にデータ（多行テーブル形式）— 次のWD行で停止
        for j in range(i + 1, min(i + 6, len(lines))):
            next_line = lines[j]
            if re.search(r"WD\d{3}", next_line):
                break
            has_type = any(kw in next_line for kw in type_kws)
            if has_type:
                parts = re.split(r"\s{2,}|\t", next_line.strip())
                type_idx = 0
                for tp in parts:
                    if type_idx < len(group) and any(k in tp for k in type_kws) and not group[type_idx]["type"]:
                        group[type_idx]["type"] = extract_type_only(tp)
                        type_idx += 1
                if len(parts) == 1 and type_idx == 0 and group and not group[0]["type"]:
                    group[0]["type"] = extract_type_only(parts[0])
            for wi, wv in enumerate(re.findall(r"W\s*(\d{3,4})", next_line)):
                if wi < len(group) and not group[wi]["w"]:
                    group[wi]["w"] = wv
            for hi, hv in enumerate(re.findall(r"H\s*(\d{3,4}|特注)", next_line)):
                if hi < len(group) and not group[hi]["h"]:
                    group[hi]["h"] = hv

        for wd in group:
            key = f"{wd['key']}_{wd['floor']}" if wd["floor"] else wd["key"]
            entries[key] = wd

        i += 1
    return entries


def parse_mokuko_pdf(pdf_bytes, api_key=None):
    """木工事図面 → WDエントリ辞書
    テキスト抽出を試み、CIDフォントなら Claude API（api_key指定時）またはeasyocr にフォールバック"""
    pages = read_pdf_text(pdf_bytes)
    full_text = "\n".join(p["text"] for p in pages)

    ocr_used = False

    # pdfplumberで取れたテキストからWD番号を試行
    entries = parse_mokuko_from_text(full_text)

    # WDが1件も取れない場合、またはCIDフォント判定 → Claude APIで画像読み取り
    claude_raw = None
    cid_detected = is_cid_font_pdf(full_text)
    print(f"[parse_mokuko_pdf] entries_from_pdfplumber={len(entries)}, cid_detected={cid_detected}, api_key={'YES' if api_key else 'NO'}")
    if (len(entries) == 0 or cid_detected) and api_key:
        try:
            print("[parse_mokuko_pdf] Calling Claude API OCR...")
            claude_text = ocr_with_claude(pdf_bytes, api_key)
            print(f"[parse_mokuko_pdf] Claude returned {len(claude_text)} chars")
            print(f"[parse_mokuko_pdf] Claude text preview:\n{claude_text[:500]}")
            claude_raw = claude_text  # デバッグ用に保存
            full_text = claude_text
            ocr_used = True
            entries = parse_mokuko_from_text(full_text)
            print(f"[parse_mokuko_pdf] Parsed entries after Claude: {list(entries.keys())}")
        except Exception as e:
            print(f"[parse_mokuko_pdf] Claude API error: {e}")
            if len(entries) == 0:
                return {}, False, f"Claude APIエラー: {e}", None
            # エントリが既にあればエラーを無視して継続

    # WDが取れない かつ APIキーもない場合
    if len(entries) == 0 and not api_key:
        try:
            pages = ocr_pdf(pdf_bytes)
            full_text = "\n".join(p["text"] for p in pages)
            ocr_used = True
            entries = parse_mokuko_from_text(full_text)
        except Exception:
            return {}, False, (
                "木工事図面からWDデータを読み取れませんでした。\n"
                "「手動入力モード」でWDデータを貼り付けてください。"
            ), None

    # 丸囲み形式も対応（木工事図面がWD①形式の場合）
    for line in full_text.splitlines():
        for m in WD_CIRCLE_PAT.finditer(line):
            n = circle_to_int(m.group(1))
            if not n:
                continue
            norm = f"WD{n}"
            w_m  = re.search(r"W\s*(\d{3,4})", line)
            h_m  = re.search(r"H\s*(\d{3,4})", line)
            if norm not in entries:
                entries[norm] = {
                    "raw": m.group(0), "key": norm, "floor": None,
                    "type": "", "w": w_m.group(1) if w_m else "",
                    "h": h_m.group(1) if h_m else "",
                }

    return entries, ocr_used, None, claude_raw


# ─── 手動入力パーサー ────────────────────────────────────────────────────────

def parse_mokuko_manual(text):
    """
    手動入力テキストからWDエントリを解析。
    以下の形式に対応:
      WD101 片開きドア W654 H2035
      WD101,片開きドア,654,2035
      WD101　片開き戸　654　2035
    また「WD101 WD102 / 片開きドア 片開きドア / W654 W778 / H2035 H2035」
    のような複数列形式にも対応。
    """
    entries = {}
    if not text or not text.strip():
        return entries

    lines = [l.strip() for l in text.strip().splitlines() if l.strip()]

    type_kws = ["片開きドア", "開き戸", "片引戸", "引き戸", "引戸", "折戸",
                "クローゼット", "上吊", "アウトセット", "ドア", "サッシ", "引"]

    # ── 形式A: 1行にWD番号が複数並ぶ「表形式」 ──────────────────────────────
    wd_lines = [(i, l) for i, l in enumerate(lines)
                if re.search(r"WD\d{3}", l) and not re.search(r"W\d{3,4}", l)]

    if wd_lines:
        for row_idx, (li, wd_line) in enumerate(wd_lines):
            # WD番号を抽出
            wds = re.findall(r"WD(\d{3})", wd_line)
            if not wds:
                continue
            group = []
            for raw_digits in wds:
                norm, floor = normalize_mokuko_wd(f"WD{raw_digits}")
                group.append({"raw": f"WD{raw_digits}", "key": norm, "floor": floor,
                              "type": "", "w": "", "h": ""})

            # 次の数行からデータ取得
            for j in range(li + 1, min(li + 6, len(lines))):
                bl = lines[j]
                # 次のWD行が来たら終了
                if re.search(r"WD\d{3}", bl) and not re.search(r"W\d{3,4}", bl):
                    break

                # 種類
                parts = re.split(r"[\t　,\s]{2,}", bl.strip())
                if any(kw in bl for kw in type_kws):
                    for pi, p in enumerate(parts):
                        if pi < len(group) and any(kw in p for kw in type_kws):
                            group[pi]["type"] = p.strip()

                # W寸法
                w_vals = re.findall(r"W\s*(\d{3,4})", bl)
                for wi, wv in enumerate(w_vals):
                    if wi < len(group) and not group[wi]["w"]:
                        group[wi]["w"] = wv

                # H寸法
                h_vals = re.findall(r"H\s*(\d{3,4}|特注|[A-Za-z]+注?)", bl)
                for hi, hv in enumerate(h_vals):
                    if hi < len(group) and not group[hi]["h"]:
                        group[hi]["h"] = hv

            for wd in group:
                key = f"{wd['key']}_{wd['floor']}" if wd["floor"] else wd["key"]
                if wd["key"]:
                    entries[key] = wd
        if entries:
            return entries

    # ── 形式B: 1行1エントリ（カンマ区切り・スペース区切り） ───────────────────
    for line in lines:
        # WD番号を探す
        m = re.search(r"WD(\d{1,3})", line)
        if not m:
            continue
        raw = f"WD{m.group(1)}"
        norm, floor = normalize_mokuko_wd(raw)

        # 種類
        typ = ""
        for kw in type_kws:
            if kw in line:
                typ = kw; break

        # W, H
        w_m = re.search(r"W\s*(\d{3,4})", line)
        h_m = re.search(r"H\s*(\d{3,4}|特注)", line)
        # カンマ区切り数値もフォールバック
        nums = re.findall(r"\b(\d{3,4})\b", line)

        w = w_m.group(1) if w_m else (nums[0] if len(nums) > 0 else "")
        h = h_m.group(1) if h_m else (nums[1] if len(nums) > 1 else "")

        key = f"{norm}_{floor}" if floor else norm
        entries[key] = {"raw": raw, "key": norm, "floor": floor,
                        "type": typ, "w": w, "h": h}

    return entries


# ─── 建具仕様・床仕様パーサー ────────────────────────────────────────────────

def parse_building_spec(text):
    """
    建具・床図面の「建具仕様」「床仕様」セクションを抽出する。
    各ページに出現する場合は最初のものを採用。
    """
    spec = {
        "door_color": "",
        "stopper_color": "",
        "catcher_color": "",
        "frame_color": "",
        "handle_color": "",
        "door_color_exc": "",
        "frame_color_exc": "",
        "floor_material": "",
        "floor_code": "",
        "floor_direction": "",
    }

    lines = text.splitlines()
    in_spec = False

    for line in lines:
        if "建具仕様" in line:
            in_spec = True

        if not in_spec:
            continue

        if ("ドア色" in line or ("ドア" in line and "色" in line)) and not spec["door_color"]:
            m = re.search(r"ドア色\s+([^\s※④]+)", line)
            if m:
                spec["door_color"] = m.group(1).strip()
            exc = re.search(r"一部変更.*?[④①②③⑤]\s*(\S+)", line)
            if exc:
                spec["door_color_exc"] = exc.group(1).strip()

        if "ストッパー" in line and not spec["stopper_color"]:
            m = re.search(r"(?:ストッパー\s+色|ストッパー色|ストッパー)\s+(\S+)", line)
            if m:
                spec["stopper_color"] = m.group(1).strip()

        if "キャッチャー" in line and not spec["catcher_color"]:
            m = re.search(r"(?:キャッチャー\s+色|キャッチャー色|キャッチャー)\s+(\S+)", line)
            if m:
                spec["catcher_color"] = m.group(1).strip()

        if "枠色" in line and "枠色は建具仕様" not in line and not spec["frame_color"]:
            m = re.search(r"枠色\s+([^\s※④]+)", line)
            if m:
                spec["frame_color"] = m.group(1).strip()
            exc = re.search(r"一部変更.*?[④①②③⑤]\s*(\S+)", line)
            if exc:
                spec["frame_color_exc"] = exc.group(1).strip()

        if "ハンドル" in line and not spec["handle_color"]:
            m = re.search(r"(?:ハンドル\s+色|ハンドル色|ハンドル)\s+(\S+)", line)
            if m:
                spec["handle_color"] = m.group(1).strip()

        if "床の張り方向" in line and not spec["floor_direction"]:
            spec["floor_direction"] = line.strip()

        if "ウッドワン" in line or "FKK" in line:
            spec["floor_material"] = line.strip()[:60]
        if re.search(r"FKK\d+", line) and not spec["floor_code"]:
            m = re.search(r"(FKK[\w\-]+)", line)
            if m:
                spec["floor_code"] = m.group(1)

        if "和室" in line and in_spec and spec["door_color"]:
            break

    return spec


def check_building_spec(spec):
    ok = []
    errors = []

    if not spec["door_color"]:
        errors.append(
            "**建具仕様・ドア色**\n"
            "  - 現状: 建具図面の「建具仕様」欄からドア色を読み取れませんでした\n"
            "  - 理由: 目視確認が必要"
        )

    if not spec["floor_direction"]:
        errors.append(
            "**床の張り方向注記**\n"
            "  - 現状: 「床仕様」欄の床の張り方向注記を読み取れませんでした\n"
            "  - 理由: 目視確認が必要（矢印方向も含む）"
        )

    if spec["door_color"]:
        ok.append(f"建具仕様を読み取りました（ドア色={spec['door_color']}、詳細は下部テーブル参照）")

    return errors, ok


# ─── チェック関数 ────────────────────────────────────────────────────────────

def floors_match(f1, f2):
    """'1F'と'p1'、'2F'と'p2'など、異なる形式の階数表現を同一視する"""
    if f1 == f2:
        return True
    if not f1 or not f2:
        return False
    m_nf = re.match(r'(\d+)F$', f1)
    m_pn = re.match(r'p(\d+)$', f2)
    if m_nf and m_pn:
        return m_nf.group(1) == m_pn.group(1)
    m_pn = re.match(r'p(\d+)$', f1)
    m_nf = re.match(r'(\d+)F$', f2)
    if m_pn and m_nf:
        return m_pn.group(1) == m_nf.group(1)
    return False


# ─── 種類カテゴリ対応表（★修正箇所） ────────────────────────────────────────
# 木工事図面と建具図面で表記が異なっても同カテゴリなら合格とする
DOOR_TYPE_CATEGORIES = {
    "片開きドア": [
        "片開きドア", "片開き戸", "開き戸",
        "トイレドア", "標準ドア",
        "⑥トイレドア", "③標準ドア",
        "⑥", "③",
    ],
    "片引戸": [
        "片引戸", "片引き戸",           # 「片引」が付く名前
        "アウトセット引き戸", "アウトセット引戸", "アウトセット",  # 「アウトセット」が付く名前
        "上吊引き戸", "上吊引戸", "上吊",
        "⑱片引き戸", "⑨アウトセット",
        "⑱", "⑨",
    ],
    "折戸": [
        "折戸", "クローゼット", "ノンレール",
        "㊸クローゼット", "クローゼットドア",
        "㊸",
    ],
}


def get_door_category(type_str):
    """種類文字列をカテゴリに変換する。該当なしはNoneを返す"""
    if not type_str:
        return None
    for category, keywords in DOOR_TYPE_CATEGORIES.items():
        for kw in keywords:
            if kw in type_str:
                return category
    return None


def _sanitize(text):
    """丸囲み数字など cp932 で扱えない文字を ASCII に変換する"""
    circle_map = {
        '①':'1','②':'2','③':'3','④':'4','⑤':'5','⑥':'6','⑦':'7','⑧':'8','⑨':'9','⑩':'10',
        '⑪':'11','⑫':'12','⑬':'13','⑭':'14','⑮':'15','⑯':'16','⑰':'17','⑱':'18','⑲':'19','⑳':'20',
        '㉑':'21','㉒':'22','㉓':'23','㉔':'24','㉕':'25','㉖':'26','㉗':'27','㉘':'28','㉙':'29','㉚':'30',
        '㉛':'31','㉜':'32','㉝':'33','㉞':'34','㉟':'35','㊱':'36','㊲':'37','㊳':'38','㊴':'39','㊵':'40',
        '㊶':'41','㊷':'42','㊸':'43','㊹':'44','㊺':'45','㊻':'46','㊼':'47','㊽':'48','㊾':'49','㊿':'50',
    }
    for ch, rep in circle_map.items():
        text = text.replace(ch, rep)
    return text.encode('cp932', errors='replace').decode('cp932')


def check_cross_reference(mokuko_entries, tategu_entries):
    try:
        return _check_cross_reference_inner(mokuko_entries, tategu_entries)
    except UnicodeEncodeError:
        # エントリ内の文字列を sanitize してリトライ
        def san_entry(e):
            return {k: (_sanitize(v) if isinstance(v, str) else v) for k, v in e.items()}
        safe_mokuko = {k: san_entry(v) for k, v in mokuko_entries.items()}
        safe_tategu = {k: san_entry(v) for k, v in tategu_entries.items()}
        return _check_cross_reference_inner(safe_mokuko, safe_tategu)


def _check_cross_reference_inner(mokuko_entries, tategu_entries):
    errors, ok = [], []

    if not mokuko_entries:
        errors.append(
            "**【木工事図面 × 建具図面】寸法・種類の照合**\n"
            "  - 現状: 木工事図面からWDデータを読み取れませんでした（CIDフォントまたは画像PDF）\n"
            "  - 理由: 目視確認が必要 — 木工事図面の各WD番号の種類・W・H寸法を建具図面と手動で照合してください"
        )
        return errors, ok

    # 照合実行
    for mk_key, me in mokuko_entries.items():
        norm  = me["key"]    # 'WD1', 'WD2', ...
        floor = me.get("floor")  # '1F', '2F', None

        # ── 建具図面側で対応エントリを探す ──────────────────────────────
        te = None
        if floor:
            te = tategu_entries.get(f"{norm}_{floor}")
        if te is None:
            if floor:
                for v in tategu_entries.values():
                    if v["key"] == norm and floors_match(floor, v.get("floor")):
                        te = v
                        break
            else:
                for v in tategu_entries.values():
                    if v["key"] == norm:
                        te = v
                        break

        floor_label = f"（{floor}）" if floor else ""
        if te is None:
            errors.append(
                f"**{me['raw']}{floor_label} → {norm}{floor_label}**\n"
                f"  - 現状: 建具図面に {norm}{floor_label} が見当たらない\n"
                f"  - 理由: 建具番号が未登録またはテーブル読み取り不可"
            )
            continue

        # ── 寸法・種類の照合 ──────────────────────────────────────────
        issues = []
        mw = re.sub(r"[^\d]", "", me.get("w", ""))
        tw = re.sub(r"[^\d]", "", te.get("w", ""))
        mh = re.sub(r"[^\d]", "", me.get("h", ""))
        th = re.sub(r"[^\d]", "", te.get("h", ""))
        def size_match(a, b):
            if a == b:
                return True
            # 一方が他方のプレフィックスの場合も一致とみなす（例: "07" == "077"）
            n = min(len(a), len(b))
            return a[:n] == b[:n]
        if mw and tw and not size_match(mw, tw):
            issues.append(f"W寸法: 木工事={me['w']} ／ 建具図面={te['w']}")
        if mh and th and not size_match(mh, th):
            issues.append(f"H寸法: 木工事={me['h']} ／ 建具図面={te['h']}")

        # ★修正: 種類はカテゴリが一致すれば合格
        if me.get("type") and te.get("type"):
            mc = get_door_category(me["type"])
            tc = get_door_category(te["type"])
            # 折戸 ↔ 片引戸 は木工事図面でOCR誤読が起きやすいため互換扱い
            COMPAT = {("折戸", "片引戸"), ("片引戸", "折戸")}
            type_ok = (
                mc is not None and mc == tc          # 両方カテゴリ一致
                or (mc, tc) in COMPAT               # 互換カテゴリ
                or me["type"] in te["type"]          # 部分文字列一致
                or te["type"] in me["type"]
            )
            if not type_ok:
                issues.append(f"種類: 木工事={me['type']} ／ 建具図面={te['type']}")

        te_floor = te.get("floor", "")
        if issues:
            errors.append(f"**{me['raw']}{floor_label} / {norm}[{te_floor}]** — 不一致\n" +
                          "\n".join(f"  - {x}" for x in issues))
        else:
            detail = []
            if te.get("type"): detail.append(te["type"])
            if te.get("w"):    detail.append(f"W={te['w']}")
            if te.get("h"):    detail.append(f"H={te['h']}")
            ok.append(f"{me['raw']}{floor_label} / {norm}[{te_floor}] ✓ （{'  '.join(detail)}）")

    return errors, ok


def check_required_fields(tategu_entries):
    errors, ok = [], []
    for key, e in tategu_entries.items():
        if not e.get("type") and not e.get("w") and not e.get("h"):
            continue
        missing = []
        # 把手デザイン・敷居有無は斜線（未記入）が正常な場合があるため必須チェックから除外
        if not e.get("hinban"):  missing.append("品番")
        label = f"{e['raw']}（{e.get('room','?')}）[{e.get('floor','')}]"
        if missing:
            errors.append(f"**{label}** — 必須項目が空欄: {', '.join(missing)}")
        else:
            ok.append(f"{label} — 必須項目すべて記載あり")
    return errors, ok


def check_sill(tategu_entries):
    errors, ok = [], []
    for key, e in tategu_entries.items():
        if not e.get("type") and not e.get("w") and not e.get("h"):
            continue
        sill = e.get("sill", "")
        sill_color = e.get("sill_color", "")
        label = f"{e['raw']}（{e.get('room','?')}）[{e.get('floor','')}]"
        if "有" in sill:
            if not sill_color or sill_color in ("-", "－", "ー"):
                errors.append(
                    f"**{label}** — 敷居有なのに敷居色が未記入\n"
                    f"  - 敷居有無=「有」／ 敷居色／種類=「{sill_color or '（空欄）'}」"
                )
            else:
                ok.append(f"{label} — 敷居有・色={sill_color} ✓")
        elif "無" in sill or sill == "":
            ok.append(f"{label} — 敷居無/未記入")
    return errors, ok


def check_na(texts):
    combined = "\n".join(texts)
    pat = re.compile(r"#N/A|#REF!|#VALUE!|#DIV/0!|#NAME\?|#NULL!", re.I)
    found = list(set(pat.findall(combined)))
    if found:
        counts = {v: combined.count(v) for v in found}
        detail = ", ".join(f"{k}（{v}箇所）" for k, v in counts.items())
        return [f"**数式エラー値が残っています**: {detail}\n"
                "  - ExcelでエラーをなくしてからPDF出力し直してください"]
    return []


def check_taste(full_text, ref_text, tategu_entries):
    errors, ok = [], []
    taste_m = re.search(r"テイスト\s*[：:]\s*([A-Za-z])|テイスト([A-Za-z])\b", full_text)
    taste = (taste_m.group(1) or taste_m.group(2)).upper() if taste_m else None

    if not taste:
        errors.append(
            "**テイスト整合性チェック**\n"
            "  - テイスト指定（A/B/C等）を図面から読み取れませんでした\n"
            "  - 理由: 目視確認が必要"
        )
        return errors, ok

    if not ref_text:
        errors.append(
            f"**テイスト整合性チェック（テイスト={taste}）**\n"
            "  - 参照資料（テイスト規定PDF）が未アップロードのため照合不可\n"
            "  - 理由: 参照資料をアップロードするか、目視確認が必要"
        )
        return errors, ok

    rules = {}
    for line in ref_text.splitlines():
        if f"テイスト{taste}" in line or f"Taste{taste}" in line:
            dm = re.search(r"建具カラー[：:]\s*(\S+)", line)
            fm = re.search(r"床カラー[：:]\s*(\S+)", line)
            if dm: rules["door_color"] = dm.group(1)
            if fm: rules["floor_color"] = fm.group(1)

    if not rules:
        errors.append(
            f"**テイスト整合性チェック（テイスト={taste}）**\n"
            "  - 参照資料からテイスト{taste}の規定色を読み取れませんでした\n"
            "  - 理由: 目視確認が必要（AI判定困難）"
        )
        return errors, ok

    for key, e in tategu_entries.items():
        iss = []
        if "door_color" in rules and e.get("door_color") and e["door_color"] != rules["door_color"]:
            iss.append(f"建具カラー: 規定={rules['door_color']} ／ 記載={e['door_color']}")
        if "floor_color" in rules and e.get("floor_color") and e["floor_color"] != rules["floor_color"]:
            iss.append(f"床カラー: 規定={rules['floor_color']} ／ 記載={e['floor_color']}")
        label = f"{e['raw']}（{e.get('room','?')}）"
        if iss:
            errors.append(f"**{label}** — テイスト{taste}と不一致\n" +
                          "\n".join(f"  - {x}" for x in iss))
        else:
            ok.append(f"{label} — テイスト{taste}規定と一致 ✓")

    return errors, ok


def check_closet_sill(tategu_entries):
    errors, ok = [], []
    closet_kws = ["クローゼット", "折戸", "フラットレール", "ノンレール", "両開き", "観音"]

    for key, e in tategu_entries.items():
        typ  = e.get("type", "")
        sill = e.get("sill", "")
        sc   = e.get("sill_color", "")
        label = f"{e['raw']}（{e.get('room','?')}）[{e.get('floor','')}]"

        if not any(kw in typ for kw in closet_kws):
            continue

        # ノンレールタイプは敷居不要（斜線が正常）
        if "ノンレール" in typ:
            ok.append(f"{label} — ノンレール折戸・敷居不要 ✓")
            continue

        if "両開き" in typ or "観音" in typ:
            if sill != "有":
                errors.append(
                    f"**{label}** — 両開き/観音開き収納の敷居\n"
                    f"  - ルール: 両開き吊り収納は必ず敷居「有」が必要\n"
                    f"  - 現状: 敷居={sill or '（未記入）'}"
                )
            else:
                ok.append(f"{label} — 両開き収納・敷居有 ✓")
            continue

        if sill != "有":
            errors.append(
                f"**{label}** — クローゼット/折戸の敷居\n"
                f"  - ルール: フラットレール（洋室）は敷居「有」シャイングレー指示が必要\n"
                f"           和室ツバ無薄下枠は敷居「有」床材色指示が必要\n"
                f"  - 現状: 敷居={sill or '（未記入）'}"
            )
        elif not sc:
            errors.append(
                f"**{label}** — クローゼット/折戸の敷居色未記入\n"
                f"  - ルール: 洋室=シャイングレー、和室=床材色\n"
                f"  - 現状: 敷居「有」だが色の指示なし"
            )
        else:
            ok.append(f"{label} — クローゼット敷居有・色={sc} ✓")

    return errors, ok


def check_floor_thickness(spec):
    errors, ok = [], []
    mat = spec.get("floor_material", "")

    if not mat:
        return errors, ok

    if "フローリング" in mat and "無垢" not in mat:
        if "12" in mat:
            ok.append("床材厚み: フローリング12mm ✓")
        else:
            errors.append(
                "**床材厚み（フローリング）**\n"
                "  - ルール: 突板/複合フローリングは12mm\n"
                f"  - 現状: {mat[:50]}\n"
                "  - 理由: 12mmの記載が確認できません。目視確認が必要"
            )

    if "無垢" in mat:
        if "15" in mat:
            ok.append("床材厚み: 無垢フローリング15mm ✓")
        else:
            errors.append(
                "**床材厚み（無垢材）**\n"
                "  - ルール: 無垢材フローリングは15mm\n"
                f"  - 現状: {mat[:50]}\n"
                "  - 理由: 15mmの記載が確認できません。目視確認が必要"
            )

    if "フローリング" in mat and "無垢" not in mat:
        if "2P" in mat:
            ok.append("床材: 突板フローリング2P指示あり ✓")
        else:
            errors.append(
                "**突板フローリングの指示**\n"
                "  - ルール: 突板（複合）フローリングは2Pで提案・指示すること\n"
                f"  - 現状: {mat[:50]}\n"
                "  - 理由: 2Pの記載が確認できません。目視確認が必要"
            )

    return errors, ok


def check_visual_items():
    return [
        "**建具配置（赤丸数字）とリストの照合**\n"
        "  - 平面図の出入り口の赤丸数字①②③…とWD1/WD2/WD3…の対応を手動確認\n"
        "  - 目視確認が必要（AI判定困難）",

        "**床の張り方向（赤い矢印）確認**\n"
        "  - 各部屋・廊下・収納ごとに矢印が漏れなく配置されているか確認\n"
        "  - ルール: 基本は長手方向 ／ リビングと玄関ホールは方向を合わせる\n"
        "  - 扉で区切られた収納内も個別に指示があるか確認（吊り戸内・吊り戸下それぞれ）\n"
        "  - 目視確認が必要（AI判定困難）",

        "**エアコン干渉チェック**\n"
        "  - ①建具の開き方向とエアコン設置位置が干渉しないか\n"
        "    　特にH23収納建具・ハイドア開き戸は注意\n"
        "  - ②アウトセット建具引き込み側にエアコンがある場合、幅800mm確保できるか\n"
        "    　既製サイズで収まるか、特注サイズが必要か確認\n"
        "  - 目視確認が必要",

        "**収納扉とコンセント・スイッチの干渉確認**\n"
        "  - 枠なしすっきりタイプの収納扉は開扉時にコンセント・スイッチと干渉しないか\n"
        "  - 把手設置の場合は壁との干渉確認 → 涙目（ウレタンクッション）の必要性をお客様へ説明\n"
        "  - 目視確認が必要",

        "**見えないゾウストッパーの指示確認**\n"
        "  - リビングドア（開き戸）が玄関ホール側に開く場合 → 備考欄に指示があるか確認\n"
        "  - 片引き戸3枚建ての場合 → どちら側に壁をふかすかの記載があるか確認\n"
        "  - 目視確認が必要",

        "**格子・ミラーの備考記載確認**\n"
        "  - 格子を設置する場合: 備考欄に「格子あり」の記載があるか\n"
        "  - クローゼットドアにミラー付の場合: 備考欄に「ミラー付」の記載があるか\n"
        "  - 目視確認が必要",

        "**把手・引手の設置位置確認**\n"
        "  - H1400以下の建具: センター設置が可能\n"
        "  - H1400以上の建具: 下から900芯に引手\n"
        "  - 吊押入れ（畳天からH850）の場合: 建具H1580 → 引手がFL+H1750（高位置）になるため設置位置を確認\n"
        "  - 目視確認が必要",

        "**アウトセット・上吊り建具の敷居確認**\n"
        "  - 床の種類が変わる箇所のアウトセット・上吊り・観音開き建具は敷居の有無・色を記載すること\n"
        "  - 目視確認が必要（床種類の変化点の把握が必要）",

        "**メインフロアと収納・WICの床材統一確認**\n"
        "  - メインフロアと収納・WICが繋がっている場合は基本的に同一材種で統一\n"
        "  - 目視確認が必要",
    ]


def extract_property_name(text):
    for pat in [r"【.+?】", r"物件名[：:]\s*(.+)", r"現場名[：:]\s*(.+)"]:
        m = re.search(pat, text)
        if m:
            return (m.group(0) if m.lastindex is None else m.group(1)).strip()[:50]
    return "（図面から読み取り不可）"


# ─── レポート生成 ────────────────────────────────────────────────────────────

def build_report(property_name, all_errors, all_ok, tategu_entries, mokuko_entries):
    lines = ["### 📋 建具図面 照合チェックレポート",
             f"**対象物件:** {property_name}", ""]
    n_tategu = len(tategu_entries)
    n_mokuko = len(mokuko_entries)
    lines.append(f"> 建具・床図面: **{n_tategu}件** 抽出 ／ 木工事図面: **{n_mokuko}件** 抽出")
    lines.append("")

    lines.append("#### 🔴 エラー・要確認項目（図面間の不一致・不備）")
    if all_errors:
        for e in all_errors:
            lines.append(f"* {e}")
            lines.append("")
    else:
        lines.append("すべてルール通りです ✓")
    lines.append("")

    lines.append("#### 🟢 合格項目")
    if all_ok:
        for item in all_ok:
            lines.append(f"* {item}")
    else:
        lines.append("（照合できた合格項目なし）")
    lines.append("")
    lines.append("---")
    lines.append("修正が必要な箇所は以上です。")
    return "\n".join(lines)


def build_spec_summary(spec):
    rows = ["| 項目 | 内容 | 例外 |", "|---|---|---|"]
    data = [
        ("ドア色",       spec["door_color"],    spec["door_color_exc"]),
        ("ストッパー色", spec["stopper_color"],  ""),
        ("キャッチャー色", spec["catcher_color"], ""),
        ("枠色",        spec["frame_color"],    spec["frame_color_exc"]),
        ("ハンドル色",   spec["handle_color"],   ""),
        ("床材",         spec["floor_material"][:40] if spec["floor_material"] else "", ""),
        ("床材品番",     spec["floor_code"],     ""),
        ("床張り方向",   spec["floor_direction"][:40] if spec["floor_direction"] else "", ""),
    ]
    for label, val, exc in data:
        rows.append(f"| {label} | {val or '（未読取）'} | {exc} |")
    return "\n".join(rows)


def build_tategu_summary(tategu_entries):
    rows = []
    rows.append("| WD番号 | 階 | 部屋名 | 種類 | W | H | 把手 | 敷居 | 敷居色 | 品番 |")
    rows.append("|---|---|---|---|---|---|---|---|---|---|")
    def v(val, fallback="記載なし"):
        s = str(val).strip() if val else ""
        return s if s else fallback

    for key in sorted(tategu_entries.keys()):
        e = tategu_entries[key]
        rows.append(
            f"| {e['raw']} | {v(e.get('floor'))} | {v(e.get('room'))} "
            f"| {v(e.get('type'))} | {v(e.get('w'))} | {v(e.get('h'))} "
            f"| {v(e.get('handle'))} | {v(e.get('sill'))} "
            f"| {v(e.get('sill_color'))} | {v(e.get('hinban'))} |"
        )
    return "\n".join(rows)


# ─── Flask ルート ────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/preview-ocr", methods=["POST"])
def preview_ocr():
    """1階・2階の建具スケジュール画像をOCRしてプレビュー用データを返す"""
    api_key = request.form.get("api_key", "").strip() or _load_api_key()
    floor1_file = request.files.get("floor1")
    floor2_file = request.files.get("floor2")

    images = []
    if floor1_file and floor1_file.filename:
        images.append(floor1_file.read())
    if floor2_file and floor2_file.filename:
        images.append(floor2_file.read())

    if not images:
        return jsonify({"error": "画像を選択してください"}), 400
    if not api_key:
        return jsonify({"error": "APIキーが設定されていません。環境変数 ANTHROPIC_API_KEY を設定してください。"}), 400

    try:
        raw_text = ocr_images_with_claude(images, api_key)
        entries = parse_mokuko_from_text(raw_text)

        display_lines = []
        for key in sorted(entries.keys()):
            e = entries[key]
            wd  = e.get("raw", key)
            tp  = e.get("type", "")
            w   = f"W{e['w']}" if e.get("w") else ""
            h   = f"H{e['h']}" if e.get("h") else ""
            display_lines.append(f"{wd},{tp},{w},{h}")

        return jsonify({
            "raw": raw_text,
            "parsed_lines": display_lines,
            "count": len(entries)
        })
    except Exception as ex:
        import traceback
        return jsonify({"error": str(ex), "trace": traceback.format_exc()}), 500


@app.route("/check", methods=["POST"])
def check():
    open("debug_check.txt", "w", encoding="utf-8").write("check route called\n")
    mokuko_files  = request.files.getlist("mokuko")
    # 1階・2階個別画像モード
    floor1_file   = request.files.get("floor1")
    floor2_file   = request.files.get("floor2")
    floor_images  = [f for f in [floor1_file, floor2_file] if f and f.filename]
    tategu_file   = request.files.get("tategu")
    ref_file      = request.files.get("reference")
    mokuko_manual = request.form.get("mokuko_manual", "").strip()
    api_key = request.form.get("api_key", "").strip() or _load_api_key()

    if not tategu_file:
        return jsonify({"error": "建具・床図面をアップロードしてください。"}), 400
    if not floor_images and not mokuko_files and not mokuko_manual:
        return jsonify({"error": "木工事図面（1階または2階の画像）を選択してください。"}), 400

    try:
        valid_mokuko = [f for f in mokuko_files if f and f.filename]
        print(f"[check] uploaded mokuko files: {[f.filename for f in valid_mokuko]}")
        print(f"[check] floor images: {[f.filename for f in floor_images]}")
        tategu_bytes = tategu_file.read()
        ref_bytes    = ref_file.read() if ref_file and ref_file.filename else None

        def is_image_file(f):
            return f.filename.lower().rsplit(".", 1)[-1] in ("png", "jpg", "jpeg")

        # floor1/floor2 個別画像が優先、なければ従来の mokuko ファイルを使用
        if floor_images:
            image_files = floor_images
            pdf_files   = []
        else:
            image_files = [f for f in valid_mokuko if is_image_file(f)]
            pdf_files   = [f for f in valid_mokuko if not is_image_file(f)]

        tategu_pages = read_pdf_text(tategu_bytes)
        tategu_text  = "\n".join(p["text"] for p in tategu_pages)
        mokuko_text  = ""
        if pdf_files:
            for mf in pdf_files:
                mb = mf.read()
                pages = read_pdf_text(mb)
                mokuko_text += "\n".join(p["text"] for p in pages) + "\n"
        ref_text = ""
        if ref_bytes:
            ref_pages = read_pdf_text(ref_bytes)
            ref_text  = "\n".join(p["text"] for p in ref_pages)

        full_text = mokuko_text + "\n" + tategu_text

        tategu_entries = parse_tategu_pdf(tategu_bytes)

        if mokuko_manual:
            mokuko_entries = parse_mokuko_manual(mokuko_manual)
            ocr_used, ocr_err, claude_raw = False, None, None

        elif image_files:
            try:
                image_bytes_list = [f.read() for f in image_files]
                print(f"[check/image] Sending {len(image_bytes_list)} image(s) to Claude API")
                claude_raw = ocr_images_with_claude(image_bytes_list, api_key or None) if api_key else None
                if claude_raw:
                    print(f"[check/image] Claude returned {len(claude_raw)} chars")
                    print(f"[check/image] Preview:\n{claude_raw[:800]}")
                    mokuko_entries = parse_mokuko_from_text(claude_raw)
                    print(f"[check/image] Parsed entries: {list(mokuko_entries.keys())}")
                    ocr_used = True
                    ocr_err = None
                else:
                    mokuko_entries = {}
                    ocr_used = False
                    ocr_err = "APIキーが設定されていません。環境変数 ANTHROPIC_API_KEY を設定してください。"
                    claude_raw = None
            except Exception as e:
                mokuko_entries = {}
                ocr_used = False
                ocr_err = f"画像読み取りエラー: {e}"
                claude_raw = None
                print(f"[check/image] Error: {e}")

        elif pdf_files:
            pdf_files[0].seek(0)
            mokuko_bytes_first = pdf_files[0].read()
            mokuko_entries, ocr_used, ocr_err, claude_raw = parse_mokuko_pdf(mokuko_bytes_first, api_key or None)

        else:
            mokuko_entries, ocr_used, ocr_err, claude_raw = {}, False, None, None

        property_name = extract_property_name(full_text)

        all_errors, all_ok = [], []

        if ocr_err:
            all_errors.append(f"**木工事図面の読み取りエラー**\n  - {ocr_err}\n  - 木工事図面の照合は目視確認が必要です")

        errs, oks = check_cross_reference(mokuko_entries, tategu_entries)
        all_errors.extend(errs); all_ok.extend(oks)

        errs, oks = check_required_fields(tategu_entries)
        all_errors.extend(errs); all_ok.extend(oks)

        errs, oks = check_sill(tategu_entries)
        all_errors.extend(errs); all_ok.extend(oks)

        errs = check_na([mokuko_text, tategu_text])
        all_errors.extend(errs)

        errs, oks = check_taste(full_text, ref_text, tategu_entries)
        all_errors.extend(errs); all_ok.extend(oks)

        bldg_spec = parse_building_spec(tategu_text)
        errs, oks = check_building_spec(bldg_spec)
        all_errors.extend(errs)
        all_ok.extend(oks)

        errs, oks = check_closet_sill(tategu_entries)
        all_errors.extend(errs); all_ok.extend(oks)

        errs, oks = check_floor_thickness(bldg_spec)
        all_errors.extend(errs); all_ok.extend(oks)

        all_errors.extend(check_visual_items())

        report  = build_report(property_name, all_errors, all_ok, tategu_entries, mokuko_entries)
        summary = build_tategu_summary(tategu_entries)
        spec_md = build_spec_summary(bldg_spec)

        return jsonify({
            "report": report,
            "summary": summary,
            "spec_md": spec_md,
            "tategu_count": len(tategu_entries),
            "mokuko_count": len(mokuko_entries),
            "ocr_used": ocr_used,
            "mokuko_empty": len(mokuko_entries) == 0,
            "claude_raw": claude_raw,
        })

    except Exception as e:
        import traceback
        return jsonify({"error": f"処理エラー: {str(e)}", "trace": traceback.format_exc()}), 500


@app.route("/debug-mokuko", methods=["POST"])
def debug_mokuko():
    mokuko_file = request.files.get("mokuko")
    api_key = request.form.get("api_key", "").strip() or _load_api_key() or None

    if not mokuko_file:
        return jsonify({"error": "木工事図面ファイルが必要です"}), 400

    try:
        pdf_bytes = mokuko_file.read()

        pages = read_pdf_text(pdf_bytes)
        raw_text = "\n".join(p["text"] for p in pages)
        is_cid = is_cid_font_pdf(raw_text)

        result = {
            "step1_raw_text_sample": raw_text[:500],
            "step1_is_cid_font": is_cid,
            "step2_api_key_received": bool(api_key),
            "step3_claude_raw": None,
            "step4_parsed_entries": {},
            "step4_entry_count": 0,
            "error": None,
        }

        if is_cid and api_key:
            try:
                claude_text = ocr_with_claude(pdf_bytes, api_key)
                result["step3_claude_raw"] = claude_text[:2000]
                entries = parse_mokuko_from_text(claude_text)
                result["step4_parsed_entries"] = {k: v for k, v in list(entries.items())[:10]}
                result["step4_entry_count"] = len(entries)
            except Exception as e:
                result["error"] = f"Claude APIエラー: {e}"
        elif not is_cid:
            entries = parse_mokuko_from_text(raw_text)
            result["step4_parsed_entries"] = {k: v for k, v in list(entries.items())[:10]}
            result["step4_entry_count"] = len(entries)
            result["note"] = "CIDフォントではないためテキスト直接解析"

        return jsonify(result)

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"建具図面チェッカー 起動中... http://localhost:{port}")
    app.run(debug=False, host="0.0.0.0", port=port)
