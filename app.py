import streamlit as st
import re
import io
import os
import base64
import pdfplumber
import fitz  # PyMuPDF

# ============================================================
# PDF テキスト抽出
# ============================================================
def read_pdf_text(pdf_bytes):
    pages = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                pages.append({"page": i + 1, "text": text})
    except Exception as e:
        pages.append({"page": 1, "text": "", "error": str(e)})
    return pages


# ============================================================
# 丸囲み数字ユーティリティ
# ============================================================
CIRCLE_CHARS = "".join(chr(c) for c in
    list(range(0x2460, 0x2474)) +
    list(range(0x3251, 0x3260)) +
    list(range(0x32B1, 0x32C0))
)
WD_CIRCLE_PAT = re.compile(r"WD([" + CIRCLE_CHARS + r"])")
WD_DIGIT_PAT  = re.compile(r"WD(\d{1,3})\b")


def circle_to_int(ch):
    c = ord(ch)
    if 0x2460 <= c <= 0x2473: return c - 0x245F
    if 0x3251 <= c <= 0x325F: return c - 0x3250 + 21
    if 0x32B1 <= c <= 0x32BF: return c - 0x32B0 + 37
    return None


def normalize_mokuko_wd(raw):
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


# ============================================================
# 建具・設計図面パーサー
# ============================================================
def split_by_keyword(line, keyword):
    parts = re.split(re.escape(keyword) + r"\s*", line)
    return [p.strip() for p in parts[1:] if p.strip()]


def find_floor_label(text):
    m = re.search(r"(\d)階平面図", text)
    return f"{m.group(1)}F" if m else None


def parse_tategu_pdf(pdf_bytes):
    entries = {}
    pages = read_pdf_text(pdf_bytes)

    for page_data in pages:
        pnum  = page_data["page"]
        text  = page_data["text"]
        lines = text.splitlines()
        floor = find_floor_label(text) or f"p{pnum}"

        i = 0
        while i < len(lines):
            line = lines[i]
            wd_matches = list(WD_CIRCLE_PAT.finditer(line))
            if not wd_matches:
                i += 1
                continue

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

            i += 1
            block = []
            while i < len(lines):
                if WD_CIRCLE_PAT.search(lines[i]):
                    break
                block.append(lines[i])
                i += 1

            n_wd = len(group)
            pending_w = {}

            for bi, bl in enumerate(block):
                if "種類" in bl and "敷居色" not in bl:
                    types = split_by_keyword(bl, "種類")
                    for ti, tv in enumerate(types):
                        if ti < n_wd and not group[ti]["type"] and tv and "敷居色" not in tv:
                            group[ti]["type"] = tv[:40]

                if "品番" in bl:
                    hinbans = split_by_keyword(bl, "品番")
                    for hi, hv in enumerate(hinbans):
                        if hi < n_wd and not group[hi]["hinban"]:
                            group[hi]["hinban"] = hv[:50]

                if "大きさ" in bl:
                    size_sections = re.split(r"大きさ\s*", bl)
                    for si, sp in enumerate(size_sections[1:]):
                        if si >= n_wd:
                            break
                        wm = re.search(r"W\s*(\d{3,4})", sp)
                        hm = re.search(r"H\s*(\d{3,4}|特注)", sp)
                        if wm:
                            group[si]["w"] = wm.group(1)
                        elif si in pending_w:
                            group[si]["w"] = pending_w[si]
                        if hm:
                            group[si]["h"] = hm.group(1)

                sw = re.match(r"^\s*(W\s*\d{3,4})\s*$", bl)
                if sw:
                    w_val = re.search(r"(\d{3,4})", sw.group(1)).group(1)
                    pending_w[n_wd - 1] = w_val

            sill_hdr_idx = None
            sill_col_hdr_idx = None
            handle_hdr_idx = None

            for bi, bl in enumerate(block):
                if "敷居有無" in bl and sill_hdr_idx is None:
                    sill_hdr_idx = bi
                if ("敷居色" in bl or "敷居色・種類" in bl) and sill_col_hdr_idx is None:
                    sill_col_hdr_idx = bi
                if "把手デザイン" in bl and handle_hdr_idx is None:
                    handle_hdr_idx = bi

            if sill_hdr_idx is not None:
                for bi in range(sill_hdr_idx + 1, min(sill_hdr_idx + 4, len(block))):
                    val = block[bi].strip()
                    if val and "敷居" not in val and "種類" not in val:
                        vals = re.findall(r"有|無", val)
                        for si, sv in enumerate(vals):
                            if si < n_wd:
                                group[si]["sill"] = sv
                        break

            if sill_col_hdr_idx is not None:
                for bi in range(sill_col_hdr_idx + 1, min(sill_col_hdr_idx + 4, len(block))):
                    val = block[bi].strip()
                    if val and "敷居" not in val:
                        pl_colors = re.findall(r"PL\s+\S+", val)
                        for ci, pc in enumerate(pl_colors):
                            if ci < n_wd:
                                group[ci]["sill_color"] = pc
                        if not pl_colors:
                            color_cands = re.findall(r"[^\s]{2,20}", val)
                            for ci, cc in enumerate(color_cands[:n_wd]):
                                group[ci]["sill_color"] = cc
                        break

            if handle_hdr_idx is not None:
                for bi in range(handle_hdr_idx + 1, min(handle_hdr_idx + 4, len(block))):
                    val = block[bi].strip()
                    if val and "把手" not in val and "敷居" not in val:
                        handles = re.findall(r"[" + CIRCLE_CHARS + r"][^\s" + CIRCLE_CHARS + r"]*", val)
                        for hi, hv in enumerate(handles):
                            if hi < n_wd and not group[hi]["handle"]:
                                group[hi]["handle"] = hv
                        if not handles and val:
                            group[0]["handle"] = val[:30]
                        break

            for wd in group:
                entries[wd["full_key"]] = wd

    return entries


# ============================================================
# CIDフォント判定 & OCR
# ============================================================
def is_cid_font_pdf(text):
    if not text:
        return True
    cid_count = text.count("(cid:")
    total     = max(len(text), 1)
    readable = sum(1 for c in text if c.isprintable() and ord(c) > 31)
    return cid_count > 5 or readable / total < 0.3


def pdf_to_images(pdf_bytes, dpi=200):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    zoom = dpi / 72
    mat  = fitz.Matrix(zoom, zoom)
    for page in doc:
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
        images.append(pix.tobytes("png"))
    return images


def ocr_images_with_claude(image_list, api_key):
    try:
        import anthropic
    except ImportError:
        raise RuntimeError("anthropic パッケージが見つかりません。`pip install anthropic` を実行してください。")

    client = anthropic.Anthropic(api_key=api_key)
    prompt_text = (
        "この画像は木工事図面（建具スケジュール表）です。\n"
        "以下の手順で、表に記載されたWD番号・種類・W寸法・H寸法を正確に読み取ってください。\n\n"
        "【読み取り手順】\n"
        "手順1: 表の一番上の行にあるWD番号（WD101, WD102, WD201など）を左から右へ順番に確認する\n"
        "手順2: 各WD番号の「列」を画像上で厳密に追う。その列の中にある種類・W寸法・H寸法を読む\n"
        "手順3: 隣の列と混同しないようWD番号と寸法が同じ列に属していることを確認してから出力する\n\n"
        "【出力形式】1つのWD番号につき1行、カンマ区切りで出力（実際の値を画像から読むこと）:\n"
        "WDxxx,種類名,W寸法,H寸法\n\n"
        "【厳守ルール】\n"
        "1. カンマ区切り形式のみ（スペース区切り不可）\n"
        "2. 1つのWD番号につき必ず1行のみ\n"
        "3. WD番号は「WD」+3桁数字形式（例: WD101, WD201）\n"
        "4. 種類は表記通り\n"
        "5. W寸法・H寸法は必ず「そのWD番号と同じ列」の値を読む\n"
        "6. 読み取れない項目は空欄（例: WD103,折戸,,）\n"
        "7. WDスケジュール表以外（金物・サッシ表・ヘッダー等）は出力しない\n"
        "8. 解説・説明文は一切書かない。データ行のみ出力する\n"
        "9. 「WD」が付いていない数字（101, 102等）はサッシ番号なので無視する\n"
        "10. WD1xx（1階）とWD2xx（2階）のみ対象。WD3xx以上は出力しない\n"
    )

    all_text = []
    for img_bytes in image_list:
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
    images = pdf_to_images(pdf_bytes, dpi=200)
    return ocr_images_with_claude(images, api_key)


# ============================================================
# 木工事図面パーサー
# ============================================================
def parse_mokuko_from_text(text):
    entries = {}
    lines = text.splitlines()

    type_kws = ["片開きドア", "開き戸", "片引戸", "引戸", "引き戸",
                "折戸", "クローゼット", "上吊", "アウトセット", "ドア", "サッシ"]

    def extract_type_only(seg):
        seg = re.sub(r"W\s*\d{3,4}", "", seg)
        seg = re.sub(r"H\s*(\d{3,4}|特注)", "", seg)
        seg = seg.strip()
        seg = re.sub(r"^W([^\d])", r"\1", seg)
        return seg.strip()

    csv_found = False
    for line in lines:
        m = re.match(r"(WD\d{3})\s*,\s*([^,]*),\s*(W[\d]+|)\s*,?\s*(H[\d]+|H特注|)", line.strip())
        if not m:
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

    i = 0
    while i < len(lines):
        line = lines[i]
        wds = list(re.finditer(r"WD(\d{3})", line))
        if not wds:
            i += 1
            continue

        group = []
        for m_wd in wds:
            raw  = m_wd.group(0)
            norm, floor = normalize_mokuko_wd(raw)
            group.append({"raw": raw, "key": norm, "floor": floor,
                          "type": "", "w": "", "h": ""})

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


def parse_mokuko_manual(text):
    entries = {}
    if not text or not text.strip():
        return entries

    lines = [l.strip() for l in text.strip().splitlines() if l.strip()]
    type_kws = ["片開きドア", "開き戸", "片引戸", "引き戸", "引戸", "折戸",
                "クローゼット", "上吊", "アウトセット", "ドア", "サッシ", "引"]

    for line in lines:
        m = re.search(r"WD(\d{1,3})", line)
        if not m:
            continue
        raw = f"WD{m.group(1)}"
        norm, floor = normalize_mokuko_wd(raw)
        typ = ""
        for kw in type_kws:
            if kw in line:
                typ = kw; break
        w_m = re.search(r"W\s*(\d{3,4})", line)
        h_m = re.search(r"H\s*(\d{3,4}|特注)", line)
        nums = re.findall(r"\b(\d{3,4})\b", line)
        w = w_m.group(1) if w_m else (nums[0] if len(nums) > 0 else "")
        h = h_m.group(1) if h_m else (nums[1] if len(nums) > 1 else "")
        key = f"{norm}_{floor}" if floor else norm
        entries[key] = {"raw": raw, "key": norm, "floor": floor,
                        "type": typ, "w": w, "h": h}
    return entries


def parse_mokuko_pdf(pdf_bytes, api_key=None):
    pages = read_pdf_text(pdf_bytes)
    full_text = "\n".join(p["text"] for p in pages)
    ocr_used = False
    entries = parse_mokuko_from_text(full_text)
    claude_raw = None
    cid_detected = is_cid_font_pdf(full_text)

    if (len(entries) == 0 or cid_detected) and api_key:
        try:
            claude_text = ocr_with_claude(pdf_bytes, api_key)
            claude_raw = claude_text
            full_text = claude_text
            ocr_used = True
            entries = parse_mokuko_from_text(full_text)
        except Exception as e:
            if len(entries) == 0:
                return {}, False, f"Claude APIエラー: {e}", None

    if len(entries) == 0 and not api_key:
        return {}, False, "木工事図面からWDデータを読み取れませんでした。", None

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


# ============================================================
# 建具仕様・建物仕様パーサー
# ============================================================
def parse_building_spec(text):
    spec = {
        "door_color": "", "stopper_color": "", "catcher_color": "",
        "frame_color": "", "handle_color": "",
        "door_color_exc": "", "frame_color_exc": "",
        "floor_material": "", "floor_code": "", "floor_direction": "",
    }
    lines = text.splitlines()
    in_spec = False
    for line in lines:
        if "建具仕様" in line:
            in_spec = True
        if not in_spec:
            continue
        if ("ドア色" in line or ("ドア" in line and "色" in line)) and not spec["door_color"]:
            m = re.search(r"ドア色\s+([^\s※①]+)", line)
            if m:
                spec["door_color"] = m.group(1).strip()
        if "ストッパー" in line and not spec["stopper_color"]:
            m = re.search(r"(?:ストッパー\s+色|ストッパー色|ストッパー)\s+(\S+)", line)
            if m:
                spec["stopper_color"] = m.group(1).strip()
        if "キャッチャー" in line and not spec["catcher_color"]:
            m = re.search(r"(?:キャッチャー\s+色|キャッチャー色|キャッチャー)\s+(\S+)", line)
            if m:
                spec["catcher_color"] = m.group(1).strip()
        if "枠色" in line and "枠色は建具仕様" not in line and not spec["frame_color"]:
            m = re.search(r"枠色\s+([^\s※①]+)", line)
            if m:
                spec["frame_color"] = m.group(1).strip()
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
        if "水回" in line and in_spec and spec["door_color"]:
            break
    return spec


# ============================================================
# チェック関数群
# ============================================================
DOOR_TYPE_CATEGORIES = {
    "片開きドア": ["片開きドア", "片開き戸", "開き戸", "トイレドア", "洗面ドア",
                   "⑥トイレドア", "③洗面ドア", "⑥", "③"],
    "片引戸": ["片引戸", "片引き戸", "アウトセット引き戸", "アウトセット引戸", "アウトセット",
               "上吊引き戸", "上吊引戸", "上吊", "⑫片引き戸", "⑨アウトセット", "⑫", "⑨"],
    "折戸": ["折戸", "クローゼット", "ノンレール", "⑭クローゼット", "クローゼットドア", "⑭"],
}


def get_door_category(type_str):
    if not type_str:
        return None
    for category, keywords in DOOR_TYPE_CATEGORIES.items():
        for kw in keywords:
            if kw in type_str:
                return category
    return None


def floors_match(f1, f2):
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


def check_cross_reference(mokuko_entries, tategu_entries):
    errors, ok = [], []
    if not mokuko_entries:
        errors.append(
            "**【木工事図面 ↔ 建具図面】寸法・種類の照合**\n"
            "  - 現状: 木工事図面からWDデータを読み取れませんでした\n"
            "  - 理由: 目視確認が必要"
        )
        return errors, ok

    for mk_key, me in mokuko_entries.items():
        norm  = me["key"]
        floor = me.get("floor")
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

        issues = []
        mw = re.sub(r"[^\d]", "", me.get("w", ""))
        tw = re.sub(r"[^\d]", "", te.get("w", ""))
        mh = re.sub(r"[^\d]", "", me.get("h", ""))
        th = re.sub(r"[^\d]", "", te.get("h", ""))
        if mw and tw and mw != tw:
            issues.append(f"W寸法: 木工事={me['w']} ≠ 建具図面={te['w']}")
        if mh and th and mh != th:
            issues.append(f"H寸法: 木工事={me['h']} ≠ 建具図面={te['h']}")

        if me.get("type") and te.get("type"):
            mc = get_door_category(me["type"])
            tc = get_door_category(te["type"])
            type_ok = (
                mc is not None and mc == tc
                or me["type"] in te["type"]
                or te["type"] in me["type"]
            )
            if not type_ok:
                issues.append(f"種類: 木工事={me['type']} ≠ 建具図面={te['type']}")

        te_floor = te.get("floor", "")
        if issues:
            errors.append(f"**{me['raw']}{floor_label} / {norm}[{te_floor}]** → 不一致\n" +
                          "\n".join(f"  - {x}" for x in issues))
        else:
            detail = []
            if te.get("type"): detail.append(te["type"])
            if te.get("w"):    detail.append(f"W={te['w']}")
            if te.get("h"):    detail.append(f"H={te['h']}")
            ok.append(f"{me['raw']}{floor_label} / {norm}[{te_floor}] ✅（{'  '.join(detail)}）")

    return errors, ok


def check_required_fields(tategu_entries):
    errors, ok = [], []
    for key, e in tategu_entries.items():
        if not e.get("room") and not e.get("type") and not e.get("w") and not e.get("h") and not e.get("hinban"):
            continue
        missing = []
        if not e.get("handle"):  missing.append("把手デザイン")
        if not e.get("sill"):    missing.append("敷居有無")
        if not e.get("hinban"):  missing.append("品番")
        label = f"{e['raw']}（{e.get('room','?')}）[{e.get('floor','')}]"
        if missing:
            errors.append(f"**{label}** → 必須項目が空欄: {', '.join(missing)}")
        else:
            ok.append(f"{label} → 必須項目すべて記載あり")
    return errors, ok


def check_sill(tategu_entries):
    errors, ok = [], []
    for key, e in tategu_entries.items():
        if not e.get("room") and not e.get("type") and not e.get("w") and not e.get("h") and not e.get("hinban"):
            continue
        sill = e.get("sill", "")
        sill_color = e.get("sill_color", "")
        label = f"{e['raw']}（{e.get('room','?')}）[{e.get('floor','')}]"
        if "有" in sill:
            if not sill_color or sill_color in ("-", "ー", "ー"):
                errors.append(
                    f"**{label}** → 敷居有なのに敷居色が未記載\n"
                    f"  - 敷居有無=「有」、敷居色・種類={sill_color or '（空欄）'}"
                )
            else:
                ok.append(f"{label} → 敷居有・色={sill_color} ✅")
        elif "無" in sill or sill == "":
            ok.append(f"{label} → 敷居無/未記載")
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
        if "両開き" in typ or "観音" in typ:
            if sill != "有":
                errors.append(
                    f"**{label}** → 両開き/観音開き収納の敷居\n"
                    f"  - ルール: 両開き吊り収納は必ず敷居「有」が必要\n"
                    f"  - 現状: 敷居={sill or '（未記載）'}"
                )
            else:
                ok.append(f"{label} → 両開き収納の敷居有 ✅")
            continue
        if sill != "有":
            errors.append(
                f"**{label}** → クローゼット/折戸の敷居\n"
                f"  - ルール: フラットレール（玄関）は敷居「有」シャイングレー指定が必要\n"
                f"           水回りデバ無し上框は敷居「有」建具色指定が必要\n"
                f"  - 現状: 敷居={sill or '（未記載）'}"
            )
        elif not sc:
            errors.append(
                f"**{label}** → クローゼット/折戸の敷居色未記載\n"
                f"  - ルール: 玄関=シャイングレー、居室=建具色\n"
                f"  - 現状: 敷居「有」だが色の指定なし"
            )
        else:
            ok.append(f"{label} → クローゼット敷居有・色={sc} ✅")
    return errors, ok


def check_visual_items():
    return [
        "**建具配置（丸囲み数字）とリストの照合**\n"
        "  - 平面図の出入り口等の丸囲み数字①②③…とWD1/WD2/WD3…の対応を手動確認\n"
        "  - 目視確認が必要（AI判定困難）",

        "**床の張り方向（張り・貼り）確認**\n"
        "  - 各部屋の建具の収納ごとに矢印が入れなく記載されているか確認\n"
        "  - ルール: 基本は長手方向。リビングと廊下ホールは方向を合わせる\n"
        "  - 手で区切られた収納内も個別に指定があるか確認（吊り戸内の吊り戸下そろえぞろい）\n"
        "  - 目視確認が必要（AI判定困難）",

        "**エアコン干渉チェック**\n"
        "  - ①建具の開き方向とエアコン設置位置が干渉しないか\n"
        "  - ②アウトセット建具引き込み側にエアコンがある場合、壁500mm確保できるか\n"
        "  - 目視確認が必要",

        "**収納扉とコンセント・スイッチの干渉確認**\n"
        "  - 枠なしすっきりタイプの収納扉は開扉時にコンセント・スイッチと干渉しないか\n"
        "  - 目視確認が必要",

        "**見えないガウストッパーの指定確認**\n"
        "  - リビングドア（開き戸）が廊下ホール側に開く場合 → 金物欄に指定があるか確認\n"
        "  - 目視確認が必要",

        "**鏡子・ミラーの金物記載確認**\n"
        "  - 鏡子を設定する場合、金物欄に「鏡子あり」の記載があるか\n"
        "  - クローゼットドアにミラー付の場合、金物欄に「ミラー付」の記載があるか\n"
        "  - 目視確認が必要",

        "**把手・引手の設置位置確認**\n"
        "  - H1400以下の建具: センター設定が可能\n"
        "  - H1400以上の建具: 下から900位に引手\n"
        "  - 目視確認が必要",

        "**アウトセット・上吊り建具の敷居確認**\n"
        "  - 床の種類が変わる箇所のアウトセット・上吊りの観音開き建具は敷居の有無・色を記載すること\n"
        "  - 目視確認が必要",

        "**メインフロアと収納・WICの床材統一確認**\n"
        "  - メインフロアと収納・WICが繋がっている場合は基本的に同一床材で統一\n"
        "  - 目視確認が必要",
    ]


def extract_property_name(text):
    for pat in [r"【.+?】", r"物件名[：:]\s*(.+)", r"現場名[：:]\s*(.+)"]:
        m = re.search(pat, text)
        if m:
            return (m.group(0) if m.lastindex is None else m.group(1)).strip()[:50]
    return "（図面から読み取り不可）"


# ============================================================
# レポート生成
# ============================================================
def build_report(property_name, all_errors, all_ok, tategu_entries, mokuko_entries):
    lines = ["### 📋 建具図面 照合チェックレポート",
             f"**対象物件:** {property_name}", ""]
    n_tategu = len(tategu_entries)
    n_mokuko = len(mokuko_entries)
    lines.append(f"> 建具・設計図面: **{n_tategu}件** 抽出 ／ 木工事図面: **{n_mokuko}件** 抽出")
    lines.append("")

    lines.append("#### 🔴 エラー・要確認項目（図面間の不一致・不足）")
    if all_errors:
        for e in all_errors:
            lines.append(f"* {e}")
            lines.append("")
    else:
        lines.append("すべてルール通りです ✅")
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


# ============================================================
# Streamlit UI
# ============================================================
st.set_page_config(page_title="建具図面チェッカー", page_icon="🏠", layout="wide")
st.title("🏠 建具図面チェッカー")
st.caption("AI自動照合システム ｜ 木工事図面 × 建具・設計図面の照合チェック")
st.divider()

# APIキー
api_key_file = os.path.join(os.path.dirname(__file__), "api_key.txt")
default_key = ""
if os.path.exists(api_key_file):
    with open(api_key_file, "r", encoding="utf-8") as f:
        default_key = f.read().strip()

with st.sidebar:
    st.markdown("""
    **チェック対象項目**
    - ✅ WD番号・種類・W寸法・H寸法の照合
    - ✅ 建具仕様（ドア色・枠色等）の読み取り
    - ✅ 必須項目（把手・敷居・品番）の記載チェック
    - ✅ 敷居有無・色の整合性チェック
    - ✅ クローゼット/折戸の敷居ルール
    - ✅ 数式エラー値（#N/A等）の検出
    - ✅ 目視確認項目リスト出力
    """)

    st.divider()
    api_key_input = st.text_input(
        "Anthropic APIキー（CIDフォントPDF読み取り用）",
        value=default_key,
        type="password",
        help="api_key.txtに記載があれば自動読み込みされます"
    )

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📂 木工事図面（建具スケジュール表）")

    mokuko_mode = st.radio(
        "入力方法",
        ["画像アップロード（1階・2階）", "手動入力"],
        horizontal=True
    )

    floor1_file = None
    floor2_file = None
    mokuko_manual_text = ""

    if mokuko_mode == "画像アップロード（1階・2階）":
        c1, c2 = st.columns(2)
        with c1:
            floor1_file = st.file_uploader("1階 建具スケジュール", type=["png", "jpg", "jpeg"], key="floor1")
        with c2:
            floor2_file = st.file_uploader("2階 建具スケジュール", type=["png", "jpg", "jpeg"], key="floor2")

        if floor1_file or floor2_file:
            if st.button("🔍 読み取り確認（OCRプレビュー）"):
                api_key = api_key_input or default_key
                if not api_key:
                    st.error("APIキーが設定されていません。サイドバーで入力してください。")
                else:
                    with st.spinner("画像をAIで読み取り中..."):
                        images = []
                        if floor1_file:
                            images.append(floor1_file.read())
                            floor1_file.seek(0)
                        if floor2_file:
                            images.append(floor2_file.read())
                            floor2_file.seek(0)
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
                            st.success(f"✅ {len(entries)}件のWDデータを読み取りました")
                            st.code("\n".join(display_lines))
                        except Exception as ex:
                            st.error(f"OCRエラー: {ex}")
    else:
        mokuko_manual_text = st.text_area(
            "WDデータを貼り付け",
            height=200,
            placeholder="WD101,片開きドア,W654,H2035\nWD102,片開きドア,W778,H2035"
        )

    st.divider()
    st.subheader("📂 建具・設計図面")
    tategu_file = st.file_uploader("建具・設計図面PDF", type=["pdf"], key="tategu")
    if tategu_file:
        st.success(f"✅ {tategu_file.name}")

    st.divider()
    st.subheader("📂 参照資料（任意）")
    ref_file = st.file_uploader("テイスト規定・チェック項目PDF", type=["pdf"], key="ref")

    st.divider()

    can_check = tategu_file and (floor1_file or floor2_file or mokuko_manual_text)
    if can_check:
        if st.button("🔍 チェック開始", type="primary", use_container_width=True):
            api_key = api_key_input or default_key

            tategu_bytes = tategu_file.read()
            ref_bytes = ref_file.read() if ref_file else None

            tategu_pages = read_pdf_text(tategu_bytes)
            tategu_text  = "\n".join(p["text"] for p in tategu_pages)

            mokuko_text = ""
            ref_text = ""
            if ref_bytes:
                ref_pages = read_pdf_text(ref_bytes)
                ref_text = "\n".join(p["text"] for p in ref_pages)

            full_text = mokuko_text + "\n" + tategu_text

            tategu_entries = parse_tategu_pdf(tategu_bytes)

            ocr_used = False
            ocr_err = None
            claude_raw = None

            if mokuko_manual_text:
                mokuko_entries = parse_mokuko_manual(mokuko_manual_text)
            elif floor1_file or floor2_file:
                with st.spinner("📄 画像をAIで読み取り中... しばらくお待ちください"):
                    try:
                        image_bytes_list = []
                        if floor1_file:
                            floor1_file.seek(0)
                            image_bytes_list.append(floor1_file.read())
                        if floor2_file:
                            floor2_file.seek(0)
                            image_bytes_list.append(floor2_file.read())

                        if api_key:
                            claude_raw = ocr_images_with_claude(image_bytes_list, api_key)
                            mokuko_entries = parse_mokuko_from_text(claude_raw)
                            ocr_used = True
                        else:
                            mokuko_entries = {}
                            ocr_err = "APIキーが設定されていません"
                    except Exception as e:
                        mokuko_entries = {}
                        ocr_err = f"画像読み取りエラー: {e}"
            else:
                mokuko_entries = {}

            property_name = extract_property_name(full_text)

            all_errors, all_ok = [], []

            if ocr_err:
                all_errors.append(f"**木工事図面の読み取りエラー**\n  - {ocr_err}")

            errs, oks = check_cross_reference(mokuko_entries, tategu_entries)
            all_errors.extend(errs); all_ok.extend(oks)

            errs, oks = check_required_fields(tategu_entries)
            all_errors.extend(errs); all_ok.extend(oks)

            errs, oks = check_sill(tategu_entries)
            all_errors.extend(errs); all_ok.extend(oks)

            errs = check_na([mokuko_text, tategu_text])
            all_errors.extend(errs)

            bldg_spec = parse_building_spec(tategu_text)
            errs_b, oks_b = [], []
            if not bldg_spec["door_color"]:
                errs_b.append("**建具仕様のドア色**\n  - 建具図面の「建具仕様」欄からドア色を読み取れませんでした\n  - 理由: 目視確認が必要")
            if not bldg_spec["floor_direction"]:
                errs_b.append("**床の張り方向記載**\n  - 「建物仕様」欄の床の張り方向記載を読み取れませんでした\n  - 理由: 目視確認が必要")
            if bldg_spec["door_color"]:
                oks_b.append(f"建具仕様を読み取りました（ドア色={bldg_spec['door_color']}、詳細は下部テーブル参照）")
            all_errors.extend(errs_b); all_ok.extend(oks_b)

            errs, oks = check_closet_sill(tategu_entries)
            all_errors.extend(errs); all_ok.extend(oks)

            all_errors.extend(check_visual_items())

            report = build_report(property_name, all_errors, all_ok, tategu_entries, mokuko_entries)

            st.session_state["tategu_report"] = report
            st.session_state["tategu_entries"] = tategu_entries
            st.session_state["mokuko_entries"] = mokuko_entries
            st.session_state["bldg_spec"] = bldg_spec
            st.session_state["ocr_used"] = ocr_used
            st.session_state["claude_raw"] = claude_raw

            if all_errors:
                st.warning(f"⚠️ {len(all_errors)}件のエラー・要確認項目があります")
            else:
                st.success("✅ すべてルール通りです")
    else:
        st.info("木工事図面と建具・設計図面をアップロードしてください")


with col2:
    st.subheader("📋 チェックレポート")

    if "tategu_report" in st.session_state:
        report = st.session_state["tategu_report"]

        if st.session_state.get("ocr_used"):
            st.info("🔍 木工事図面はOCRで読み取りました。認識精度に限りがあるため、照合結果は目視でもご確認ください。")

        st.markdown(report)

        # 建具仕様サマリー
        spec = st.session_state.get("bldg_spec", {})
        if spec and spec.get("door_color"):
            with st.expander("📊 建具仕様・建物仕様（自動抽出結果）", expanded=True):
                spec_data = {
                    "項目": ["ドア色", "ストッパー色", "キャッチャー色", "枠色", "ハンドル色", "床材", "床品番", "床張り方向"],
                    "内容": [
                        spec.get("door_color", "") or "（未読取）",
                        spec.get("stopper_color", "") or "（未読取）",
                        spec.get("catcher_color", "") or "（未読取）",
                        spec.get("frame_color", "") or "（未読取）",
                        spec.get("handle_color", "") or "（未読取）",
                        (spec.get("floor_material", "") or "（未読取）")[:40],
                        spec.get("floor_code", "") or "（未読取）",
                        (spec.get("floor_direction", "") or "（未読取）")[:40],
                    ]
                }
                st.table(spec_data)

        # 建具データ一覧
        tategu_entries = st.session_state.get("tategu_entries", {})
        if tategu_entries:
            with st.expander("📊 建具・設計図面 抽出データ一覧（照合確認用）"):
                rows = []
                for key in sorted(tategu_entries.keys()):
                    e = tategu_entries[key]
                    rows.append({
                        "WD番号": e["raw"],
                        "階": e.get("floor", ""),
                        "部屋名": e.get("room", ""),
                        "種類": e.get("type", ""),
                        "W": e.get("w", ""),
                        "H": e.get("h", ""),
                        "把手": e.get("handle", ""),
                        "敷居": e.get("sill", ""),
                        "敷居色": e.get("sill_color", ""),
                        "品番": e.get("hinban", ""),
                    })
                st.dataframe(rows, use_container_width=True)

        # Claude OCR生テキスト
        if st.session_state.get("claude_raw"):
            with st.expander("📊 木工事図面 Claude読み取り結果（生テキスト）"):
                st.code(st.session_state["claude_raw"])

        st.divider()
        st.download_button(
            label="📥 レポートをダウンロード",
            data=report,
            file_name="建具チェックレポート.txt",
            mime="text/plain",
            use_container_width=True
        )
    else:
        st.info("ファイルをアップロードして「チェック開始」を押してください")
