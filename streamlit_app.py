import streamlit as st
import sys
import os
from dotenv import load_dotenv
# override=True: シェル側に空のANTHROPIC_API_KEY等が残っていても.envの値で上書きする
load_dotenv(os.path.join(os.path.dirname(__file__), '.env'), override=True)

# ─── ページ設定（最初に呼ぶ必要がある）─────────────────────────────────────
st.set_page_config(
    page_title="建具図面チェッカー",
    page_icon="🏠",
    layout="wide",
)

# ─── app.py からビジネスロジックをインポート ─────────────────────────────────
from app import (
    parse_tategu_pdf,
    parse_mokuko_pdf,
    parse_mokuko_manual,
    parse_mokuko_from_text,
    parse_building_spec,
    check_cross_reference,
    check_required_fields,
    check_sill,
    check_na,
    check_taste,
    check_building_spec,
    check_closet_sill,
    check_floor_thickness,
    check_visual_items,
    build_report,
    build_tategu_summary,
    build_spec_summary,
    extract_property_name,
    ocr_images_with_claude,
    read_pdf_text,
)


def get_api_key() -> str:
    """Streamlit Secrets → 環境変数 ANTHROPIC_API_KEY の順でAPIキーを取得する"""
    try:
        key = st.secrets.get("ANTHROPIC_API_KEY", "")
        if key:
            return key
    except Exception:
        pass
    return os.environ.get("ANTHROPIC_API_KEY", "")


# ─── タイトル & 説明 ──────────────────────────────────────────────────────────
st.title("🏠 建具図面チェッカー")
st.markdown(
    """
**建具・床図面**（PDF）と **木工事図面**（PDF または画像）をアップロードすると、
WD番号の寸法・種類・敷居・品番などの整合性を自動チェックします。
CIDフォントPDFは Claude API ビジョンで自動読み取り対応。
"""
)

# ─── APIキー状態インジケータ ──────────────────────────────────────────────────
_initial_api_key = get_api_key()
if _initial_api_key:
    st.success(f"✅ Claude API キー設定済み（CIDフォントPDFも自動OCR可能）")
else:
    st.warning(
        "⚠️ Claude API キーが未設定です。CIDフォント由来の建具図面PDFはOCRできません。"
        "プロジェクトフォルダの `.env` に `ANTHROPIC_API_KEY=...` を設定して、Streamlitを再起動してください。"
    )

# ─── ファイルアップロード ─────────────────────────────────────────────────────
st.markdown("---")
col_left, col_right = st.columns(2)

with col_left:
    st.subheader("① 建具・床図面（必須）")
    tategu_file = st.file_uploader(
        "PDF をアップロード",
        type=["pdf"],
        key="tategu",
        help="建具・床図面PDFをアップロードしてください（WD①②③形式）",
    )

with col_right:
    st.subheader("② 木工事図面")
    mokuko_mode = st.radio(
        "入力方法を選択",
        ["画像（1F / 2F）", "PDF", "手動入力"],
        horizontal=True,
        key="mokuko_mode",
    )

    floor1_file = floor2_file = mokuko_file = None
    mokuko_manual = ""

    if mokuko_mode == "画像（1F / 2F）":
        floor1_file = st.file_uploader(
            "1階 建具スケジュール画像",
            type=["png", "jpg", "jpeg"],
            key="floor1",
        )
        floor2_file = st.file_uploader(
            "2階 建具スケジュール画像",
            type=["png", "jpg", "jpeg"],
            key="floor2",
        )
    elif mokuko_mode == "PDF":
        mokuko_file = st.file_uploader(
            "木工事図面 PDF",
            type=["pdf"],
            key="mokuko",
        )
    else:
        mokuko_manual = st.text_area(
            "WDデータを貼り付け",
            placeholder="例: WD101,片開きドア,W654,H2035\nWD102,片引き戸,W778,H2035",
            height=160,
            key="manual",
        )

# 参照資料
st.markdown("---")
st.subheader("③ 参照資料（任意）")
ref_file = st.file_uploader(
    "テイスト規定PDF など",
    type=["pdf"],
    key="ref",
    help="テイスト整合性チェックに使用します。なくてもチェック可能です。",
)

# ─── チェック実行ボタン ───────────────────────────────────────────────────────
st.markdown("---")
run_btn = st.button("▶ チェック実行", type="primary", use_container_width=True)

if run_btn:
    # 入力バリデーション
    if not tategu_file:
        st.error("建具・床図面をアップロードしてください。")
        st.stop()

    if mokuko_mode == "画像（1F / 2F）" and not floor1_file and not floor2_file:
        st.error("1階または2階の建具スケジュール画像をアップロードしてください。")
        st.stop()

    if mokuko_mode == "PDF" and not mokuko_file:
        st.error("木工事図面PDFをアップロードしてください。")
        st.stop()

    if mokuko_mode == "手動入力" and not mokuko_manual.strip():
        st.error("WDデータを入力してください。")
        st.stop()

    with st.spinner("チェック中… しばらくお待ちください"):
        api_key = get_api_key()

        # ── ファイル読み込み ──
        tategu_bytes = tategu_file.read()
        ref_bytes = ref_file.read() if ref_file else None

        # ── テキスト抽出 ──
        tategu_pages = read_pdf_text(tategu_bytes)
        tategu_text = "\n".join(p["text"] for p in tategu_pages)

        ref_text = ""
        if ref_bytes:
            ref_pages = read_pdf_text(ref_bytes)
            ref_text = "\n".join(p["text"] for p in ref_pages)

        # ── 建具エントリ解析（CIDフォントPDFは Claude OCR フォールバック） ──
        # parse_tategu_pdf 内のprint出力を捕捉してUIに表示
        import io as _io
        from contextlib import redirect_stdout as _rstdout
        _parse_buf = _io.StringIO()
        with _rstdout(_parse_buf):
            tategu_entries, tategu_ocr_spec, tategu_ocr_raw = parse_tategu_pdf(tategu_bytes, api_key or None)
        tategu_parse_log = _parse_buf.getvalue()

        tategu_err = None
        if len(tategu_entries) == 0:
            from app import is_cid_font_pdf as _cidchk
            _full_t = "\n".join(p["text"] for p in read_pdf_text(tategu_bytes))
            if _cidchk(_full_t):
                if not api_key:
                    tategu_err = (
                        "建具図面PDFがCIDフォント由来でテキスト抽出できません。"
                        "Claude APIキーが未設定のためOCRも実行できませんでした。"
                        "`.env` の `ANTHROPIC_API_KEY` を設定してください。"
                    )
                else:
                    tategu_err = (
                        "建具図面PDFがCIDフォント由来です。Claude OCRを試みましたが"
                        "WDエントリを抽出できませんでした。"
                    )
            else:
                tategu_err = (
                    "建具図面PDFからWDエントリを1件も抽出できませんでした。"
                    "PDFの形式や記載内容を確認してください。"
                )

        # ── 建具側パースログをUIに表示 ──
        if tategu_parse_log.strip():
            with st.expander("🔍 建具PDFパース詳細ログ（クリックで展開）"):
                st.code(tategu_parse_log, language="text")
                st.caption(
                    "チェックポイント: `entries_from_pdfplumber` が0なら pdfplumberで読めていません。"
                    "`cid_detected=True` ならCIDフォントPDFでOCRが必要です。"
                    "`api_key=YES` かつ `Calling Claude API OCR fallback...` が出ればOCRが走っています。"
                )

        # ── Claude OCR の生出力をUIに表示（OCRが走った場合のみ） ──
        if tategu_ocr_raw:
            with st.expander("📋 Claude OCR の生出力（クリックで展開／フィールド欠落の確認用）"):
                for i, page_text in enumerate(tategu_ocr_raw):
                    st.markdown(f"**ページ {i+1}**")
                    st.code(page_text, language="text")
                st.caption(
                    "各WD行は9フィールド（8個の「|」）固定。区切り数が足りない場合や、"
                    "画像で見えている値（品番・敷居色など）が空欄になっている場合は、Claude OCRが読み落としています。"
                    "その場合はプロンプト調整が必要です。"
                )

        # ── 木工事エントリ解析 ──
        mokuko_entries = {}
        mokuko_text = ""
        ocr_used = False
        ocr_err = None
        claude_raw = None

        if mokuko_mode == "手動入力":
            mokuko_entries = parse_mokuko_manual(mokuko_manual)

        elif mokuko_mode == "画像（1F / 2F）":
            image_bytes_list = []
            if floor1_file:
                image_bytes_list.append(floor1_file.read())
            if floor2_file:
                image_bytes_list.append(floor2_file.read())

            if not api_key:
                ocr_err = "APIキーが設定されていません。Streamlit Secretsに ANTHROPIC_API_KEY を設定してください。"
            else:
                try:
                    claude_raw = ocr_images_with_claude(image_bytes_list, api_key)
                    mokuko_entries = parse_mokuko_from_text(claude_raw)
                    ocr_used = True
                except Exception as e:
                    ocr_err = f"画像読み取りエラー: {e}"

        elif mokuko_mode == "PDF":
            mokuko_bytes = mokuko_file.read()
            pages = read_pdf_text(mokuko_bytes)
            mokuko_text = "\n".join(p["text"] for p in pages)
            mokuko_entries, ocr_used, ocr_err, claude_raw = parse_mokuko_pdf(
                mokuko_bytes, api_key or None
            )

        full_text = mokuko_text + "\n" + tategu_text
        property_name = extract_property_name(full_text)

        # ── チェック実行 ──
        all_errors, all_ok = [], []

        if ocr_err:
            all_errors.append(
                f"**木工事図面の読み取りエラー**\n"
                f"  - {ocr_err}\n"
                f"  - 木工事図面の照合は目視確認が必要です"
            )

        if tategu_err:
            all_errors.append(
                f"**建具図面の読み取りエラー**\n"
                f"  - {tategu_err}\n"
                f"  - 建具図面側のWD情報が取れていないため、以降の照合・必須項目チェックは目視確認が必要です"
            )

        errs, oks = check_cross_reference(mokuko_entries, tategu_entries)
        all_errors.extend(errs)
        all_ok.extend(oks)

        errs, oks = check_required_fields(tategu_entries)
        all_errors.extend(errs)
        all_ok.extend(oks)

        errs, oks = check_sill(tategu_entries)
        all_errors.extend(errs)
        all_ok.extend(oks)

        errs = check_na([mokuko_text, tategu_text])
        all_errors.extend(errs)

        errs, oks = check_taste(full_text, ref_text, tategu_entries)
        all_errors.extend(errs)
        all_ok.extend(oks)

        # OCRが走っていればOCR抽出のspecを優先、無ければpdfplumberテキストからパース
        bldg_spec = tategu_ocr_spec if tategu_ocr_spec else parse_building_spec(tategu_text)
        errs, oks = check_building_spec(bldg_spec)
        all_errors.extend(errs)
        all_ok.extend(oks)

        errs, oks = check_closet_sill(tategu_entries)
        all_errors.extend(errs)
        all_ok.extend(oks)

        errs, oks = check_floor_thickness(bldg_spec)
        all_errors.extend(errs)
        all_ok.extend(oks)

        all_errors.extend(check_visual_items())

        report = build_report(property_name, all_errors, all_ok, tategu_entries, mokuko_entries)
        summary = build_tategu_summary(tategu_entries)
        spec_md = build_spec_summary(bldg_spec)

    # ── 結果表示 ──
    st.success(
        f"チェック完了 — 建具図面: {len(tategu_entries)}件 ／ 木工事図面: {len(mokuko_entries)}件 抽出"
    )
    if ocr_used:
        st.info("木工事図面: Claude API ビジョンで読み取りました")

    tab1, tab2, tab3 = st.tabs(["📋 チェックレポート", "📐 建具一覧", "🎨 仕様サマリー"])

    with tab1:
        st.markdown(report)

    with tab2:
        st.markdown(summary)

    with tab3:
        st.markdown(spec_md)

    if claude_raw:
        with st.expander("Claude OCR 生データ（デバッグ用）"):
            st.text(claude_raw)
