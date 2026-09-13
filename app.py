import io
import datetime
import streamlit as st
import openpyxl

import engine

st.set_page_config(page_title="FR TRIM Auto-Fill", page_icon="🧵", layout="centered")

# --------------------------------------------------------------------------
# i18n
# --------------------------------------------------------------------------

LANGS = {"ko": "한국어", "en": "English", "vi": "Tiếng Việt"}

TEXT = {
    "ko": {
        "title": "🧵 FR Chart TRIM 자동 채우기",
        "caption": (
            "여러 개의 Balance Sheet(스타일별) + FR(Friday Review) 차트 1개를 올리면, "
            "각 스타일의 TRIM 영역을 Balance Sheet 데이터로 자동으로 채운 FR 파일을 만들어줍니다."
        ),
        "rules_expander": "ℹ️ 처리 규칙 (한번 확인해보세요)",
        "rules_md": """
- Balance Sheet의 **GENERAL TRIM / SPECIAL TRIM** 항목표에서 **Item / ETD / ETA / Tracking# / Shipped Qty** (1st order and ship date 기준) 만 가져옵니다.
- 값이 비어 있어도 1st shipment는 **그대로** 항목 라인을 만듭니다 (건너뛰지 않음).
- 거래처(Mill/Supplier) 이름은 가져오지 않습니다.
- 이름에 **"Hangtag"** 가 들어간 항목은 **사이즈별로 줄을 나눠서** 넣습니다.
- **2nd order and ship date**가 있는 항목은, ETD/ETA/Tracking#/Shipped Qty가 **전부 채워져 있을 때만** `항목명-2nd` 라인을 추가로 넣습니다 (비어있으면 추가하지 않음).
- FR 차트의 TRIM 칸(스타일당 1개 병합 셀)을 풀어서, **Item/ETD/ETA/Tracking#/Shipped Qty 5개 컬럼 x N개 행**의 실제 표(테두리 포함)로 바꿔 넣습니다.
- 필요한 줄 수가 기존 스타일 블록의 행 수보다 많으면, 해당 스타일 블록 전체를 자동으로 늘리고 (다른 컬럼들의 병합, 수식, 아래 스타일들 위치도 같이 안전하게 조정) 부족한 만큼 채웁니다.
- 헤더에 **BULK FB / TRIM** 그룹 제목과 구분선을 넣고, TRIM 아래에 컬럼 라벨(Item/ETD/ETA/Tracking#/Shipped Qty)을 넣습니다.
""",
        "split_keyword_label": "사이즈별로 줄을 나눌 항목 키워드 (기본: hangtag)",
        "step1_header": "1. Balance Sheet 업로드 (여러 개 선택 가능, 스타일당 1개)",
        "bs_uploader_label": "Balance Sheet 파일들",
        "step2_header": "2. FR (Friday Review) 차트 업로드 (1개)",
        "fr_uploader_label": "FR 차트 파일",
        "sheet_select_label": "FR 파일 안에서 사용할 시트를 선택하세요",
        "fr_read_error": "FR 파일을 읽는 중 문제가 발생했습니다: {e}",
        "process_button": "🚀 처리 시작",
        "spinner": "처리 중입니다...",
        "process_error": "처리 중 오류가 발생했습니다: {e}",
        "results_header": "처리 결과",
        "result_ok": "스타일 {style}: {lines}줄 반영 완료 (행 {start}~{end}{added})",
        "rows_added_suffix": ", {n}행 추가",
        "result_warn": "스타일 {style} ({sheet}): {status}",
        "status_no_style": "Balance Sheet에서 스타일 번호를 찾지 못했습니다",
        "status_no_trim_section": "GENERAL TRIM 섹션을 찾지 못했습니다",
        "status_style_not_found_in_fr": "FR 차트에서 이 스타일을 찾지 못했습니다",
        "status_fr_columns_not_found": "FR 차트에서 STYLE NO / TRIM 컬럼을 찾지 못했습니다",
        "download_button": "⬇️ 업데이트된 FR 파일 다운로드",
        "lang_label": "Language / 언어 / Ngôn ngữ",
    },
    "en": {
        "title": "🧵 FR Chart TRIM Auto-Fill",
        "caption": (
            "Upload multiple Balance Sheets (one per style) plus one FR (Friday Review) chart, "
            "and this will automatically fill in each style's TRIM section using the Balance Sheet data."
        ),
        "rules_expander": "ℹ️ How it works (worth a quick read)",
        "rules_md": """
- Pulls only **Item / ETD / ETA / Tracking# / Shipped Qty** (1st order and ship date) from the Balance Sheet's **GENERAL TRIM / SPECIAL TRIM** item table.
- Blank values **are still included** as their own line for the 1st shipment (nothing is skipped).
- Vendor / Mill-Supplier name is **not** included, to keep it as condensed as possible.
- Any item whose name contains **"Hangtag"** is **split into one line per size**.
- If an item has a **2nd order and ship date**, an extra `<item name>-2nd` line is added **only when** ETD/ETA/Tracking#/Shipped Qty are **all filled in** (blanks are not carried over for 2nd+ shipments).
- The FR chart's TRIM cell (one merged cell per style) is unmerged and rebuilt as a real bordered table: **5 columns (Item / ETD / ETA / Tracking# / Shipped Qty) x N rows**.
- If more rows are needed than the style's current row block has, the whole block is safely grown (other merged columns, formulas, and the position of styles below are all adjusted automatically).
- Header banners **BULK FB / TRIM** with a divider, plus column labels (Item/ETD/ETA/Tracking#/Shipped Qty) under TRIM, are set up automatically.
""",
        "split_keyword_label": "Keyword for items to split by size (default: hangtag)",
        "step1_header": "1. Upload Balance Sheets (multiple allowed, one per style)",
        "bs_uploader_label": "Balance Sheet files",
        "step2_header": "2. Upload the FR (Friday Review) chart (one file)",
        "fr_uploader_label": "FR chart file",
        "sheet_select_label": "Choose which sheet to use in the FR file",
        "fr_read_error": "There was a problem reading the FR file: {e}",
        "process_button": "🚀 Start processing",
        "spinner": "Processing...",
        "process_error": "An error occurred while processing: {e}",
        "results_header": "Results",
        "result_ok": "Style {style}: {lines} line(s) written (rows {start}-{end}{added})",
        "rows_added_suffix": ", {n} row(s) added",
        "result_warn": "Style {style} ({sheet}): {status}",
        "status_no_style": "Could not find a style number in this Balance Sheet",
        "status_no_trim_section": "Could not find a GENERAL TRIM section",
        "status_style_not_found_in_fr": "This style was not found in the FR chart",
        "status_fr_columns_not_found": "Could not find the STYLE NO / TRIM columns in the FR chart",
        "download_button": "⬇️ Download updated FR file",
        "lang_label": "Language / 언어 / Ngôn ngữ",
    },
    "vi": {
        "title": "🧵 Tự động điền TRIM cho FR Chart",
        "caption": (
            "Tải lên nhiều Balance Sheet (mỗi style một file) cùng với 1 file FR (Friday Review) chart, "
            "công cụ sẽ tự động điền phần TRIM của từng style bằng dữ liệu từ Balance Sheet."
        ),
        "rules_expander": "ℹ️ Quy tắc xử lý (nên xem qua một lần)",
        "rules_md": """
- Chỉ lấy **Item / ETD / ETA / Tracking# / Shipped Qty** (theo 1st order and ship date) từ bảng mục **GENERAL TRIM / SPECIAL TRIM** trong Balance Sheet.
- Với 1st shipment, dòng vẫn được tạo **kể cả khi giá trị để trống** (không bỏ qua).
- **Không** lấy tên nhà cung cấp (Mill/Supplier) để giữ nội dung gọn nhất có thể.
- Các mục có tên chứa **"Hangtag"** sẽ được **tách riêng theo từng size**.
- Nếu mục có **2nd order and ship date**, dòng bổ sung `<tên mục>-2nd` chỉ được thêm vào **khi cả 4 giá trị** ETD/ETA/Tracking#/Shipped Qty **đều có đầy đủ** (nếu còn trống thì không thêm).
- Ô TRIM trong FR chart (1 ô gộp cho mỗi style) sẽ được tách ra và dựng lại thành bảng thật có viền: **5 cột (Item / ETD / ETA / Tracking# / Shipped Qty) x N dòng**.
- Nếu số dòng cần nhiều hơn số hàng hiện có của khối style đó, toàn bộ khối sẽ được **tự động mở rộng** (các cột gộp khác, công thức, và vị trí của các style phía dưới đều được điều chỉnh an toàn).
- Tiêu đề nhóm **BULK FB / TRIM** cùng đường phân cách, và nhãn cột (Item/ETD/ETA/Tracking#/Shipped Qty) bên dưới TRIM sẽ được tự động thêm vào.
""",
        "split_keyword_label": "Từ khóa để tách dòng theo size (mặc định: hangtag)",
        "step1_header": "1. Tải lên Balance Sheet (có thể chọn nhiều file, mỗi style 1 file)",
        "bs_uploader_label": "Các file Balance Sheet",
        "step2_header": "2. Tải lên FR (Friday Review) chart (1 file)",
        "fr_uploader_label": "File FR chart",
        "sheet_select_label": "Chọn sheet muốn dùng trong file FR",
        "fr_read_error": "Có lỗi khi đọc file FR: {e}",
        "process_button": "🚀 Bắt đầu xử lý",
        "spinner": "Đang xử lý...",
        "process_error": "Đã xảy ra lỗi trong quá trình xử lý: {e}",
        "results_header": "Kết quả xử lý",
        "result_ok": "Style {style}: đã ghi {lines} dòng (hàng {start}-{end}{added})",
        "rows_added_suffix": ", đã thêm {n} hàng",
        "result_warn": "Style {style} ({sheet}): {status}",
        "status_no_style": "Không tìm thấy mã style trong Balance Sheet này",
        "status_no_trim_section": "Không tìm thấy phần GENERAL TRIM",
        "status_style_not_found_in_fr": "Không tìm thấy style này trong FR chart",
        "status_fr_columns_not_found": "Không tìm thấy cột STYLE NO / TRIM trong FR chart",
        "download_button": "⬇️ Tải file FR đã cập nhật",
        "lang_label": "Language / 언어 / Ngôn ngữ",
    },
}

if "lang" not in st.session_state:
    st.session_state.lang = "ko"

lang_code = st.selectbox(
    TEXT[st.session_state.lang]["lang_label"],
    options=list(LANGS.keys()),
    format_func=lambda k: LANGS[k],
    index=list(LANGS.keys()).index(st.session_state.lang),
    key="lang_selector",
)
st.session_state.lang = lang_code
T = TEXT[lang_code]


def status_text(code):
    return T.get(f"status_{code}", code)


# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------

st.title(T["title"])
st.caption(T["caption"])

with st.expander(T["rules_expander"], expanded=False):
    st.markdown(T["rules_md"])

split_keyword = (st.text_input(T["split_keyword_label"], value="hangtag").strip().lower() or "hangtag")

st.subheader(T["step1_header"])
bs_files = st.file_uploader(T["bs_uploader_label"], type=["xlsx"], accept_multiple_files=True, key="bs_files")

st.subheader(T["step2_header"])
fr_file = st.file_uploader(T["fr_uploader_label"], type=["xlsx"], key="fr_file")

fr_sheet_name = None
if fr_file is not None:
    try:
        fr_file.seek(0)
        probe_wb = openpyxl.load_workbook(io.BytesIO(fr_file.read()), read_only=True)
        sheet_names = probe_wb.sheetnames
        fr_file.seek(0)
        if len(sheet_names) == 1:
            fr_sheet_name = sheet_names[0]
        else:
            fr_sheet_name = st.selectbox(T["sheet_select_label"], sheet_names)
    except Exception as e:
        st.error(T["fr_read_error"].format(e=e))

process_clicked = st.button(
    T["process_button"], type="primary", disabled=not (bs_files and fr_file and fr_sheet_name)
)

if process_clicked:
    with st.spinner(T["spinner"]):
        try:
            fr_file.seek(0)
            fr_wb = openpyxl.load_workbook(io.BytesIO(fr_file.read()))

            bs_wbs = []
            for f in bs_files:
                f.seek(0)
                bs_wbs.append(openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True))

            results = engine.process(fr_wb, fr_sheet_name, bs_wbs, split_keyword=split_keyword)

            buf = io.BytesIO()
            fr_wb.save(buf)
            buf.seek(0)

            st.session_state["result_bytes"] = buf.getvalue()
            st.session_state["results"] = results
            st.session_state["fr_filename"] = fr_file.name
        except engine.EngineError as e:
            st.error(T["process_error"].format(e=status_text(str(e))))
            st.session_state.pop("result_bytes", None)
        except Exception as e:
            st.error(T["process_error"].format(e=e))
            st.session_state.pop("result_bytes", None)

if "results" in st.session_state:
    st.subheader(T["results_header"])
    for r in st.session_state["results"]:
        status = r.get("status", "")
        style = r.get("style", r.get("sheet"))
        if status == "ok":
            added = T["rows_added_suffix"].format(n=r["rows_added"]) if r.get("rows_added") else ""
            st.success(
                T["result_ok"].format(
                    style=style, lines=r["lines"], start=r["block_start"], end=r["block_end"], added=added
                )
            )
        else:
            st.warning(T["result_warn"].format(style=style, sheet=r.get("sheet"), status=status_text(status)))

if "result_bytes" in st.session_state:
    base_name = st.session_state.get("fr_filename", "FR_updated.xlsx")
    if base_name.lower().endswith(".xlsx"):
        base_name = base_name[:-5]
    out_name = f"{base_name}_TRIM_UPDATED_{datetime.date.today().isoformat()}.xlsx"
    st.download_button(
        T["download_button"],
        data=st.session_state["result_bytes"],
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
