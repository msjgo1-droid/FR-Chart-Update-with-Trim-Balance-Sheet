import io
import datetime
import streamlit as st
import openpyxl

import engine

st.set_page_config(page_title="FR TRIM 자동 채우기", page_icon="🧵", layout="centered")

st.title("🧵 FR Chart TRIM 자동 채우기")
st.caption(
    "여러 개의 Balance Sheet(스타일별) + FR(Friday Review) 차트 1개를 올리면, "
    "각 스타일의 TRIM 영역을 Balance Sheet 데이터로 자동으로 채운 FR 파일을 만들어줍니다."
)

with st.expander("ℹ️ 처리 규칙 (한번 확인해보세요)", expanded=False):
    st.markdown(
        """
- Balance Sheet의 **GENERAL TRIM / SPECIAL TRIM** 항목표에서 **Item / ETD / ETA / Tracking# / Shipped Qty** (1st order and ship date 기준) 만 가져옵니다.
- 값이 비어 있어도 1st shipment는 **그대로** 항목 라인을 만듭니다 (건너뛰지 않음).
- 거래처(Mill/Supplier) 이름은 가져오지 않습니다.
- 이름에 **"Hangtag"** 가 들어간 항목은 **사이즈별로 줄을 나눠서** 넣습니다.
- **2nd order and ship date**가 있는 항목은, ETD/ETA/Tracking#/Shipped Qty가 **전부 채워져 있을 때만** `항목명-2nd` 라인을 추가로 넣습니다 (비어있으면 추가하지 않음).
- FR 차트의 TRIM 칸(스타일당 1개 병합 셀)을 풀어서, **Item/ETD/ETA/Tracking#/Shipped Qty 5개 컬럼 x N개 행**의 실제 표(테두리 포함)로 바꿔 넣습니다.
- 필요한 줄 수가 기존 스타일 블록의 행 수보다 많으면, 해당 스타일 블록 전체를 자동으로 늘리고 (다른 컬럼들의 병합, 수식, 아래 스타일들 위치도 같이 안전하게 조정) 부족한 만큼 채웁니다.
- 헤더에 **BULK FB / TRIM** 그룹 제목과 구분선을 넣고, TRIM 아래에 컬럼 라벨(Item/ETD/ETA/Tracking#/Shipped Qty)을 넣습니다.
        """
    )

split_keyword = st.text_input(
    "사이즈별로 줄을 나눌 항목 키워드 (기본: hangtag)", value="hangtag"
).strip().lower() or "hangtag"

st.subheader("1. Balance Sheet 업로드 (여러 개 선택 가능, 스타일당 1개)")
bs_files = st.file_uploader(
    "Balance Sheet 파일들", type=["xlsx"], accept_multiple_files=True, key="bs_files"
)

st.subheader("2. FR (Friday Review) 차트 업로드 (1개)")
fr_file = st.file_uploader("FR 차트 파일", type=["xlsx"], key="fr_file")

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
            fr_sheet_name = st.selectbox("FR 파일 안에서 사용할 시트를 선택하세요", sheet_names)
    except Exception as e:
        st.error(f"FR 파일을 읽는 중 문제가 발생했습니다: {e}")

process_clicked = st.button("🚀 처리 시작", type="primary", disabled=not (bs_files and fr_file and fr_sheet_name))

if process_clicked:
    with st.spinner("처리 중입니다..."):
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
        except Exception as e:
            st.error(f"처리 중 오류가 발생했습니다: {e}")
            st.session_state.pop("result_bytes", None)

if "results" in st.session_state:
    st.subheader("처리 결과")
    for r in st.session_state["results"]:
        status = r.get("status", "")
        style = r.get("style", r.get("sheet"))
        if status == "완료":
            st.success(
                f"스타일 {style}: {r['lines']}줄 반영 완료 "
                f"(행 {r['block_start']}~{r['block_end']}"
                + (f", {r['rows_added']}행 추가" if r.get("rows_added") else "")
                + ")"
            )
        else:
            st.warning(f"스타일 {style} ({r.get('sheet')}): {status}")

if "result_bytes" in st.session_state:
    base_name = st.session_state.get("fr_filename", "FR_updated.xlsx")
    if base_name.lower().endswith(".xlsx"):
        base_name = base_name[:-5]
    out_name = f"{base_name}_TRIM_UPDATED_{datetime.date.today().isoformat()}.xlsx"
    st.download_button(
        "⬇️ 업데이트된 FR 파일 다운로드",
        data=st.session_state["result_bytes"],
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
