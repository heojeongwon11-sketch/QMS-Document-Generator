"""
관리계획서(CP) 기반 하위 문서 자동 생성 웹앱

생성 문서:
- 스크류 점검일지
- 설비일상점검표
- 셋업 & 일일 성형조건 기록지

실행 방법:
    python -m streamlit run app.py


"""
import io
import os
from datetime import datetime

import streamlit as st

from generate_screw import generate as generate_screw
from generate_equipment import generate as generate_equipment
from generate_molding import generate as generate_molding

APP_DIR = os.path.dirname(os.path.abspath(__file__))
SCREW_TEMPLATE_PATH = os.path.join(APP_DIR, "screw_template.xlsx")
EQUIPMENT_TEMPLATE_PATH = os.path.join(APP_DIR, "equipment_template.xlsx")
MOLDING_TEMPLATE_PATH = os.path.join(APP_DIR, "molding_template.xlsx")

st.set_page_config(page_title="하위문서 자동 생성기", page_icon="📋", layout="centered")

st.title("📋 하위문서 자동 생성기")
st.caption("관리계획서(CP)를 업로드하면 하위문서들을 자동으로 만들어 드립니다.")

st.subheader("1. 관리계획서(CP) 업로드")
cp_file = st.file_uploader(
    "CP 파일 (.xlsx 또는 .xls)",
    type=["xlsx", "xls"],
    accept_multiple_files=False,
    key="cp_uploader",
)
st.caption("※ .xls(구형) / .xlsx 모두 업로드 가능합니다. 한 번에 1개 파일만 업로드됩니다.")

if cp_file is None:
    st.info("먼저 CP 파일을 업로드해 주세요.")
    st.stop()

cp_bytes = cp_file.read()  # 여러 탭에서 재사용하기 위해 미리 읽어둠

st.subheader("2. 생성할 문서 선택")
tab_screw, tab_equipment, tab_molding = st.tabs(
    ["🔩 스크류 점검일지", "🛠️ 설비일상점검표", "📝 성형조건기록지"]
)

# ------------------------------------------------------------------
# 스크류 점검일지
# ------------------------------------------------------------------
with tab_screw:
    with st.expander("설정", expanded=False):
        use_custom_screw_template = st.toggle("스크류 점검일지 템플릿을 직접 업로드", value=False, key="screw_custom_template_toggle")
        custom_screw_template_file = None
        if use_custom_screw_template:
            custom_screw_template_file = st.file_uploader(
                "스크류 점검일지 템플릿(.xlsx) 업로드", type=["xlsx"], key="screw_template_uploader"
            )

    st.caption(
        "상단(특이사항 안내/마모도 판정표/표 헤더)은 템플릿 그대로 유지되고, "
        "제목의 설비호기 번호만 CP 값으로 바뀝니다. 로그 입력 영역은 빈 칸으로 생성됩니다."
    )

    if st.button("스크류 점검일지 생성하기", type="primary", key="screw_generate_btn"):
        try:
            template_source = custom_screw_template_file if custom_screw_template_file else SCREW_TEMPLATE_PATH
            output_buffer = io.BytesIO()

            result = generate_screw(
                cp_path=io.BytesIO(cp_bytes),
                template_path=template_source,
                output_path=output_buffer,
            )
            output_buffer.seek(0)

            st.success("스크류 점검일지 생성 완료!")
            st.write(f"설비호기: **{result['설비호기']}**")

            out_name = f"스크류_점검일지_{result['설비호기']}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            st.download_button(
                label="⬇️ 스크류 점검일지 다운로드",
                data=output_buffer,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="screw_download_btn",
            )
        except Exception as e:
            st.error(f"생성 중 오류가 발생했습니다: {e}")
            st.exception(e)

# ------------------------------------------------------------------
# 설비일상점검표
# ------------------------------------------------------------------
with tab_equipment:
    with st.expander("설정", expanded=False):
        use_custom_eq_template = st.toggle("설비일상점검표 템플릿을 직접 업로드", value=False, key="eq_custom_template_toggle")
        custom_eq_template_file = None
        if use_custom_eq_template:
            custom_eq_template_file = st.file_uploader(
                "설비일상점검표 템플릿(.xlsx) 업로드", type=["xlsx"], key="eq_template_uploader"
            )

    st.caption(
        "구분(건조기/온조기/사출기 등)은 템플릿 구성을 그대로 따르고, 점검항목/규격은 CP 값으로 채웁니다. "
        "연도·월·설비기종은 CP에 없는 정보라 빈 칸으로 둡니다. 날짜별 체크 칸은 매일 작업자가 채우는 영역입니다."
    )

    if st.button("설비일상점검표 생성하기", type="primary", key="eq_generate_btn"):
        try:
            template_source = custom_eq_template_file if custom_eq_template_file else EQUIPMENT_TEMPLATE_PATH
            output_buffer = io.BytesIO()

            result = generate_equipment(
                cp_path=io.BytesIO(cp_bytes),
                template_path=template_source,
                output_path=output_buffer,
            )
            output_buffer.seek(0)

            st.success("설비일상점검표 생성 완료!")
            st.write(f"설비호기: **{result['설비호기']}**")
            if result["CP에서_못찾은_항목"]:
                st.warning(f"CP에서 값을 찾지 못한 항목: {', '.join(result['CP에서_못찾은_항목'])}")

            out_name = f"설비일상점검표_{result['설비호기']}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            st.download_button(
                label="⬇️ 설비일상점검표 다운로드",
                data=output_buffer,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="eq_download_btn",
            )
        except Exception as e:
            st.error(f"생성 중 오류가 발생했습니다: {e}")
            st.exception(e)

# ------------------------------------------------------------------
# 셋업 & 일일 성형조건 기록지
# ------------------------------------------------------------------
with tab_molding:
    with st.expander("설정", expanded=False):
        use_custom_mold_template = st.toggle("성형조건기록지 템플릿을 직접 업로드", value=False, key="mold_custom_template_toggle")
        custom_mold_template_file = None
        if use_custom_mold_template:
            custom_mold_template_file = st.file_uploader(
                "성형조건기록지 템플릿(.xlsx) 업로드", type=["xlsx"], key="mold_template_uploader"
            )

    st.caption(
        "CP의 관리기록방법이 '셋업 & 일일 성형 조건 기록지'인 행을 위에서부터 No./공정단계/공정특성/"
        "규격치·공차·단위로 채웁니다. 표 시작 행·열 위치는 템플릿에서 자동으로 찾습니다. "
        "날짜별 성형 기록 칸은 작업자 수기 입력용이라 빈 칸입니다."
    )

    if st.button("성형조건기록지 생성하기", type="primary", key="mold_generate_btn"):
        try:
            template_source = custom_mold_template_file if custom_mold_template_file else MOLDING_TEMPLATE_PATH
            output_buffer = io.BytesIO()

            result = generate_molding(
                cp_path=io.BytesIO(cp_bytes),
                template_path=template_source,
                output_path=output_buffer,
            )
            output_buffer.seek(0)

            st.success("성형조건기록지 생성 완료!")
            st.write(f"제품번호: **{result['제품번호']}**")
            if result["CP에서_못찾은_항목"]:
                st.warning(f"CP에서 값을 찾지 못한 항목: {', '.join(result['CP에서_못찾은_항목'])}")

            out_name = f"성형조건기록지_{result['제품번호']}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            st.download_button(
                label="⬇️ 성형조건기록지 다운로드",
                data=output_buffer,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="mold_download_btn",
            )
        except Exception as e:
            st.error(f"생성 중 오류가 발생했습니다: {e}")
            st.exception(e)
