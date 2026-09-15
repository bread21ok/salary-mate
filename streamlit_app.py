"""
streamlit_app.py
=================
급여메이트 웹앱 - 3개 엑셀 파일을 업로드하면 급여내역서를 자동 계산해서
다운로드할 수 있게 해주는 Streamlit 앱.

실행 방법: streamlit run streamlit_app.py
"""

import streamlit as st
from datetime import date

import os
from main_calculator import main, detect_target_year_month

st.set_page_config(page_title="급여메이트", page_icon="💰", layout="centered")

st.title("💰 급여메이트")
st.caption("기간제근로자 급여 자동 계산 시스템")

st.markdown("---")

# 급여내역서 양식(빈 서식)은 매번 바뀌는 데이터가 아니라 앱에 고정으로 포함된 파일입니다.
# 이 파일과 같은 폴더(깃허브 저장소)에 "급여내역서_양식.xlsx" 를 함께 올려두어야 합니다.
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "급여내역서_양식.xlsx")

st.subheader("1. 파일 업로드")
st.caption("급여내역서 양식은 앱에 이미 포함되어 있어 따로 업로드하지 않아도 됩니다.")

op_file = st.file_uploader("운영현황.xlsx", type=["xlsx"])
personal_file = st.file_uploader("개인부담금.xlsx", type=["xlsx"])
local_file = st.file_uploader("자치단체부담금.xlsx", type=["xlsx"])

if not os.path.exists(TEMPLATE_PATH):
    st.error(
        "⚠️ 급여내역서_양식.xlsx 파일을 찾을 수 없습니다. "
        "이 파일을 깃허브 저장소에 streamlit_app.py와 같은 위치에 올려주세요."
    )

st.subheader("2. 처리 대상 연/월")

today = date.today()
c1, c2 = st.columns(2)
with c1:
    target_year = st.number_input("연도", min_value=2020, max_value=2100, value=today.year, step=1)
with c2:
    target_month = st.number_input("월", min_value=1, max_value=12, value=today.month, step=1)

st.markdown("---")

run_button = st.button("🚀 계산 실행", type="primary", use_container_width=True)

if run_button:
    missing = []
    if not op_file: missing.append("운영현황.xlsx")
    if not personal_file: missing.append("개인부담금.xlsx")
    if not local_file: missing.append("자치단체부담금.xlsx")

    if missing:
        st.error(f"다음 파일을 업로드해주세요: {', '.join(missing)}")
    elif not os.path.exists(TEMPLATE_PATH):
        st.error("급여내역서_양식.xlsx 파일이 앱 폴더에 없어 계산을 진행할 수 없습니다.")
    else:
        with st.spinner("계산 중입니다..."):
            try:
                result_bytesio, messages = main(
                    op_file,
                    personal_file,
                    local_file,
                    TEMPLATE_PATH,
                    int(target_year),
                    int(target_month),
                )
            except Exception as e:
                st.error(f"계산 중 오류가 발생했습니다: {e}")
                st.stop()

        st.success("계산이 완료되었습니다.")

        if messages:
            st.subheader("⚠️ 확인이 필요한 안내 사항")
            for m in messages:
                st.warning(m)
        else:
            st.info("특별한 경고/오류 없이 모든 인원이 정상 처리되었습니다.")

        st.subheader("3. 결과 다운로드")
        st.download_button(
            label="📥 급여내역서 다운로드",
            data=result_bytesio,
            file_name=f"급여내역서_{int(target_year)}년{int(target_month)}월.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
