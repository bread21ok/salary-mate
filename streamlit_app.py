"""
streamlit_app.py
=================
급여메이트 웹앱
왼쪽 사이드바 메뉴: 급여계산 / 챗봇 / 참고자료집

실행 방법: streamlit run streamlit_app.py
"""

import os
from datetime import date

import streamlit as st

from main_calculator import main as run_payroll_calculation

# ---------------------------------------------------------------------
# 기본 설정 + 디자인(CSS)
# ---------------------------------------------------------------------
st.set_page_config(page_title="급여메이트", page_icon="💰", layout="wide")

CUSTOM_CSS = """
<style>
/* 전체 배경 & 기본 폰트 */
.stApp {
    background-color: #F7F8FA;
}

/* 사이드바 스타일 */
section[data-testid="stSidebar"] {
    background-color: #1F2937;
}
section[data-testid="stSidebar"] * {
    color: #F3F4F6 !important;
}
section[data-testid="stSidebar"] .stRadio > label {
    font-size: 1.05rem;
}

/* 카드처럼 보이는 컨테이너 */
div[data-testid="stVerticalBlockBorderWrapper"] {
    background-color: white;
    border-radius: 14px;
    padding: 0.5rem;
}

/* 버튼 */
.stButton > button, .stDownloadButton > button {
    border-radius: 10px;
    font-weight: 600;
}

/* 타이틀 영역 */
.app-header {
    padding: 0.6rem 0 1.2rem 0;
}
.app-header h1 {
    margin-bottom: 0;
}
.app-header p {
    color: #6B7280;
    margin-top: 0.2rem;
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "급여내역서_양식.xlsx")
REFERENCE_DIR = os.path.join(os.path.dirname(__file__), "reference_docs")

# ---------------------------------------------------------------------
# 사이드바 메뉴
# ---------------------------------------------------------------------
with st.sidebar:
    st.markdown("## 💰 급여메이트")
    st.caption("기간제근로자 급여 자동화 도구")
    st.markdown("---")
    page = st.radio(
        "메뉴",
        options=["급여계산", "챗봇", "참고자료집"],
        label_visibility="collapsed",
    )

st.markdown(
    '<div class="app-header"><h1>💰 급여메이트</h1>'
    '<p>기간제근로자 급여 자동 계산 · 상담 챗봇 · 참고자료 모음</p></div>',
    unsafe_allow_html=True,
)


# =======================================================================
# 페이지 1: 급여계산
# =======================================================================
def render_salary_page():
    with st.container(border=True):
        st.subheader("1. 파일 업로드")
        st.caption("급여내역서 양식은 앱에 이미 포함되어 있어 따로 업로드하지 않아도 됩니다.")

        col1, col2, col3 = st.columns(3)
        with col1:
            op_file = st.file_uploader("운영현황.xlsx", type=["xlsx"])
        with col2:
            personal_file = st.file_uploader("개인부담금.xlsx", type=["xlsx"])
        with col3:
            local_file = st.file_uploader("자치단체부담금.xlsx", type=["xlsx"])

        if not os.path.exists(TEMPLATE_PATH):
            st.error(
                "⚠️ 급여내역서_양식.xlsx 파일을 찾을 수 없습니다. "
                "이 파일을 깃허브 저장소에 streamlit_app.py와 같은 위치에 올려주세요."
            )

    with st.container(border=True):
        st.subheader("2. 처리 대상 연/월")
        today = date.today()
        c1, c2 = st.columns(2)
        with c1:
            target_year = st.number_input("연도", min_value=2020, max_value=2100, value=today.year, step=1)
        with c2:
            target_month = st.number_input("월", min_value=1, max_value=12, value=today.month, step=1)

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
                    result_bytesio, messages = run_payroll_calculation(
                        op_file, personal_file, local_file, TEMPLATE_PATH,
                        int(target_year), int(target_month),
                    )
                except Exception as e:
                    st.error(f"계산 중 오류가 발생했습니다: {e}")
                    st.stop()

            st.success("계산이 완료되었습니다.")

            with st.container(border=True):
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


# =======================================================================
# 페이지 2: 챗봇 (Claude API 연동)
# =======================================================================
SYSTEM_PROMPT = (
    "당신은 '급여메이트' 웹앱의 상담 챗봇입니다. "
    "기간제근로자 급여, 4대보험, 소득세, 이 앱의 사용법에 대한 질문에 "
    "친절하고 쉽게 한국어로 답변하세요. 확실하지 않은 법령/세율 정보는 "
    "반드시 관할 기관(국세청, 4대보험 공단 등)에 확인하라고 안내하세요."
)


def render_chatbot_page():
    with st.container(border=True):
        st.subheader("💬 급여메이트 챗봇")
        st.caption("급여·4대보험·세금이나 이 앱 사용법에 대해 자유롭게 물어보세요.")

    api_key = st.secrets.get("ANTHROPIC_API_KEY", None)
    if not api_key:
        st.error(
            "⚠️ Claude API 키가 설정되어 있지 않습니다.\n\n"
            "Streamlit Cloud 앱 관리 화면 → Settings → Secrets 에서\n"
            'ANTHROPIC_API_KEY = "여기에_API_키_입력"\n'
            "형식으로 추가해주세요."
        )
        return

    try:
        import anthropic
    except ImportError:
        st.error("anthropic 패키지가 설치되어 있지 않습니다. requirements.txt에 'anthropic'을 추가해주세요.")
        return

    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []

    for msg in st.session_state.chat_history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    user_input = st.chat_input("궁금한 점을 입력하세요...")

    if user_input:
        st.session_state.chat_history.append({"role": "user", "content": user_input})
        with st.chat_message("user"):
            st.markdown(user_input)

        with st.chat_message("assistant"):
            placeholder = st.empty()
            try:
                client = anthropic.Anthropic(api_key=api_key)
                response = client.messages.create(
                    model="claude-sonnet-5",
                    max_tokens=1024,
                    system=SYSTEM_PROMPT,
                    messages=[
                        {"role": m["role"], "content": m["content"]}
                        for m in st.session_state.chat_history
                    ],
                )
                answer = "".join(
                    block.text for block in response.content if block.type == "text"
                )
            except Exception as e:
                answer = f"오류가 발생했습니다: {e}"

            placeholder.markdown(answer)

        st.session_state.chat_history.append({"role": "assistant", "content": answer})

    if st.session_state.chat_history:
        if st.button("🗑️ 대화 초기화"):
            st.session_state.chat_history = []
            st.rerun()


# =======================================================================
# 페이지 3: 참고자료집
# =======================================================================
def render_reference_page():
    with st.container(border=True):
        st.subheader("📚 참고자료집")
        st.caption("담당자가 올려둔 참고 문서를 다운로드할 수 있습니다.")

        if not os.path.isdir(REFERENCE_DIR):
            st.info("아직 등록된 참고자료가 없습니다.")
            return

        files = sorted(
            f for f in os.listdir(REFERENCE_DIR)
            if os.path.isfile(os.path.join(REFERENCE_DIR, f)) and f != "README.txt"
        )

        if not files:
            st.info(
                "아직 등록된 참고자료가 없습니다.\n\n"
                "관리자: 깃허브 저장소의 reference_docs 폴더에 파일을 올리면 여기에 자동으로 나타납니다."
            )
            return

        for filename in files:
            filepath = os.path.join(REFERENCE_DIR, filename)
            with open(filepath, "rb") as f:
                file_bytes = f.read()

            col1, col2 = st.columns([4, 1])
            with col1:
                st.markdown(f"📄 **{filename}**")
            with col2:
                st.download_button(
                    "다운로드", data=file_bytes, file_name=filename,
                    key=f"dl_{filename}", use_container_width=True,
                )
            st.markdown("---")


# =======================================================================
# 라우팅
# =======================================================================
if page == "급여계산":
    render_salary_page()
elif page == "챗봇":
    render_chatbot_page()
elif page == "참고자료집":
    render_reference_page()
