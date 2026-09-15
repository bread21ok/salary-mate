"""
streamlit_app.py
=================
급여메이트 웹앱
왼쪽 사이드바 메뉴: 급여계산 / 자료 검색 / 참고자료집

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
        options=["급여계산", "자료 검색", "참고자료집"],
        label_visibility="collapsed",
    )

st.markdown(
    '<div class="app-header"><h1>💰 급여메이트</h1>'
    '<p>기간제근로자 급여 자동 계산 · 자료 검색 · 참고자료 모음</p></div>',
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
# 페이지 2: 자료 검색 (AI 없이, 참고자료집 안에서 키워드로 관련 내용을 찾아 보여줌)
# =======================================================================
SUPPORTED_EXTENSIONS = {"txt", "pdf", "docx"}


def _get_reference_version_key():
    """reference_docs 폴더 안 파일 목록/수정시각을 캐시 판단용 키로 만든다."""
    if not os.path.isdir(REFERENCE_DIR):
        return ()
    items = []
    for fname in sorted(os.listdir(REFERENCE_DIR)):
        path = os.path.join(REFERENCE_DIR, fname)
        if os.path.isfile(path) and fname != "README.txt":
            items.append((fname, os.path.getmtime(path), os.path.getsize(path)))
    return tuple(items)


@st.cache_data(show_spinner="참고자료집을 읽는 중입니다...")
def _load_reference_texts(_version_key):
    """reference_docs 폴더의 파일들을 읽어 {파일명: 텍스트} 형태로 반환한다."""
    texts = {}
    unsupported = []

    if not os.path.isdir(REFERENCE_DIR):
        return texts, unsupported

    for fname in sorted(os.listdir(REFERENCE_DIR)):
        path = os.path.join(REFERENCE_DIR, fname)
        if not os.path.isfile(path) or fname == "README.txt":
            continue

        ext = fname.lower().rsplit(".", 1)[-1] if "." in fname else ""

        if ext not in SUPPORTED_EXTENSIONS:
            unsupported.append(fname)
            continue

        try:
            if ext == "txt":
                with open(path, encoding="utf-8", errors="ignore") as f:
                    content = f.read()
            elif ext == "pdf":
                from pypdf import PdfReader
                reader = PdfReader(path)
                content = "\n".join((page.extract_text() or "") for page in reader.pages)
            elif ext == "docx":
                import docx
                doc = docx.Document(path)
                content = "\n".join(p.text for p in doc.paragraphs)
            else:
                content = ""
            texts[fname] = content.strip()
        except Exception as e:
            texts[fname] = f"[읽기 오류: {e}]"

    return texts, unsupported


def _split_into_chunks(text: str, max_chunk_chars: int = 350) -> list:
    """문서 텍스트를 검색하기 좋은 크기(약 350자)의 덩어리(청크)로 나눈다."""
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    chunks = []
    buffer = ""
    for line in lines:
        if buffer and len(buffer) + len(line) + 1 > max_chunk_chars:
            chunks.append(buffer)
            buffer = line
        else:
            buffer = (buffer + " " + line).strip()
    if buffer:
        chunks.append(buffer)
    return chunks


def _search_reference(query: str, texts: dict, top_k: int = 5) -> list:
    """
    query에 포함된 단어들이 얼마나 등장하는지로 각 문서 청크에 점수를 매겨
    가장 관련성 높은 상위 top_k개를 반환한다.
    반환 항목: {"file": 파일명, "text": 청크내용, "score": 점수}
    """
    keywords = [w for w in query.strip().split() if len(w) >= 1]
    if not keywords:
        return []

    results = []
    for fname, content in texts.items():
        if not content:
            continue
        for chunk in _split_into_chunks(content):
            score = 0
            for kw in keywords:
                score += chunk.count(kw)
            if query.strip() in chunk:
                score += 5  # 질문 문장 전체가 그대로 들어있으면 가산점
            if score > 0:
                results.append({"file": fname, "text": chunk, "score": score})

    results.sort(key=lambda r: r["score"], reverse=True)
    return results[:top_k]


def _highlight(text: str, query: str) -> str:
    """검색어에 포함된 단어들을 텍스트에서 굵게 강조 표시한다."""
    import re
    keywords = sorted({w for w in query.strip().split() if len(w) >= 1}, key=len, reverse=True)
    highlighted = text
    for kw in keywords:
        pattern = re.escape(kw)
        highlighted = re.sub(f"({pattern})", r"**\1**", highlighted)
    return highlighted


def render_search_page():
    with st.container(border=True):
        st.subheader("🔎 자료 검색")
        st.caption("참고자료집 문서 안에서 입력한 단어가 포함된 내용을 찾아 보여드립니다. (AI를 사용하지 않는 순수 검색 기능입니다)")

    version_key = _get_reference_version_key()
    reference_texts, unsupported_files = _load_reference_texts(version_key)

    with st.expander(f"📎 현재 인식된 참고자료 ({len([t for t in reference_texts.values() if t])}건)", expanded=False):
        if reference_texts:
            for fname, content in reference_texts.items():
                ok = "✅" if content else "⚠️ 읽기 실패"
                st.markdown(f"- {ok} {fname}")
        else:
            st.markdown("등록된 참고자료가 없습니다. '참고자료집' 메뉴 안내를 참고해 파일을 추가해주세요.")

        if unsupported_files:
            st.warning(
                "다음 파일은 형식이 지원되지 않아 검색되지 않습니다 (PDF/DOCX/TXT로 변환 후 다시 올려주세요): "
                + ", ".join(unsupported_files)
            )

    query = st.text_input(
        "검색어를 입력하세요",
        placeholder="예: 장애인 자녀 고용보험 / 65세 이상 취득 / 계약만료 정산",
    )
    search_clicked = st.button("🔍 검색", type="primary")

    if search_clicked or query:
        if not query.strip():
            st.warning("검색어를 입력해주세요.")
        else:
            results = _search_reference(query, reference_texts, top_k=5)

            if results:
                st.success(f"관련 내용 {len(results)}건을 찾았습니다.")
                for i, r in enumerate(results, start=1):
                    with st.container(border=True):
                        st.markdown(f"**{i}. 출처: 📄 {r['file']}**")
                        st.markdown(_highlight(r["text"], query))
            else:
                st.warning("등록된 참고자료에서 관련 내용을 찾지 못했습니다.")
                import urllib.parse
                search_url = "https://www.google.com/search?q=" + urllib.parse.quote(query)
                st.markdown(
                    f"자료에 없는 내용이라면, 아래 링크로 직접 웹에서 유사 사례를 찾아보실 수 있어요.\n\n"
                    f"🔗 [Google에서 \"{query}\" 검색해보기]({search_url})\n\n"
                    f"※ 이 링크는 자동 검색이 아니라, 클릭하면 새 탭에서 직접 검색 결과를 확인하는 참고용 링크입니다."
                )


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
elif page == "자료 검색":
    render_search_page()
elif page == "참고자료집":
    render_reference_page()
