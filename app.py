import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import sqlite3
import re
import os
import base64

# =============================================================================
# 0. 기본 설정 및 아이콘 / 데이터 / 배경 파일 경로
# =============================================================================
ICON_FILE = "icon.png"
DB_FILE = "genealogy.db"
DATA_FILE = "genealogy.xlsx"

# 저장소 업로드 배경 이미지 파일 탐색 (background.png, bg.png 등)
BG_PATHS = [
    "background.png",
    "background.jpg",
    "bg.png",
    "bg.jpg",
    "artifacts/file_generation/ttl=63d/output/background.png",
    "artifacts/file_generation/ttl=63d/output/bg.png"
]
bg_file = next((p for p in BG_PATHS if os.path.exists(p)), None)

# 폴백 경로 탐색
if not os.path.exists(ICON_FILE):
    alt_icon = "artifacts/file_generation/ttl=63d/output/icon.png"
    if os.path.exists(alt_icon):
        ICON_FILE = alt_icon

DB_PATHS = [
    "genealogy.db",
    "artifacts/file_generation/ttl=63d/output/genealogy.db"
]
db_file_path = next((p for p in DB_PATHS if os.path.exists(p)), DB_FILE)

if not os.path.exists(DATA_FILE):
    if os.path.exists("완귀공파 족보 검색 시스템.xlsx"):
        DATA_FILE = "완귀공파 족보 검색 시스템.xlsx"
    elif os.path.exists("artifacts/file_generation/ttl=63d/output/genealogy.xlsx"):
        DATA_FILE = "artifacts/file_generation/ttl=63d/output/genealogy.xlsx"
    elif os.path.exists("artifacts/file_generation/ttl=63d/output/완귀공파 족보 통합.xlsx"):
        DATA_FILE = "artifacts/file_generation/ttl=63d/output/완귀공파 족보 통합.xlsx"

# ⭐ st.set_page_config는 반드시 모든 st.* 명령보다 가장 먼저 호출되어야 함
st.set_page_config(
    page_title="광주안씨 완귀공파 종중 족보 검색 시스템",
    page_icon=ICON_FILE if os.path.exists(ICON_FILE) else "📜",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ⭐ 이미지 Base64 인코딩을 @st.cache_resource로 메모리에 영구 캐싱하여 모바일 부팅 시 디스크 재읽기 제거
@st.cache_resource
def get_base64_image(image_path):
    if image_path and os.path.exists(image_path):
        try:
            with open(image_path, "rb") as f:
                return f"data:image/png;base64,{base64.b64encode(f.read()).decode()}"
        except Exception:
            return None
    return None

icon_b64 = get_base64_image(ICON_FILE)
bg_b64 = get_base64_image(bg_file) if bg_file else None

def render_html(html_content):
    """
    마크다운 파서가 4칸 이상의 들여쓰기 공백을 코드 블록(<pre><code>)으로 오인하는 현상을
    원천 방지하기 위해 모든 줄의 앞뒤 공백을 제거하고 순수 HTML로 안전하게 렌더링합니다.
    """
    if not html_content:
        return
    clean_lines = [line.strip() for line in str(html_content).splitlines() if line.strip()]
    clean_html = chr(10).join(clean_lines)
    st.markdown(clean_html, unsafe_allow_html=True)

def inject_global_ui_handlers():
    """
    ⭐ 전역 UI 제어 스크립트:
    검색창(인명·계보·단어 검색)에서 엔터(Enter)를 누르는 즉시 모바일 자판(가상 키보드) 강제 숨김(blur)
    """
    components.html(
        """
        <script>
            (function() {
                try {
                    var parentDoc = window.parent.document;

                    // 엔터 키 입력 시 즉시 입력창 포커스 해제하여 모바일 자판 숨김
                    if (!parentDoc._genealogyKeyHandlerAdded) {
                        parentDoc.addEventListener('keydown', function(e) {
                            if (e.key === 'Enter' || e.keyCode === 13) {
                                setTimeout(function() {
                                    try {
                                        if (parentDoc.activeElement && parentDoc.activeElement.tagName === 'INPUT') {
                                            parentDoc.activeElement.blur();
                                        }
                                        var allInputs = parentDoc.querySelectorAll('input[type="text"], input:not([type])');
                                        allInputs.forEach(function(inp) { inp.blur(); });
                                    } catch (err) {}
                                }, 50);
                            }
                        }, true);
                        parentDoc._genealogyKeyHandlerAdded = true;
                    }
                } catch (e) {}
            })();
        </script>
        """,
        height=0
    )

def scroll_to_search_top():
    """
    ⭐ 검색 엔터 후:
    1) 모바일 가상 자판(키보드)이 확실히 내려가도록 포커스 해제(blur) 처리
    2) 인명검색 라인이 화면 제일 위에 오도록 상단으로 자동 스크롤
    """
    components.html(
        """
        <script>
            (function() {
                function dismissKeyboardAndScroll() {
                    try {
                        var parentDoc = window.parent.document;
                        var allInputs = parentDoc.querySelectorAll('input');
                        allInputs.forEach(function(inp) {
                            inp.blur();
                        });
                        if (parentDoc.activeElement && typeof parentDoc.activeElement.blur === 'function') {
                            parentDoc.activeElement.blur();
                        }
                        var target = parentDoc.getElementById('search-section-anchor');
                        if (target) {
                            target.scrollIntoView({ behavior: 'smooth', block: 'start' });
                        }
                    } catch (e) {}
                }
                setTimeout(dismissKeyboardAndScroll, 60);
                setTimeout(dismissKeyboardAndScroll, 220);
            })();
        </script>
        """,
        height=0
    )

def get_query_param(key):
    try:
        if hasattr(st, "query_params") and key in st.query_params:
            return st.query_params[key]
        elif hasattr(st, "experimental_get_query_params"):
            params = st.experimental_get_query_params()
            if key in params and params[key]:
                return params[key][0]
    except Exception:
        pass
    return None

# =============================================================================
# 1. 디자인 스타일 CSS
# =============================================================================
CUSTOM_CSS = """
<style>
html, body, [class*="css"], .stMarkdown, p, div {
    font-family: 'HCR Dotum', '함초롱돋움', 'Apple SD Gothic Neo', 'Malgun Gothic', 'Noto Sans KR', sans-serif;
}

/* 기본 상단 헤더 숨김 */
header[data-testid="stHeader"] {
    display: none !important;
}

/* ⭐ 화면에 있는 모든 창의 바깥쪽 좌우 여백을 기존 대비 50% 축소 (0.65rem -> 0.32rem) */
.main .block-container, 
[data-testid="stAppViewBlockContainer"], 
.block-container {
    max-width: 100% !important;
    width: 100% !important;
    margin-left: auto !important;
    margin-right: auto !important;
    padding-top: 1.0rem !important;
    padding-bottom: 1.2rem !important;
    padding-left: 0.32rem !important;
    padding-right: 0.32rem !important;
    box-sizing: border-box !important;
}

div[data-testid="stVerticalBlock"] {
    gap: 0.28rem !important;
}
div[data-testid="stElementContainer"],
.element-container {
    margin-top: 0px !important;
    margin-bottom: 1px !important;
}

/* 저장소 배경 이미지 적용 및 투명도 제로(100% 완전 불투명) */
.stApp {
    background-color: #162b52;
    background-size: cover !important;
    background-position: center top !important;
    background-attachment: fixed !important;
    opacity: 1 !important;
}

/* ⭐ 제일 위 페이지 타이틀 창: 배경 색상 황색(#facc15) + 아이콘을 왼쪽 끝으로 최대한 밀착 */
.title-container {
    width: 100% !important;
    box-sizing: border-box !important;
    margin-top: 8px !important;
    margin-bottom: 14px !important;
    padding: 8px 10px 8px 2px !important;
    border: 2px solid #b45309 !important;
    border-bottom: 3.5px solid #8c2d19 !important;
    background: #facc15 !important;
    background-color: #facc15 !important;
    opacity: 1 !important;
    border-radius: 6px;
    box-shadow: 0 3px 8px rgba(0,0,0,0.25);
    display: flex !important;
    flex-direction: row !important;
    align-items: center !important;
    justify-content: flex-start !important;
    gap: 6px !important;
}

.title-left-icon {
    flex-shrink: 0 !important;
    display: flex !important;
    align-items: center !important;
    justify-content: flex-start !important;
    margin-left: 0px !important;
    padding-left: 0px !important;
}

.title-right-content {
    flex: 1 !important;
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    text-align: center !important;
}

.title-clan-name {
    font-size: 13.5pt;
    font-weight: 800;
    color: #1e3c72;
    line-height: 1.25;
    letter-spacing: -0.3px;
    margin-bottom: 2px;
}

.title-main-center {
    font-size: 16.5pt;
    font-weight: 800;
    color: #111827;
    text-align: center;
    width: 100%;
    margin-top: 1px;
    margin-bottom: 3px;
    line-height: 1.25;
    letter-spacing: -0.5px;
}

.title-sub-bottom {
    font-size: 10.5pt;
    font-weight: 700;
    color: #1e293b;
    text-align: center;
    width: 100%;
    margin-top: 1px;
    line-height: 1.35;
    letter-spacing: -0.2px;
}

/* 위 타이틀 창과 간격을 넓힌 안내 문구 라인 */
.compact-search-header {
    width: 100% !important;
    text-align: center !important;
    margin-top: 12px !important;
    margin-bottom: 2px !important;
    padding: 2px 2px !important;
    line-height: 1.35 !important;
}
.compact-search-subtitle {
    font-size: 10.8pt !important;
    font-weight: 700 !important;
    color: #fef08a !important;
    text-align: center !important;
    margin-top: 0px !important;
    margin-bottom: 0px !important;
    line-height: 1.35 !important;
    letter-spacing: -0.3px !important;
    text-shadow: 0 1px 3px rgba(0,0,0,0.65);
}

/* 검색 유형 선택 라인과 아래 검색어 입력창 간격을 50% 이상 축소 */
div[data-testid="stRadio"] {
    margin-top: 0px !important;
    margin-bottom: 0px !important;
    padding-top: 0px !important;
    padding-bottom: 0px !important;
}
div[role="radiogroup"] {
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: nowrap !important;
    align-items: center !important;
    justify-content: center !important;
    gap: 14px !important;
    width: 100% !important;
    margin-top: 0px !important;
    margin-bottom: 0px !important;
    padding: 0px 2px !important;
    overflow-x: auto !important;
}
div[role="radiogroup"] > label,
div[role="radiogroup"] label[data-baseweb="radio"] {
    white-space: nowrap !important;
    flex-shrink: 0 !important;
    margin-right: 2px !important;
    margin-bottom: 0px !important;
    padding: 1px 2px !important;
    line-height: 1.25 !important;
}
div[role="radiogroup"] label p,
div[role="radiogroup"] label span {
    font-size: 12.5pt !important;
    font-weight: 800 !important;
    color: #ffffff !important;
    white-space: nowrap !important;
    line-height: 1.25 !important;
    letter-spacing: -0.3px !important;
    text-shadow: 0 1px 3px rgba(0,0,0,0.65);
}

/* 검색어 입력창: 테두리 흰색(3.5px solid #ffffff), 배경 청색, 세로 중앙 정렬 */
div[data-testid="stTextInput"] {
    margin-top: 0px !important;
    margin-bottom: 4px !important;
    padding-top: 0px !important;
    padding-bottom: 0px !important;
}
div[data-testid="stTextInputRootElement"] {
    border: 3.5px solid #ffffff !important;
    border-radius: 8px !important;
    background-color: #1d4ed8 !important;
    box-shadow: 0 2px 6px rgba(0,0,0,0.3) !important;
    box-sizing: border-box !important;
    overflow: hidden !important;
}
div[data-testid="stTextInput"] div[data-baseweb="input"],
div[data-testid="stTextInput"] div[data-baseweb="base-input"] {
    background-color: #1d4ed8 !important;
    border: none !important;
    box-shadow: none !important;
    height: 46px !important;
    min-height: 46px !important;
    display: flex !important;
    align-items: center !important;
}
div[data-testid="stTextInput"] input {
    background-color: #1d4ed8 !important;
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
    caret-color: #ffffff !important;
    font-weight: 800 !important;
    font-size: 13.5pt !important;
    padding: 0px 12px !important;
    height: 100% !important;
    line-height: 46px !important;
    vertical-align: middle !important;
}
div[data-testid="stTextInput"] input::placeholder {
    color: #cbd5e1 !important;
    -webkit-text-fill-color: #cbd5e1 !important;
    font-size: 9.6pt !important;
    font-weight: 400 !important;
    opacity: 1 !important;
    line-height: 46px !important;
    vertical-align: middle !important;
}
div[data-testid="stTextInput"] input::-webkit-input-placeholder {
    color: #cbd5e1 !important;
    -webkit-text-fill-color: #cbd5e1 !important;
    font-size: 9.6pt !important;
    font-weight: 400 !important;
    opacity: 1 !important;
    line-height: 46px !important;
    vertical-align: middle !important;
}

/* 검색 결과 없을 때 안내 문구: 글자 색상 흰색(#ffffff) */
.no-result-notice {
    width: 100% !important;
    box-sizing: border-box !important;
    background-color: rgba(30, 58, 138, 0.88) !important;
    border: 1.5px solid #ffffff !important;
    color: #ffffff !important;
    padding: 9px 12px !important;
    border-radius: 6px !important;
    font-size: 10.8pt !important;
    font-weight: 500 !important;
    margin-top: 4px !important;
    margin-bottom: 6px !important;
    line-height: 1.45 !important;
    text-align: center !important;
    text-shadow: 0 1px 2px rgba(0,0,0,0.6) !important;
}

/* ⭐ '초기 검색 인물로 복귀' 버튼: 바탕색 녹색(#15803d), 글씨체 흰색(#ffffff) */
div[data-testid="stButton"] > button {
    background-color: #15803d !important;
    background: #15803d !important;
    color: #ffffff !important;
    border: 2px solid #ffffff !important;
    border-radius: 6px !important;
    font-weight: 800 !important;
    font-size: 11pt !important;
    padding: 6px 12px !important;
    box-shadow: 0 2px 5px rgba(0,0,0,0.3) !important;
}
div[data-testid="stButton"] > button:hover,
div[data-testid="stButton"] > button:active,
div[data-testid="stButton"] > button:focus {
    background-color: #166534 !important;
    background: #166534 !important;
    color: #ffffff !important;
    border-color: #fef08a !important;
}
div[data-testid="stButton"] > button *,
div[data-testid="stButton"] > button p,
div[data-testid="stButton"] > button span {
    color: #ffffff !important;
    font-weight: 800 !important;
    font-size: 11pt !important;
}

/* 복수 후보 안내 배너 */
.multi-ancestor-banner {
    width: 100% !important;
    box-sizing: border-box !important;
    background-color: #eff6ff !important;
    border: 1.5px solid #3b82f6 !important;
    color: #1d4ed8 !important;
    padding: 7px 10px;
    border-radius: 6px;
    font-size: 11pt !important;
    font-weight: 400 !important;
    margin-top: 3px !important;
    margin-bottom: 5px !important;
    line-height: 1.45 !important;
}

/* 드롭다운 명부 */
div[data-testid="stSelectbox"] {
    margin-top: 2px !important;
    margin-bottom: 5px !important;
}
div[data-testid="stSelectbox"] *,
div[data-testid="stSelectbox"] div[data-baseweb="select"],
div[data-testid="stSelectbox"] div[data-baseweb="select"] div,
div[data-testid="stSelectbox"] div[data-baseweb="select"] span {
    font-size: 11.5pt !important;
    font-weight: 800 !important;
    color: #0f172a !important;
    line-height: 1.4 !important;
}
ul[data-testid="stSelectboxVirtualDropdown"] li,
ul[data-testid="stSelectboxVirtualDropdown"] li *,
div[role="listbox"] li,
div[role="listbox"] li *,
div[role="option"],
div[role="option"] * {
    font-size: 11.5pt !important;
    font-weight: 800 !important;
    color: #0f172a !important;
    line-height: 1.45 !important;
}

/* 검색 결과 인물 카드 */
.result-card {
    width: 100% !important;
    box-sizing: border-box !important;
    background-color: #ffffff;
    border: 2px solid #1e3c72;
    border-top: 4px solid #1e3c72;
    border-radius: 8px;
    padding: 10px 12px;
    margin-top: 4px !important;
    margin-bottom: 8px !important;
    box-shadow: 0 4px 12px rgba(0,0,0,0.2);
}
.result-card-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 2px solid #e2e8f0;
    padding-bottom: 5px;
    margin-bottom: 7px;
    line-height: 1.35;
}
.result-name-big {
    font-size: 15.5pt;
    font-weight: 800;
    color: #000000;
    display: flex;
    align-items: center;
    gap: 8px;
}
.result-badge {
    font-size: 11.5pt;
    font-weight: 800;
    background: #eff6ff;
    color: #1d4ed8;
    padding: 2px 8px;
    border-radius: 8px;
    border: 1px solid #bfdbfe;
}
/* ⭐ 모든 족보 페이지 표시: '시직공파보' 줄바꿈 '1권 433p 5단' 형식 + 글자크기 1p 축소 (10.5pt -> 9.5pt) */
.result-book-loc {
    font-size: 9.5pt !important;
    font-weight: 800 !important;
    color: #1e293b !important;
    background: #f1f5f9 !important;
    padding: 2px 8px !important;
    border-radius: 6px !important;
    border: 1px solid #cbd5e1 !important;
    text-align: right !important;
    line-height: 1.22 !important;
    flex-shrink: 0 !important;
}

/* 한글 약력 컨테이너 */
.meta-para-container {
    width: 100% !important;
    box-sizing: border-box !important;
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 8px 11px;
    margin-bottom: 8px;
    line-height: 1.52 !important;
    font-size: 10.5pt;
    color: #0f172a;
}
.meta-para-item {
    margin-bottom: 4px;
    line-height: 1.52 !important;
    word-break: keep-all;
    color: #0f172a;
}
.meta-para-label {
    font-weight: 700;
    color: #1e3c72;
    margin-right: 4px;
}
.meta-para-value {
    color: #1e293b;
}

/* 한문 약력 원문 전용 컨테이너 */
.hanmun-raw-container {
    width: 100% !important;
    box-sizing: border-box !important;
    background-color: #fcfcfd;
    border: 1px solid #cbd5e1;
    border-left: 4px solid #1e3c72;
    border-radius: 6px;
    padding: 8px 12px;
    margin-top: 2px !important;
    margin-bottom: 2px !important;
    line-height: 1.55 !important;
    font-size: 10.5pt;
    color: #1e293b;
}

.red-correction {
    color: #dc2626 !important;
    font-weight: 700;
    background-color: #fef2f2;
    padding: 0px 3px;
    border-radius: 3px;
    border-bottom: 1px dashed #f87171;
}

.search-word-red {
    color: #ff0000 !important;
    font-weight: 700 !important;
    background-color: #fee2e2;
    padding: 0px 4px;
    border-radius: 3px;
    border: 1px solid #f87171;
}

/* ============================================================================
   ⭐ 계보도 검색 결과창 세대별 큰 박스 및 좌·우 2영역 분할, 균일 가로 길이 스타일
   1) 인명 테두리는 한 겹만 표시 (내부 <span> 등의 중첩 테두리 원천 제거)
   2) 인명 박스를 좌·우 2개 영역으로 나누어 좌영역에는 남형제, 우영역에는 여형제(사위) 배치
   3) 모든 함자 박스의 가로 길이를 일정하게 통일 (width: 106px)
   4) 본인 창 본인 함자 박스: 바탕색 매우 진한 황토색(#7c2d12), 글자색 흰색(#ffffff), 기준 표시 삭제
   ============================================================================ */
.tree-diagram-wrapper {
    width: 100% !important;
    box-sizing: border-box !important;
    background: #ffffff !important;
    border: 2px solid #1e3c72 !important;
    border-radius: 8px !important;
    padding: 8px 10px !important;
    margin-top: 2px !important;
    margin-bottom: 6px !important;
}
.tree-header-bar {
    display: flex !important;
    align-items: center !important;
    justify-content: space-between !important;
    border-bottom: 2px solid #e2e8f0 !important;
    padding-bottom: 5px !important;
    margin-bottom: 6px !important;
    gap: 8px !important;
}
/* 세대별 큰 박스: 바탕색 흰색(#ffffff) */
.tree-level-row {
    display: flex !important;
    align-items: stretch !important;
    margin-bottom: 0px !important;
    padding: 4px 6px !important;
    background: #ffffff !important;
    background-color: #ffffff !important;
    border-radius: 5px !important;
    border: 1px solid #cbd5e1 !important;
    line-height: 1.3 !important;
}
/* 본인 창(세대 큰 박스): 바탕색 흰색(#ffffff) */
.tree-level-row-self {
    display: flex !important;
    align-items: stretch !important;
    margin-bottom: 0px !important;
    padding: 4px 6px !important;
    background: #ffffff !important;
    background-color: #ffffff !important;
    border-radius: 5px !important;
    border: 1.5px solid #7c2d12 !important;
    line-height: 1.3 !important;
}
/* 세대와 세대를 연결하는 화살표의 길이를 절반으로 축소 */
.tree-connector-line {
    text-align: center;
    color: #1e3c72;
    font-weight: 900;
    font-size: 8.5pt;
    line-height: 0.85 !important;
    margin: -1px 0px !important;
    padding: 0px !important;
}
.tree-gen-badge {
    min-width: 72px;
    font-weight: 700;
    color: #1e3c72;
    font-size: 9.2pt;
    padding: 2px 4px;
    background: #dbeafe;
    border-radius: 4px;
    text-align: center;
    margin-right: 6px;
    line-height: 1.2;
    flex-shrink: 0;
    display: flex;
    align-items: center;
    justify-content: center;
}
/* 본인 창의 좌측 '본인 00세' 배지 */
.tree-gen-badge-self {
    min-width: 72px;
    font-weight: 800 !important;
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
    font-size: 9.2pt;
    padding: 2px 4px;
    background: #7c2d12 !important;
    background-color: #7c2d12 !important;
    border-radius: 4px;
    text-align: center;
    margin-right: 6px;
    line-height: 1.2;
    flex-shrink: 0;
    display: flex;
    align-items: center;
    justify-content: center;
}

/* ⭐ 인명 박스를 좌·우 2개 영역으로 나누는 컨테이너 (좌영역: 남형제, 우영역: 여형제/사위) */
.tree-split-wrapper {
    display: flex !important;
    flex-direction: row !important;
    align-items: flex-start !important;
    gap: 6px !important;
    flex: 1 !important;
    width: 100% !important;
    background: #ffffff !important;
}
.tree-left-male-zone {
    flex: 1 1 50% !important;
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: wrap !important;
    gap: 4px !important;
    align-items: center !important;
    align-content: flex-start !important;
    min-height: 24px !important;
}
.tree-right-female-zone {
    flex: 1 1 50% !important;
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: wrap !important;
    gap: 4px !important;
    align-items: center !important;
    align-content: flex-start !important;
    border-left: 1px dashed #cbd5e1 !important;
    padding-left: 6px !important;
    min-height: 24px !important;
}

/* ⭐ 개별 함자 박스:
   - 테두리는 오직 바깥 <a> 태그에 한 겹만 적용 (내부 요소 border: none)
   - 가로 길이를 모두 일정하게 고정 (width: 106px) */
.tree-node-box {
    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    text-align: center !important;
    text-decoration: none !important;
    cursor: pointer !important;
    width: 106px !important;
    min-width: 106px !important;
    max-width: 106px !important;
    box-sizing: border-box !important;
    border-radius: 4px !important;
    padding: 3px 3px !important;
    font-size: 8.8pt !important;
    line-height: 1.25 !important;
    font-weight: 700 !important;
    white-space: nowrap !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
    box-shadow: none !important;
}
/* 내부 요소에는 절대로 테두리가 생기지 않도록 강제 (한 겹 테두리 원칙) */
.tree-node-box * {
    border: none !important;
    background: transparent !important;
    box-shadow: none !important;
    padding: 0 !important;
    margin: 0 !important;
}

/* 1) 직계 선조 함자 박스: 황색(#fde047) 바탕에 검은색(#000000) 글자, 한 겹 테두리 */
a.tree-node-direct {
    background-color: #fde047 !important;
    background: #fde047 !important;
    border: 1px solid #ca8a04 !important;
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    font-weight: 800 !important;
}
/* 2) 본인 창의 본인 함자 박스: 매우 진한 황토색(#7c2d12) 바탕에 흰색(#ffffff) 글자, 한 겹 테두리 */
a.tree-node-self {
    background-color: #7c2d12 !important;
    background: #7c2d12 !important;
    border: 1px solid #431407 !important;
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
    font-weight: 800 !important;
}
/* 3) 남형제 함자 박스: 옅은 청색(#dbeafe) 바탕에 검은색(#000000) 글자, 한 겹 테두리 */
a.tree-node-male {
    background-color: #dbeafe !important;
    background: #dbeafe !important;
    border: 1px solid #60a5fa !important;
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    font-weight: 700 !important;
}
/* 4) 여형제(사위) 함자 박스: 흰색(#ffffff) 바탕에 검은색(#000000) 글자, 한 겹 테두리 */
a.tree-node-female {
    background-color: #ffffff !important;
    background: #ffffff !important;
    border: 1px solid #94a3b8 !important;
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    font-weight: 700 !important;
}

/* 종중 주요 유적 및 문화재 안내: 바탕색 녹색(#15803d), 헤더 흰색 글씨 */
.heritage-sub-container {
    width: 100% !important;
    box-sizing: border-box !important;
    background-color: #15803d !important;
    border: 2px solid #14532d !important;
    border-radius: 8px;
    padding: 9px 12px;
    margin-top: 6px;
    margin-bottom: 6px;
    box-shadow: 0 3px 8px rgba(0,0,0,0.22);
}
.heritage-sub-header {
    font-size: 11.2pt;
    font-weight: 800;
    color: #ffffff !important;
    margin-bottom: 6px;
    line-height: 1.4;
    text-shadow: 0 1px 2px rgba(0,0,0,0.5);
}
.heritage-sub-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 6px;
}
.heritage-sub-card {
    background: #ffffff;
    border: 1px solid #dcfce7;
    border-radius: 6px;
    padding: 6px 8px;
    text-decoration: none !important;
    display: block;
}
.heritage-sub-card:hover {
    background: #f0fdf4;
    border-color: #fef08a;
}
.h-card-name {
    font-size: 10pt;
    font-weight: 700;
    color: #14532d;
    margin-bottom: 2px;
    line-height: 1.35;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.h-card-desc {
    font-size: 8.8pt;
    color: #334155;
    line-height: 1.35;
}

/* ⭐ 모바일에서도 모든 창의 바깥쪽 좌우 여백을 50% 축소 (12px -> 6px) 및 좌우 대칭 유지 */
@media (max-width: 768px) {
    html,
    body,
    #root {
        width: 100% !important;
        max-width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
        box-sizing: border-box !important;
    }

    .stApp,
    [data-testid="stAppViewContainer"],
    section.main,
    [data-testid="stAppViewBlockContainer"],
    [data-testid="stMain"] {
        width: 100% !important;
        max-width: 100% !important;
        margin-left: 0 !important;
        margin-right: 0 !important;
        padding-left: 0 !important;
        padding-right: 0 !important;
        box-sizing: border-box !important;
    }

    [data-testid="stMainBlockContainer"],
    .main .block-container,
    .block-container {
        position: relative !important;
        left: calc((100vw - 100%) / 2) !important;
        width: 100% !important;
        max-width: 100% !important;
        margin-left: auto !important;
        margin-right: auto !important;
        padding-left: 6px !important;
        padding-right: 6px !important;
        box-sizing: border-box !important;
    }

    [data-testid="stMarkdownContainer"],
    .stMarkdown {
        width: 100% !important;
        max-width: 100% !important;
        box-sizing: border-box !important;
    }
}
</style>
"""

render_html(CUSTOM_CSS)
inject_global_ui_handlers()

# =============================================================================
# 2. 데이터 로드 및 초고속(0.05초) 메모리 인덱싱 (@st.cache_resource 영구 공유)
# =============================================================================
VOL_ORDER_MAP = {"A": 1, "B": 2, "C": 3, "D": 4}

KO_SEARCH_PATTERNS = [re.compile(p) for p in [
    r'자는\s*([^\.,\n\s]+)',
    r'초자는?\s*([^\.,\n\s]+)',
    r'호는\s*([^\.,\n\s]+)',
    r'일휘는?\s*([^\.,\n\s]+)',
    r'초휘는?\s*([^\.,\n\s]+)',
    r'초명은?\s*([^\.,\n\s]+)',
    r'통명은?\s*([^\.,\n\s]+)',
]]

CN_SEARCH_PATTERNS = [re.compile(p) for p in [
    r'字\s*([^\.,\n\s]+)',
    r'初字\s*([^\.,\n\s]+)',
    r'號\s*([^\.,\n\s]+)',
    r'一諱\s*([^\.,\n\s]+)',
    r'初諱\s*([^\.,\n\s]+)',
    r'初名\s*([^\.,\n\s]+)',
    r'通名\s*([^\.,\n\s]+)',
]]

def parse_gid_tuple(gid):
    """고유ID(예: C-529-4-1-37)를 (권, 페이지, 단, 순서, 세수)로 분리"""
    m = re.match(r"^([A-Z])-(\d+)-(\d+)-(\d+)-(\d+)$", str(gid).strip())
    if m:
        return m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4)), int(m.group(5))
    return "", 0, 0, 0, 0

def format_book_location_html(gid):
    """
    ⭐ 모든 족보 페이지 표시 방법:
    '시직공파보' 줄바꿈하고 '1권 433p 5단' 과 같은 형식
    - A, B, C로 시작하면 '1권', D로 시작하면 '2권'
    """
    vol, page, tier, _, _ = parse_gid_tuple(gid)
    if not vol:
        return str(gid).strip()
    vol_num = "1권" if vol in ("A", "B", "C") else ("2권" if vol == "D" else f"{vol}권")
    return f"시직공파보<br>{vol_num} {page}p {tier}단"

def get_sibling_sort_key(row, all_sibling_rels=None):
    """
    ⭐ 동일 부친을 둔 형제·자매(자녀) 정렬 키:
    1) 남성 우위 원칙: 남성(1자, 2자, 3자..., 계자, 승사손)이 여성(1녀, 2녀, 3녀...)보다 무조건 우선
    2) 동성 내에서는 장유 서열(1자 < 2자 < 3자... / 1녀 < 2녀 < 3녀...) 우선
       - '계자'/'승사손'은 형제 중 '1자'가 없으면 장자(1) 위치에, '1자'가 있으면 '1자' 바로 다음(1.5)에 배치
    3) 앞선 권(A<B<C<D) -> 앞선 페이지(page) -> 앞선 단(tier) -> 앞선 순서(seq) 우선
    """
    rel = str(row.get("관계", "")).strip()
    vol, page, tier, seq, _ = parse_gid_tuple(row.get("고유ID", ""))
    vol_idx = VOL_ORDER_MAP.get(vol, 9)

    if "녀" in rel:
        gender_rank = 1
        m = re.match(r"^(\d+)녀$", rel)
        rel_num = float(m.group(1)) if m else 99.0
    else:
        gender_rank = 0
        m = re.match(r"^(\d+)자$", rel)
        if m:
            rel_num = float(m.group(1))
        elif rel in ("계자", "승사손"):
            rel_num = 1.5 if (all_sibling_rels and "1자" in all_sibling_rels) else 1.0
        else:
            rel_num = 99.0

    return (gender_rank, rel_num, vol_idx, page, tier, seq)

def sort_siblings_df(sib_df):
    if sib_df is None or len(sib_df) <= 1:
        return sib_df
    rels = set(sib_df["관계"].astype(str).str.strip().tolist())
    recs = sib_df.to_dict("records")
    recs.sort(key=lambda r: get_sibling_sort_key(r, rels))
    return pd.DataFrame(recs)

# ⭐ @st.cache_resource를 사용하여 서버 메모리에 단 1회만 올리고 모든 모바일 접속자가 즉시(0.00초) 공유
@st.cache_resource
def load_and_preprocess_genealogy():
    raw_df = None
    status_msg = ""

    # 1순위: SQLite DB 로드
    if os.path.exists(db_file_path):
        try:
            conn = sqlite3.connect(db_file_path)
            raw_df = pd.read_sql("SELECT * FROM genealogy", conn).fillna("")
            conn.close()
            if not raw_df.empty:
                status_msg = "성공 (SQLite DB)"
        except Exception:
            raw_df = None

    # 2순위: Excel 로드
    if (raw_df is None or raw_df.empty) and os.path.exists(DATA_FILE):
        try:
            wb = openpyxl.load_workbook(DATA_FILE, read_only=True)
            sheet_name = "족보전사결과" if "족보전사결과" in wb.sheetnames else wb.sheetnames[0]
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return None, None, None, "시트에 데이터가 없습니다."
            header = [str(c).strip() if c is not None else f"Col{idx}" for idx, c in enumerate(rows[0])]
            raw_df = pd.DataFrame(rows[1:], columns=header).fillna("")
            status_msg = "성공 (Excel)"
        except Exception as e:
            return None, None, None, str(e)

    if raw_df is None or raw_df.empty:
        return None, None, None, f"데이터 파일({DATA_FILE})을 찾을 수 없습니다."

    # 1. 초고속 딕셔너리 인덱싱(O(N))으로 동명이인 부친 고유ID(_father_gid) 및 통합 검색 인덱스(_search_blob) 생성
    records = raw_df.to_dict("records")
    by_name = {}
    search_blobs = []

    for r in records:
        vol, page, tier, seq, gen = parse_gid_tuple(r.get("고유ID", ""))
        r["_vol"] = vol
        r["_page"] = page
        r["_tier"] = tier
        r["_seq"] = seq
        r["_gen"] = gen
        nm = str(r.get("인명", "")).strip()
        by_name.setdefault(nm, []).append(r)

        # 인명·자·초자·호·일휘·초휘·초명·통명 사전 인덱싱 (검색 시 정규식 재실행 제거)
        ko_name = re.sub(r'[^가-힣]+', '', nm)
        cn_name = re.sub(r'[^一-龥]+', '', nm)
        tokens = [nm, ko_name, cn_name]
        if ko_name and cn_name:
            tokens.append(f"{ko_name}_{cn_name}")
            tokens.append(f"{ko_name}{cn_name}")
        if ko_name:
            tokens.append('안' + ko_name)
        if cn_name:
            tokens.append('安' + cn_name)

        fl_ko = str(r.get('한글 약력 정리', '')).split('\n')[0]
        for pat in KO_SEARCH_PATTERNS:
            for m in pat.finditer(fl_ko):
                tokens.append(m.group(1).strip())

        fl_cn = str(r.get('한문 약력 정리', '')).split('\n')[0]
        for pat in CN_SEARCH_PATTERNS:
            for m in pat.finditer(fl_cn):
                tokens.append(m.group(1).strip())

        search_blobs.append("\n".join(t for t in tokens if t))

    father_gids = []
    for r in records:
        f_name = str(r.get("부친", "")).strip()
        if not f_name or f_name == "覯구" or f_name == "-":
            r["_father_gid"] = ""
            father_gids.append("")
            continue
        cands_all = by_name.get(f_name, [])
        if not cands_all:
            r["_father_gid"] = ""
            father_gids.append("")
            continue
        if len(cands_all) == 1:
            fg = str(cands_all[0]["고유ID"])
            r["_father_gid"] = fg
            father_gids.append(fg)
            continue

        # 1차: 세수(본인 세수 - 1) 일치 후보
        c_gen = r["_gen"]
        cands = [c for c in cands_all if c["_gen"] == c_gen - 1] or cands_all
        if len(cands) == 1:
            fg = str(cands[0]["고유ID"])
            r["_father_gid"] = fg
            father_gids.append(fg)
            continue

        # 2차: 후보의 부친 == 본인의 조부 일치 후보
        gf_name = str(r.get("조부", "")).strip()
        if gf_name and gf_name != "-":
            cands_gf = [c for c in cands if str(c.get("부친", "")).strip() == gf_name]
            if len(cands_gf) == 1:
                fg = str(cands_gf[0]["고유ID"])
                r["_father_gid"] = fg
                father_gids.append(fg)
                continue
            elif len(cands_gf) > 1:
                cands = cands_gf

        # 3차: 같은 권(Volume)에서 페이지가 작거나 같으면서 가장 가까운 선조 선택
        c_vol, c_page = r["_vol"], r["_page"]
        same_vol = [c for c in cands if c["_vol"] == c_vol and c["_page"] <= c_page]
        if same_vol:
            same_vol.sort(key=lambda x: (-x["_page"], x["_seq"]))
            fg = str(same_vol[0]["고유ID"])
            r["_father_gid"] = fg
            father_gids.append(fg)
            continue

        # 4차: 생년/배위 상세 기록이 있는 본 기록 우선
        non_empty = [c for c in cands if ("생" in str(c.get("한글 약력 정리", "")) or "배 :" in str(c.get("한글 약력 정리", "")))]
        if len(non_empty) == 1:
            fg = str(non_empty[0]["고유ID"])
            r["_father_gid"] = fg
            father_gids.append(fg)
            continue

        fg = str(cands[0]["고유ID"])
        r["_father_gid"] = fg
        father_gids.append(fg)

    raw_df["_father_gid"] = father_gids
    raw_df["_search_blob"] = search_blobs

    # 2. 계보도 초고속(0.001초) 렌더링을 위한 O(1) 부모-자녀 해시맵 및 사전 정렬
    gid_to_row = {r["고유ID"]: r for r in records}
    children_by_fgid = {}
    for r in records:
        fg = r["_father_gid"]
        if fg:
            children_by_fgid.setdefault(fg, []).append(r)
        elif str(r.get("부친", "")).strip() == "覯구":
            children_by_fgid.setdefault("__ROOT_GU__", []).append(r)

    for fg, ch_list in children_by_fgid.items():
        if len(ch_list) > 1:
            rels = {str(x.get("관계", "")).strip() for x in ch_list}
            ch_list.sort(key=lambda x: get_sibling_sort_key(x, rels))

    return raw_df, gid_to_row, children_by_fgid, status_msg

df, gid_to_row, children_by_fgid, load_msg = load_and_preprocess_genealogy()

# =============================================================================
# 3. 보조 유틸리티 함수: '한글_한문' 표기, 계자/출계 표기, 문장 끝 마침표 처리
# =============================================================================
# 1) 과거 관련 통합 검색어 그룹 (1단어만 입력해도 그룹 전체 검색)
EXAM_GROUP_KEYWORDS = [
    "과거", "문과", "급제", "사마", "별시", "대과", "무과", "잡과", "장원",
    "진사", "생원", "향시", "회시", "복시", "현랑과", "현량과", "증광시",
    "전시", "식년시", "알성시", "정시", "춘당대시", "갑과", "을과", "병과"
]

# 2) 문집·저술 관련 통합 검색어 그룹 (문집, 유집, 유고, 실기, 서집록, 저서, 논문 중 1단어만 입력해도 전체 통합 검색)
LITERARY_GROUP_KEYWORDS = [
    "문집", "유집", "유고", "실기", "서집록", "저서", "논문"
]

def sort_search_candidates_df(cand_df):
    """
    인명·계보·단어 검색 결과 후보 명부도
    1) 남성 우위 원칙 (남성 선조 우선, 여성 후순위)
    2) 앞선 권(상권 A·B·C -> 하권 D) 우선
    3) 앞선 페이지(page) 우선
    4) 앞선 단(tier) 및 순서(seq) 우선
    원칙에 따라 정렬
    """
    if cand_df is None or len(cand_df) <= 1:
        return cand_df
    recs = cand_df.to_dict("records")
    def cand_key(r):
        rel = str(r.get("관계", "")).strip()
        gender_rank = 1 if "녀" in rel else 0
        vol, page, tier, seq, _ = parse_gid_tuple(r.get("고유ID", ""))
        vol_idx = VOL_ORDER_MAP.get(vol, 9)
        return (gender_rank, vol_idx, page, tier, seq)
    recs.sort(key=cand_key)
    return pd.DataFrame(recs)

def fast_search_by_name(all_df, q):
    """
    사전 구축된 _search_blob 컬럼을 활용하여 2,887명 전체에서 함자·성명·자·초자·호·일휘·초휘·초명·통명을
    0.004초 만에 초고속 벡터 검색
    """
    q = q.strip()
    if not q:
        return all_df.iloc[0:0]
    q_alt = q[1:] if q.startswith('안') and len(q) > 1 else q
    pat = re.escape(q) if q == q_alt else f"{re.escape(q)}|{re.escape(q_alt)}"
    mask = all_df["_search_blob"].str.contains(pat, na=False, regex=True)
    return sort_search_candidates_df(all_df[mask].reset_index(drop=True))

def extract_generation_info(gid):
    m = re.search(r"-(\d+)$", str(gid).strip())
    if m:
        gen = int(m.group(1))
        wangwi_desc_gen = gen - 24
        desc_str = "완귀공 본인" if wangwi_desc_gen == 0 else f"완귀공의 {wangwi_desc_gen}대손"
        return gen, f"{gen}세  {desc_str}"
    return None, "세수 미상"

def split_name(name_raw):
    s = str(name_raw).strip()
    m = re.match(r"^([一-龥]+)?([가-힣]+)?$", s)
    if m and m.group(1) and m.group(2):
        return m.group(2), m.group(1)
    if re.fullmatch(r"^[가-힣]+$", s):
        return s, ""
    if re.fullmatch(r"^[一-龥]+$", s):
        return "", s
    return s, ""

def clean_ko_name(raw_val):
    s = str(raw_val).strip()
    ko, _ = split_name(s)
    if ko:
        return ko
    cleaned = re.sub(r"[一-龥]+", "", s).strip()
    return cleaned if cleaned else s

def format_ko_hanja_attached(raw_val):
    """
    ⭐ 모든 창의 한글 한문 표시 방식: '한글_한문' (예: '嶒증' -> '증_嶒', '禧洙희수' -> '희수_禧洙')
    """
    s = str(raw_val).strip()
    ko, hj = split_name(s)
    if ko and hj:
        return f"{ko}_{hj}"
    return s

def get_adoption_info(row):
    rel = str(row.get("관계", "")).strip()
    ko_bio = str(row.get("한글 약력 정리", "")).strip()
    ko_raw = str(row.get("한글 약력 원문", "")).strip()
    cn_raw = str(row.get("한문 약력 원문", "")).strip()
    father_ko = clean_ko_name(row.get("부친", ""))

    if rel == "계자" or "생부 " in ko_bio:
        m = re.search(r"생부\s*([^\s의\.]+)의?", ko_bio)
        saengbu_ko = clean_ko_name(m.group(1)) if m else ""
        return "입계", father_ko, saengbu_ko

    if "양부" in ko_bio or "출계" in ko_raw or "出系" in cn_raw:
        m = re.search(r"양부\s*(.+?)의\s*계자", ko_bio)
        yangbu_ko = clean_ko_name(m.group(1)) if (m and m.group(1).strip()) else ""
        if not yangbu_ko:
            m2 = re.search(r"([^\s]+)의\s*후사로\s*출계|출계\s*([^\s]+)", ko_raw)
            if m2:
                yangbu_ko = clean_ko_name(m2.group(1) or m2.group(2))
        return "출계", father_ko, yangbu_ko

    return "일반", father_ko, ""

def format_tree_node_label(row):
    """
    ⭐ 계보도 인물 박스 내부 표기:
    - 기본: '한글_한문' (예: '희수_禧洙')
    - 계자와 출계의 경우에는 한자 함자 뒤에 '_계자' 또는 '_출계' 형식으로 표시 (예: '종경_宗慶_계자', '여행_汝行_출계')
    - 그 외 자녀 표시((1자), (2자) 등)나 기준 표시는 붙이지 않음
    """
    base_name = format_ko_hanja_attached(row.get("인명", ""))
    adopt_type, _, _ = get_adoption_info(row)
    if adopt_type == "입계":
        return f"{base_name}_계자"
    elif adopt_type == "출계":
        return f"{base_name}_출계"
    return base_name

def format_candidate_label(row):
    """
    ⭐ 인명·계보 검색 시 결과 명부(선택 드롭다운) 표기 방식 ('한글_한문' 적용):
    - 일반: '증_嶒: 2자, 부 구, 조부 보문'
    - 입계: '종경_宗慶: 계자, 부 증, 생부 영, 조부 구'
    - 출계: '여행_汝行: 3자(출계), 부 후정, 양부 인정, 조부 세영'
    """
    person_combined = format_ko_hanja_attached(row.get("인명", ""))
    gfather_ko = clean_ko_name(row.get("조부", ""))
    rel = str(row.get("관계", "")).strip()
    adopt_type, father_ko, other_father_ko = get_adoption_info(row)

    if adopt_type == "입계":
        rel_str = rel if rel == "계자" else f"{rel}(계자)"
        parts = [f"{person_combined}: {rel_str}"]
        if father_ko:
            parts.append(f"부 {father_ko}")
        if other_father_ko:
            parts.append(f"생부 {other_father_ko}")
        if gfather_ko:
            parts.append(f"조부 {gfather_ko}")
        return ", ".join(parts)

    elif adopt_type == "출계":
        rel_str = f"{rel}(출계)" if "출계" not in rel else rel
        parts = [f"{person_combined}: {rel_str}"]
        if father_ko:
            parts.append(f"부 {father_ko}")
        if other_father_ko:
            parts.append(f"양부 {other_father_ko}")
        if gfather_ko:
            parts.append(f"조부 {gfather_ko}")
        return ", ".join(parts)

    else:
        parts = [f"{person_combined}: {rel}"]
        if father_ko:
            parts.append(f"부 {father_ko}")
        if gfather_ko:
            parts.append(f"조부 {gfather_ko}")
        return ", ".join(parts)

def highlight_corrections_html(text, search_term=None):
    if not text:
        return ""
    keywords = [
        "1699년", "1699", "一六九九", "1626년", "1626", "一六二六",
        "1558년", "1558", "一五五八", "1816년", "1816", "一八一六年",
        "1814년", "1814", "一八一四", "1771년", "1771", "一七七一",
        "1740년", "1740", "一七四〇", "1750년", "1750", "一七五〇",
        "1775년", "1775", "一七七5", "1898년", "1898", "一八九八",
        "1910년", "1910", "一九一〇", "1863년", "1863", "一八六三",
        "1850년", "1850", "一八五〇", "1870년", "1870", "一八七〇",
        "순종", "純宗", "헌종", "憲宗", "철종", "哲宗", "고종", "高宗",
        "기해", "己亥", "무신", "戊申", "자좌 오향", "자좌 오", "子坐 午向", "子坐 午",
        "내등말", "內嶝末", "오보허", "五步許", "북록 묘좌", "北麓 卯坐",
        "신재 주세붕", "慎齋 周世鵬", "병와 이형상", "瓶窩 李衡祥",
        "정헌 이종상", "定軒 李種祥", "순암 정복", "順庵 鼎福", "척암 김도화", "拓庵 金道和"
    ]
    formatted = str(text)
    for kw in keywords:
        if kw in formatted:
            formatted = formatted.replace(kw, f'<span class="red-correction">{kw}</span>')

    if search_term:
        if isinstance(search_term, (list, tuple, set)):
            terms = [str(t).strip() for t in search_term if str(t).strip()]
            terms.sort(key=len, reverse=True)
            if terms:
                pat = "|".join(re.escape(t) for t in terms)
                formatted = re.sub(
                    f"({pat})",
                    r'<span class="search-word-red">\1</span>',
                    formatted,
                    flags=re.IGNORECASE
                )
        elif str(search_term).strip():
            sterm = str(search_term).strip()
            formatted = re.sub(
                f"({re.escape(sterm)})",
                r'<span class="search-word-red">\1</span>',
                formatted,
                flags=re.IGNORECASE
            )
    return formatted

def extract_munjip_info(ko_bio):
    """
    약력 텍스트에서 문집·유고·실기·유집·유사·서집록·저서·논문 정보를 추출하여 '문집' 항목으로 반환
    """
    for line in str(ko_bio).split("\n"):
        if line.startswith("문집 :"):
            return line.replace("문집 :", "").strip()
        if line.startswith("학문·인품·업적 :"):
            c = line.replace("학문·인품·업적 :", "").strip()
            if "설서 완귀선생 유사" in c:
                return "설서 완귀선생 유사(說書 完龜先生 遺事)."
            if "문집이 있어 세상에 행해지며" in c:
                return "문집이 있어 세상에 행해짐."
            if "실기 상하 두 책이 있다" in c:
                return "실기(實記) 상·하 두 책이 있음."
            if "유고와 화부재강록이 있다" in c:
                return "유고(遺稿)와 화부재강록이 있음."
            if "고종 신축 1901년에 적변으로 유고가 소실되었다" in c:
                return "유고(고종 신축 1901년 적변으로 소실)."
            if "집안에 소장된 유고가 있" in c or "유고가 집안에 소장" in c:
                return "집안에 소장된 유고(遺稿)가 있음."
            if "유집이 있다" in c:
                return "유집(遺集)이 있음."
            if "학술논문집이 있다" in c:
                return "학술논문집이 있음."
            if "문집이 있다" in c:
                return "문집(文集)이 있음."
            if "유고가 있" in c:
                return "유고(遺稿)가 있음."
            if "서집록이 있다" in c:
                return "서집록(書集錄)이 있음."
            m = re.search(r"(저서[^\.]+있다)", c)
            if m:
                return m.group(1).strip() + "."
    return ""

SENTENCE_END_PATTERN = re.compile(
    r'(하였다|있었다|되었다|받았다|보임되었다|역임하였다|전하였다|없었다|추중하였다|존중하였다|칭송하였다|마쳤다|올랐다|돌아왔다|소실되었다|세웠다|받들었다|수훈하였다|길렀다|높았다|있도다|빛나리라|격절되었도다|만하도다|아름다움을 이었다|실려 있다|기록이다|것이다|있다|지냈다|찬술하였다|개찬하였다|배향되었다|행사한다|따님이다)\s+(?!(?:고|며|니|하고|하니|하며|하였으니|하였다|하더라도))'
)

def ensure_sentence_periods(text, with_linebreak=False):
    """
    검색 시 결과창 한글 약력에서 문장이 끝났을 때 마침표(.)를 찍어 주는 함수
    """
    if not text:
        return ""
    s = str(text).strip()
    if not s or s == "-":
        return s
    s = SENTENCE_END_PATTERN.sub(r'\1. ', s)
    parts = [p.strip() for p in re.split(r'\.\s*', s) if p.strip()]
    formatted_parts = [p if p.endswith('.') else p + '.' for p in parts]
    if with_linebreak:
        return "<br>".join(formatted_parts)
    return " ".join(formatted_parts)

def parse_person_hangul_details(row):
    """
    한글 약력 항목 구성 및 모든 문장 끝 마침표(.) 적용
    """
    gid = row.get("고유ID", "")
    _, gen_str = extract_generation_info(gid)
    gfather_ko = clean_ko_name(row.get("조부", ""))
    adopt_type, father_ko, other_father_ko = get_adoption_info(row)

    if adopt_type == "입계" and other_father_ko:
        father_display = f"{father_ko}(생부 {other_father_ko})"
    elif adopt_type == "출계" and other_father_ko:
        father_display = f"{father_ko}(양부 {other_father_ko})"
    else:
        father_display = father_ko

    fg_parts = []
    if father_display and father_display != "-":
        fg_parts.append(f"부친 {father_display}")
    if gfather_ko and gfather_ko != "-":
        fg_parts.append(f"조부 {gfather_ko}")
    father_gfather_str = ", ".join(fg_parts)

    ko_bio = str(row.get("한글 약력 정리", ""))
    first_ko = ko_bio.split("\n")[0] if ko_bio else ""

    details = [
        ("세수", ensure_sentence_periods(gen_str)),
    ]
    if father_gfather_str:
        details.append(("부친·조부", ensure_sentence_periods(father_gfather_str)))

    m_ja = re.search(r'자는\s*([^\.,\n]+)', first_ko)
    m_ho = re.search(r'호는\s*([^\.,\n]+)', first_ko)
    jaho_parts = []
    if m_ja:
        jaho_parts.append(f"자 {m_ja.group(1).strip()}")
    if m_ho:
        jaho_parts.append(f"호 {m_ho.group(1).strip()}")
    if jaho_parts:
        details.append(("자·호", ensure_sentence_periods(", ".join(jaho_parts))))

    m = re.search(r'초자는?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("초자", ensure_sentence_periods(m.group(1).strip())))
    m = re.search(r'초휘는?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("초휘", ensure_sentence_periods(m.group(1).strip())))
    m = re.search(r'일휘는?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("일휘", ensure_sentence_periods(m.group(1).strip())))
    m = re.search(r'통명은?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("통명", ensure_sentence_periods(m.group(1).strip())))
    m = re.search(r'초명은?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("초명", ensure_sentence_periods(m.group(1).strip())))

    first_parts = [p.strip() for p in first_ko.split(".") if p.strip()]
    birth_str = ""
    death_str = ""
    other_first_parts = []
    for p in first_parts:
        if re.match(r"^(자는|호는|초자는?|초휘는?|일휘는?|통명은?|초명은?|생부\s|양부\s)", p):
            continue
        if p.endswith("생") or "년생" in p:
            birth_str = p
        elif "졸" in p or "향년" in p or "수 " in p or "별세" in p:
            death_str = p
        elif p != "-":
            other_first_parts.append(p)

    if birth_str:
        details.append(("생년", ensure_sentence_periods(birth_str)))
    if death_str:
        details.append(("졸년", ensure_sentence_periods(death_str)))
    if other_first_parts:
        details.append(("신상", ensure_sentence_periods(". ".join(other_first_parts))))

    gwageo_val = ""
    hakmun_val = ""
    myo_list = []
    bae_list = []

    for line in ko_bio.split("\n"):
        if line.startswith("과거·관직 :"):
            gwageo_val = line.replace("과거·관직 :", "").strip()
        elif line.startswith("학문·인품·업적 :"):
            hakmun_val = line.replace("학문·인품·업적 :", "").strip()
        elif line.startswith("묘 :"):
            myo_list.append(line.replace("묘 :", "").strip())
        elif line.startswith("배 :") or line.startswith("배위 :"):
            bae_list.append(re.sub(r'^(배|배위)\s*:\s*', '', line).strip())

    if gwageo_val:
        details.append(("과거·관직", ensure_sentence_periods(gwageo_val)))

    munjip_val = extract_munjip_info(ko_bio)
    if munjip_val:
        details.append(("문집", ensure_sentence_periods(munjip_val)))

    if hakmun_val:
        details.append(("학력·인품·업적", ensure_sentence_periods(hakmun_val, with_linebreak=True)))

    for myo in myo_list:
        details.append(("묘", ensure_sentence_periods(myo)))
    for bae in bae_list:
        details.append(("배위", ensure_sentence_periods(bae)))

    return details

def format_hangul_only_paragraph_html(row, search_term=None):
    details = parse_person_hangul_details(row)
    lines_html = []

    for lbl, ko_val in details:
        val = str(ko_val).strip()
        if not val or val == "-":
            continue
        if search_term:
            val = highlight_corrections_html(val, search_term)
        lines_html.append(
            f"<div class='meta-para-item'><span class='meta-para-label'>• {lbl} :</span> <span class='meta-para-value'>{val}</span></div>"
        )

    return "".join(lines_html)

def render_person_result_card(target_row, search_term=None):
    """
    인물별 약력 제목란 표기:
    좌측: 한글_한문 0자  |  우측: 시직공파보 줄바꿈 1권 000p 0단 (글자크기 1p 축소)
    """
    combined_name = format_ko_hanja_attached(target_row['인명'])
    adopt_type, _, _ = get_adoption_info(target_row)
    rel_badge_text = f"{target_row['관계']}(출계)" if adopt_type == "출계" else str(target_row['관계'])
    book_loc_html = format_book_location_html(target_row.get('고유ID', ''))

    meta_paragraphs_html = format_hangul_only_paragraph_html(target_row, search_term=search_term)
    cn_raw = highlight_corrections_html(target_row.get('한문 약력 원문', ''), search_term=search_term)

    render_html(f"""
    <div id="search-result-anchor"></div>
    <div class="result-card">
        <div class="result-card-header">
            <div class="result-name-big">
                <span>{combined_name}</span>
                <span class="result-badge">{rel_badge_text}</span>
            </div>
            <div class="result-book-loc">{book_loc_html}</div>
        </div>
        
        <div style="font-weight:800; color:#1e3c72; margin-bottom:4px; font-size:11pt;">
            📌 한글 약력
        </div>
        <div class="meta-para-container">
            {meta_paragraphs_html}
        </div>
        
        <div style="font-weight:800; color:#1e3c72; margin-top:6px; margin-bottom:4px; font-size:11pt;">
            📜 한문 약력 원문
        </div>
        <div class="hanmun-raw-container">
            {cn_raw if cn_raw else '수록된 한문 약력 원문이 없습니다.'}
        </div>
    </div>
    """)

def get_siblings_list_fast(person_row):
    """
    사전 정렬된 해시맵(children_by_fgid)에서 O(1)로 형제·자매 리스트를 즉시 반환
    """
    f_gid = str(person_row.get("_father_gid", "")).strip()
    if f_gid:
        return children_by_fgid.get(f_gid, [person_row])
    f_name = str(person_row.get("부친", "")).strip()
    if f_name == "覯구":
        return children_by_fgid.get("__ROOT_GU__", [person_row])
    return [person_row]

def build_split_zones_html(male_htmls, female_htmls):
    """
    ⭐ 인명 박스를 좌·우 2개 영역으로 나누어 좌영역에는 남형제, 우영역에는 여형제(사위)를 배치
    """
    left_content = "".join(male_htmls)
    right_content = "".join(female_htmls)
    return (
        f'<div class="tree-split-wrapper">'
        f'<div class="tree-left-male-zone">{left_content}</div>'
        f'<div class="tree-right-female-zone">{right_content}</div>'
        f'</div>'
    )

def render_tree_hierarchy(target_row, all_df):
    """
    ⭐ 계보 검색 결과창 표시 규칙:
    1) 계보도 검색 결과창 박스 인명 테두리는 한 겹만 표시
    2) 계자와 출계의 경우에는 한자 함자 뒤 '00_계자' 또는 '00_출계' 형식으로 표시 (예: '종경_宗慶_계자', '여행_汝行_출계')
    3) 인명 박스를 좌·우 2개 영역으로 나누고 좌영역에는 남형제, 우영역에는 여형제(사위)를 배치
    4) 본인 창에 본인 함자 박스 뒤에 있는 기준 표시 삭제, 본인 함자 박스 바탕색은 매우 진한 황토색(#7c2d12), 글자색상은 흰색(#ffffff)
    5) 함자 박스의 가로 길이를 모두 일정하게 통일 (width: 106px)
    """
    target_gid = target_row['고유ID']
    target_name = target_row['인명']
    person_label = format_ko_hanja_attached(target_name)
    book_loc_html = format_book_location_html(target_gid)

    # 1. 상대(선대) 조상 탐색: O(1) 딕셔너리 조회로 24세 완귀공까지 즉시 추적
    ancestors = []
    curr = target_row
    up = 1
    while True:
        f_gid = str(curr.get('_father_gid', '')).strip()
        if not f_gid:
            break
        f_row = gid_to_row.get(f_gid)
        if not f_row:
            break
        sibs = get_siblings_list_fast(f_row)
        ancestors.append({"level": up, "row": f_row, "siblings": sibs})
        curr = f_row
        up += 1
        if up > 30:
            break

    # 계보 검색 결과창 제목(좌측) + 빨간 글씨 안내문 + 족보 페이지 표시(우측)
    diagram_parts = [
        f'<div id="search-result-anchor"></div>'
        f'<div class="tree-diagram-wrapper">'
        f'<div class="tree-header-bar">'
        f'<div style="font-weight:800; color:#000000; font-size:12pt; line-height:1.35;">'
        f'🌳 계보도 [{person_label}]<br>'
        f'<span style="font-size:9.5pt; font-weight:700; color:#dc2626;">인물 박스 클릭시 새 계보 생성</span>'
        f'</div>'
        f'<div class="result-book-loc">{book_loc_html}</div>'
        f'</div>'
    ]
    
    # 24세 완귀공(최상단)부터 부친까지 순서대로 렌더링
    for idx_a, anc in enumerate(reversed(ancestors)):
        lvl = anc['level']
        r = anc['row']
        g_num, _ = extract_generation_info(r['고유ID'])
        sibs = anc['siblings']
        
        if idx_a > 0:
            diagram_parts.append('<div class="tree-connector-line">▼</div>')

        if lvl == 1:
            lbl_badge = "부친"
        elif lvl == 2:
            lbl_badge = "조부"
        elif lvl == 3:
            lbl_badge = "증조"
        elif lvl == 4:
            lbl_badge = "고조"
        else:
            lbl_badge = f"{lvl}대조"

        # 5대조 이상(lvl >= 5)은 직계만 단독 표시 (좌영역에 직계 배치)
        if lvl >= 5:
            r_gid = r['고유ID']
            r_label = format_tree_node_label(r)
            node_tag = f'<a href="?tree_gid={r_gid}" target="_self" class="tree-node-box tree-node-direct">{r_label}</a>'
            split_html = build_split_zones_html([node_tag], [])
            diagram_parts.append(f"""
            <div class="tree-level-row">
                <span class="tree-gen-badge">▲ {lbl_badge}({g_num}세)</span>
                {split_html}
            </div>
            """)
        else:
            # 4대조 이하(고조·증조·조부·부친): 좌영역(남형제) / 우영역(여형제·사위) 분할 배치
            male_htmls = []
            female_htmls = []
            for s in sibs:
                is_direct = (s['고유ID'] == r['고유ID'])
                s_label = format_tree_node_label(s)
                s_rel = str(s['관계']).strip()
                s_gid = s['고유ID']

                if is_direct:
                    n_cls = "tree-node-box tree-node-direct"
                elif "녀" in s_rel:
                    n_cls = "tree-node-box tree-node-female"
                else:
                    n_cls = "tree-node-box tree-node-male"

                node_tag = f'<a href="?tree_gid={s_gid}" target="_self" class="{n_cls}">{s_label}</a>'
                if "녀" in s_rel:
                    female_htmls.append(node_tag)
                else:
                    male_htmls.append(node_tag)

            split_html = build_split_zones_html(male_htmls, female_htmls)
            diagram_parts.append(f"""
            <div class="tree-level-row">
                <span class="tree-gen-badge">▲ {lbl_badge}({g_num}세)</span>
                {split_html}
            </div>
            """)

    if len(ancestors) > 0:
        diagram_parts.append('<div class="tree-connector-line">▼</div>')

    # 2. 본인 및 형제 행:
    #    - 좌영역(남형제) / 우영역(여형제·사위) 분할 배치
    #    - 본인 함자 박스 뒤에 있는 기준 표시 삭제
    #    - 본인 함자 박스 바탕색은 매우 진한 황토색(#7c2d12), 글자색상은 흰색(#ffffff)
    t_sibs = get_siblings_list_fast(target_row)
    tg_num, _ = extract_generation_info(target_row['고유ID'])

    t_male_htmls = []
    t_female_htmls = []
    for s in t_sibs:
        is_t = (s['고유ID'] == target_row['고유ID'])
        s_label = format_tree_node_label(s)
        s_rel = str(s['관계']).strip()
        s_gid = s['고유ID']

        if is_t:
            n_cls = "tree-node-box tree-node-self"
        elif "녀" in s_rel:
            n_cls = "tree-node-box tree-node-female"
        else:
            n_cls = "tree-node-box tree-node-male"

        node_tag = f'<a href="?tree_gid={s_gid}" target="_self" class="{n_cls}">{s_label}</a>'
        if "녀" in s_rel:
            t_female_htmls.append(node_tag)
        else:
            t_male_htmls.append(node_tag)

    t_split_html = build_split_zones_html(t_male_htmls, t_female_htmls)
    diagram_parts.append(f"""
    <div class="tree-level-row-self">
        <span class="tree-gen-badge-self">● 본인<br>{tg_num}세</span>
        {t_split_html}
    </div>
    """)

    # 3. 하대 후손 탐색: 4대손(자·손·증손·현손)까지 좌영역(남형제) / 우영역(여형제·사위) 분할 배치
    max_down = 4
    down_labels = {
        1: "자",
        2: "손",
        3: "증손",
        4: "현손"
    }
    current_parents = [target_row]
    for down_lvl in range(1, max_down + 1):
        gen_children_records = []
        for p_row in current_parents:
            p_gid = p_row['고유ID']
            sorted_children = children_by_fgid.get(p_gid, [])
            for ch in sorted_children:
                gen_children_records.append(ch)

        if not gen_children_records:
            break

        male_items = []
        female_items = []

        for ch in gen_children_records:
            ch_label = format_tree_node_label(ch)
            ch_rel = str(ch['관계']).strip()
            ch_gid = ch['고유ID']

            if "녀" in ch_rel:
                node_html = f'<a href="?tree_gid={ch_gid}" target="_self" class="tree-node-box tree-node-female">{ch_label}</a>'
                female_items.append((ch, node_html))
            else:
                node_html = f'<a href="?tree_gid={ch_gid}" target="_self" class="tree-node-box tree-node-male">{ch_label}</a>'
                male_items.append((ch, node_html))

        d_male_htmls = [html_str for _, html_str in male_items]
        d_female_htmls = [html_str for _, html_str in female_items]
        d_split_html = build_split_zones_html(d_male_htmls, d_female_htmls)

        ordered_items = male_items + female_items
        next_gen_rows = [ch_row for ch_row, _ in ordered_items]

        diagram_parts.append('<div class="tree-connector-line">▼</div>')
        dg_num = (tg_num + down_lvl) if tg_num else (24 + down_lvl)
        d_badge = down_labels.get(down_lvl, f"{down_lvl}대손")
        diagram_parts.append(f"""
        <div class="tree-level-row">
            <span class="tree-gen-badge">▼ {d_badge}({dg_num}세)</span>
            {d_split_html}
        </div>
        """)
        current_parents = next_gen_rows

    diagram_parts.append('</div>')
    render_html("".join(diagram_parts))

    render_person_result_card(target_row)

# =============================================================================
# 4. 최상단 헤더(황색 배경 + 좌측 초대형 아이콘 밀착 + 우측 타이틀) 및 검색 인터페이스
# =============================================================================
icon_img_tag = f'<img src="{icon_b64}" style="width:118px; height:118px; object-fit:contain; display:block; margin-left:0px;">' if icon_b64 else '<span style="font-size:68px;">📜</span>'

bg_style_tag = f"<style>.stApp {{ background-image: url('{bg_b64}') !important; background-size: cover !important; background-position: center top !important; opacity: 1 !important; }}</style>" if bg_b64 else ""
if bg_style_tag:
    render_html(bg_style_tag)

# 타이틀 상자 출력 (모든 창의 한글 한문 표시: 한글_한문 -> 증_嶒)
render_html(f"""
<div class="title-container">
    <div class="title-left-icon">
        {icon_img_tag}
    </div>
    <div class="title-right-content">
        <div class="title-clan-name">광주안씨 완귀공파 종중</div>
        <div class="title-main-center">족보 검색 시스템</div>
        <div class="title-sub-bottom">- 완귀공(휘 증_嶒, 24세) ~ 44세<br>총 2,887인(고위 기준) -</div>
    </div>
</div>
""")

# 검색 엔터 시 인명검색 라인이 화면 제일 위에 위치하도록 앵커 배치
render_html("""
<div id="search-section-anchor"></div>
<div class="compact-search-header">
    <div class="compact-search-subtitle">다음 중 하나를 선택하고 검색어를 한글로 입력</div>
</div>
""")

if df is None:
    st.error(f"데이터 로드 오류: {load_msg}")
    st.info("파일 이름이 `genealogy.xlsx` 또는 `genealogy.db`인지 확인해 주시기 바랍니다.")
    st.stop()

# URL 파라미터로 계보 클릭 시 처리
param_tree_gid = get_query_param("tree_gid")
if param_tree_gid:
    st.session_state['active_tree_gid'] = param_tree_gid
    st.session_state['default_search_type'] = "계보 검색"

# 검색 방식 선택 (인명 검색, 계보 검색, 단어 검색)
default_idx = 0
if st.session_state.get('default_search_type') == "계보 검색":
    default_idx = 1

search_type = st.radio(
    "검색 종류 선택",
    ["인명 검색", "계보 검색", "단어 검색"],
    index=default_idx,
    horizontal=True,
    label_visibility="collapsed"
)

# 검색어 안내문은 기존 그대로 유지하되, 내부 검색은 자·초자·일휘·초휘까지 모두 포함
if search_type == "인명 검색":
    query_val = st.text_input(
        "검색어 입력",
        placeholder="함자, 성명, 호, 초명, 통명을 한글로 입력",
        key="name_query",
        label_visibility="collapsed"
    )
elif search_type == "계보 검색":
    query_val = st.text_input(
        "검색어 입력",
        placeholder="함자, 성명, 호, 초명, 통명을 한글로 입력",
        key="lineage_query",
        label_visibility="collapsed"
    )
else:
    query_val = st.text_input(
        "검색어 입력",
        placeholder="관직, 과거, 학문 등 모든 약력 내용 검색 가능",
        key="word_query",
        label_visibility="collapsed"
    )

# =============================================================================
# 5. 검색 결과 창
# =============================================================================
if search_type == "인명 검색" and query_val.strip():
    q = query_val.strip()
    matched_df = fast_search_by_name(df, q)
    
    if len(matched_df) == 0:
        render_html(f"""
        <div class="no-result-notice">
            입력하신 검색어 '{q}'에 해당하는 인물을 찾을 수 없습니다.
        </div>
        """)
        scroll_to_search_top()
    else:
        if len(matched_df) > 1:
            render_html(f"""
            <div class="multi-ancestor-banner">
                📢 동명이인 또는 복수 후보가 총 {len(matched_df)}위 검색되었습니다. 아래에서 조회할 인물을 선택하세요.
            </div>
            """)
            selected_idx = st.selectbox(
                "선조 선택",
                range(len(matched_df)),
                format_func=lambda i: format_candidate_label(matched_df.iloc[i]),
                label_visibility="collapsed"
            )
            target_row = matched_df.iloc[selected_idx]
        else:
            target_row = matched_df.iloc[0]
            
        render_person_result_card(target_row)
        scroll_to_search_top()

elif search_type == "계보 검색":
    q = query_val.strip()

    # 검색어가 비어 있으면(검색 전에는) 이전 계보 선택 상태를 초기화하여 아무것도 표시하지 않음
    if not q:
        st.session_state['last_lineage_query'] = ''
        st.session_state['active_tree_gid'] = None
        if hasattr(st, "query_params") and "tree_gid" in st.query_params:
            del st.query_params["tree_gid"]
    elif q != st.session_state.get('last_lineage_query', ''):
        st.session_state['last_lineage_query'] = q
        st.session_state['active_tree_gid'] = None
        if hasattr(st, "query_params") and "tree_gid" in st.query_params:
            del st.query_params["tree_gid"]

    if q:
        l_matches = fast_search_by_name(df, q)
        if len(l_matches) == 0:
            render_html(f"""
            <div class="no-result-notice">
                계보를 작성할 선조 인명('{q}')을 찾을 수 없습니다.
            </div>
            """)
            scroll_to_search_top()
        else:
            # 1) 검색 후 복수 후보(명부)가 있는 경우 상단에 항상 명부를 먼저 표시
            if len(l_matches) > 1:
                render_html(f"""
                <div class="multi-ancestor-banner">
                    📢 동명이인 또는 복수 후보가 총 {len(l_matches)}위 검색되었습니다. 아래에서 계보 기준 인물을 선택하세요.
                </div>
                """)
                sel_idx = st.selectbox(
                    "계보 기준 인물 선택",
                    range(len(l_matches)),
                    format_func=lambda i: format_candidate_label(l_matches.iloc[i]),
                    label_visibility="collapsed"
                )
                initial_target = l_matches.iloc[sel_idx]
            else:
                initial_target = l_matches.iloc[0]

            # 사용자가 드롭다운 명부에서 다른 인물로 변경하면 즉시 해당 인물의 계보로 전환
            if st.session_state.get('last_selected_initial_gid') != initial_target['고유ID']:
                st.session_state['last_selected_initial_gid'] = initial_target['고유ID']
                st.session_state['active_tree_gid'] = None
                if hasattr(st, "query_params") and "tree_gid" in st.query_params:
                    del st.query_params["tree_gid"]

            # 계보도 도형 박스를 클릭해 다른 인물로 이동한 경우 해당 인물을, 아니면 초기 선택 인물을 표시
            tree_target = initial_target
            active_gid = st.session_state.get('active_tree_gid')
            if active_gid and active_gid != initial_target['고유ID']:
                if active_gid in gid_to_row:
                    tree_target = gid_to_row[active_gid]

            # 검색 전에는 표시하지 않고, 검색 후 계보도에서 다른 인물을 클릭해 이동했을 때만 복귀 박스 표시
            if active_gid and active_gid != initial_target['고유ID']:
                if st.button("🔄 초기 검색 인물로 복귀", use_container_width=True):
                    st.session_state['active_tree_gid'] = None
                    if hasattr(st, "query_params") and "tree_gid" in st.query_params:
                        del st.query_params["tree_gid"]
                    st.rerun()

            render_tree_hierarchy(tree_target, df)
            scroll_to_search_top()

elif search_type == "단어 검색" and query_val.strip():
    term = query_val.strip()

    # 1) 과거 통합 검색어 그룹
    if term in EXAM_GROUP_KEYWORDS:
        search_terms_list = EXAM_GROUP_KEYWORDS
        regex_pattern = "|".join(re.escape(k) for k in EXAM_GROUP_KEYWORDS)
        banner_term_label = f"'{term}' (과거·문과·급제·사마·진사·생원 등 통합 검색)"
    # 2) 문집·저술 통합 검색어 그룹 (문집, 유집, 유고, 실기, 서집록, 저서, 논문)
    elif term in LITERARY_GROUP_KEYWORDS:
        search_terms_list = LITERARY_GROUP_KEYWORDS
        regex_pattern = "|".join(re.escape(k) for k in LITERARY_GROUP_KEYWORDS)
        banner_term_label = f"'{term}' (문집·유집·유고·실기·서집록·저서·논문 통합 검색)"
    else:
        search_terms_list = [term]
        regex_pattern = re.escape(term)
        banner_term_label = f"'{term}'"

    mask = (
        df['한글 약력 정리'].astype(str).str.contains(regex_pattern, case=False, na=False, regex=True) |
        df['한문 약력 정리'].astype(str).str.contains(regex_pattern, case=False, na=False, regex=True) |
        df['한글 약력 원문'].astype(str).str.contains(regex_pattern, case=False, na=False, regex=True) |
        df['한문 약력 원문'].astype(str).str.contains(regex_pattern, case=False, na=False, regex=True)
    )
    word_matches = df[mask].reset_index(drop=True)
    word_matches = sort_search_candidates_df(word_matches)
    
    if len(word_matches) == 0:
        render_html(f"""
        <div class="no-result-notice">
            입력하신 단어 '{term}'을(를) 포함하고 있는 약력 기록이 없습니다.
        </div>
        """)
        scroll_to_search_top()
    else:
        render_html(f"""
        <div class="multi-ancestor-banner">
            📢 검색어 <span class="search-word-red">{banner_term_label}</span>을(를) 포함하는 약력을 가진 선조가 총 {len(word_matches):,}위 검색되었습니다. 아래 목록에서 선조를 선택하세요.
        </div>
        """)
        chosen_idx = st.selectbox(
            "조회할 선조를 선택하세요",
            range(len(word_matches)),
            format_func=lambda i: format_candidate_label(word_matches.iloc[i]),
            label_visibility="collapsed"
        )
        chosen_row = word_matches.iloc[chosen_idx]
        
        render_person_result_card(chosen_row, search_term=search_terms_list)
        scroll_to_search_top()

# =============================================================================
# 6. 종중 주요 유적 및 문화재 안내 ('클릭 시 국가유산 포털로 이동' 문구 삭제 + 녹색 바탕)
# =============================================================================
render_html("""
<div class="heritage-sub-container">
    <div class="heritage-sub-header">🏛️ 종중 주요 유적 및 문화재 안내</div>
    <div class="heritage-sub-grid">
        <a href="https://www.heritage.go.kr/heri/cul/culSelectDetail.do?pageNo=1_1_1_1&sngl=Y&ccbaCpno=2443700200000#audioDiv" target="_blank" class="heritage-sub-card">
            <div class="h-card-name">1. 영천 완귀정(경상북도 민속문화재 제20호) <span>↗</span></div>
            <div class="h-card-desc">완귀 안증(1494~1553) 건립 | 경북 영천시 도남동</div>
        </a>
        <a href="https://www.heritage.go.kr/heri/cul/culSelectDetail.do?pageNo=1_1_1_1&sngl=Y&ccbaCpno=3413705940000" target="_blank" class="heritage-sub-card">
            <div class="h-card-name">2. 영천 육우당(육우당_六友堂) <span>↗</span></div>
            <div class="h-card-desc">경상북도 영천시 도남동 566 (2012.09.17. 지정)</div>
        </a>
        <a href="https://www.heritage.go.kr/heri/cul/culSelectDetail.do?pageNo=1_1_1_1&sngl=N&ccbaAsno=0007220000000&ccbaCpno=3412603100001" target="_blank" class="heritage-sub-card">
            <div class="h-card-name">3. 청도 만회정(만회정_晩悔亭) <span>↗</span></div>
            <div class="h-card-desc">경북 청도군 운문면 지촌리 907 (2026.02.19. 지정)</div>
        </a>
        <div class="heritage-sub-card">
            <div class="h-card-name">4. 삼괴당(삼괴당_三槐堂)</div>
            <div class="h-card-desc">광주안씨 문중 선조 강학 및 추모 유적</div>
        </div>
    </div>
</div>
""")

# =============================================================================
# 7. 하단 푸터 (Footer): '구축(고위 기준 2,887명)' 표기 + 줄바꿈 연락처 + 비볼드체
# =============================================================================
render_html("""
<div style="text-align:center; padding:8px 4px 6px 4px; color:#ffffff; font-size:9.3pt; font-weight:400; border-top:1px solid #475569; margin-top:8px; line-height:1.5; text-shadow: 0 1px 2px rgba(0,0,0,0.7);">
    📚 데이터 출처: 2008년 발간 시직공파보 PDF파일 변환후 데이터베이스 구축(고위 기준 2,887명)<br>
    &lt;연락처 안동수 010 5097 6403&gt;
</div>
""")
