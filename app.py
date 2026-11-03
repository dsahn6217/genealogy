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

# ⭐ st.set_page_config는 반드시 모든 st.* 명령보다 가장 먼저 호출되어야 함
st.set_page_config(
    page_title="광주안씨 완귀공파 종중 족보 검색 시스템",
    page_icon=ICON_FILE if os.path.exists(ICON_FILE) else "📜",
    layout="wide",
    initial_sidebar_state="collapsed"
)

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
    1) 검색창(인명·계보·단어 검색)에서 엔터(Enter)를 누르는 즉시 모바일 자판(가상 키보드) 강제 숨김(blur)
    2) 슬라이더 막대 오른쪽 끝이나 손잡이 위에 뜨는 빨간색 숫자(ThumbValue / 1~6 숫자 텍스트) 실시간 완전 제거
    """
    components.html(
        """
        <script>
            (function() {
                try {
                    var parentDoc = window.parent.document;

                    // 1. 엔터 키 입력 시 즉시 입력창 포커스 해제하여 모바일 자판 숨김
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

                    // 2. 슬라이더 내부의 빨간 숫자 표시(1~6) 요소를 찾아 완전히 숨김
                    function hideSliderRedNumbers() {
                        try {
                            var sliders = parentDoc.querySelectorAll('[data-testid="stSlider"]');
                            sliders.forEach(function(sl) {
                                var nodes = sl.querySelectorAll('div, span, p');
                                nodes.forEach(function(el) {
                                    if (el.children.length === 0 && el.textContent) {
                                        var txt = el.textContent.trim();
                                        if (/^[1-6]$/.test(txt)) {
                                            el.style.setProperty('display', 'none', 'important');
                                            el.style.setProperty('visibility', 'hidden', 'important');
                                            el.style.setProperty('opacity', '0', 'important');
                                            el.style.setProperty('font-size', '0px', 'important');
                                            el.style.setProperty('color', 'transparent', 'important');
                                        }
                                    }
                                });
                            });
                        } catch (err) {}
                    }

                    hideSliderRedNumbers();
                    setTimeout(hideSliderRedNumbers, 100);
                    setTimeout(hideSliderRedNumbers, 300);

                    if (!parentDoc._sliderObserverAdded) {
                        var obs = new MutationObserver(function() {
                            hideSliderRedNumbers();
                        });
                        obs.observe(parentDoc.body, { childList: true, subtree: true });
                        parentDoc._sliderObserverAdded = true;
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
# 1. 디자인 스타일 CSS (타이틀 박스 황색 배경, 녹색 복귀 버튼, 빨간 숫자 제거 등)
# =============================================================================
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700;800;900&family=Nanum+Myeongjo:wght@400;700&display=swap');

html, body, [class*="css"], .stMarkdown, p, div {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', 'Malgun Gothic', sans-serif;
}

/* 기본 상단 헤더 숨김 */
header[data-testid="stHeader"] {
    display: none !important;
}

.main .block-container, 
[data-testid="stAppViewBlockContainer"], 
.block-container {
    max-width: 100% !important;
    width: 100% !important;
    margin-left: auto !important;
    margin-right: auto !important;
    padding-top: 1.0rem !important;
    padding-bottom: 1.2rem !important;
    padding-left: 0.65rem !important;
    padding-right: 0.65rem !important;
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
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
    font-size: 13.5pt;
    font-weight: 800;
    color: #1e3c72;
    line-height: 1.25;
    letter-spacing: -0.3px;
    margin-bottom: 2px;
}

.title-main-center {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
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
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
    font-size: 10.5pt;
    font-weight: 700;
    color: #1e293b;
    text-align: center;
    width: 100%;
    margin-top: 1px;
    line-height: 1.35;
    letter-spacing: -0.2px;
}

/* ⭐ 위 타이틀 창과 간격을 넓힌 안내 문구 라인 */
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

/* ⭐ 직계 상대/하대 선조수 선택 슬라이더:
   1) 문구: '직계 상대 선조수 선택' / '직계 하대 선조수 선택'
   2) 하얀색 대수 글자(1대~6대): 크기 축소(9.2pt), 비볼드체(font-weight: 400)
   3) 빨간 막대 오른쪽 끝 위에 뜨는 빨간색 숫자 표시 완전 삭제 */
.slider-custom-header {
    display: flex;
    flex-direction: column;
    margin-top: 0px !important;
    margin-bottom: -14px !important;
    padding: 0px 2px !important;
}
.slider-title-text {
    color: #ffffff !important;
    font-size: 11pt !important;
    font-weight: 800 !important;
    text-shadow: 0 1px 3px rgba(0,0,0,0.85) !important;
    line-height: 1.15 !important;
    margin-bottom: 2px !important;
}
.slider-ticks-row {
    display: flex !important;
    flex-direction: row !important;
    justify-content: space-between !important;
    align-items: center !important;
    width: 100% !important;
    padding: 0px 4px !important;
    margin-bottom: 0px !important;
    box-sizing: border-box !important;
}
.slider-tick-item {
    color: #ffffff !important;
    font-size: 9.2pt !important;
    font-weight: 400 !important;
    text-shadow: 0 1px 2px rgba(0,0,0,0.85) !important;
    line-height: 1.0 !important;
    text-align: center !important;
}

div[data-testid="stSlider"] {
    margin-top: -8px !important;
    margin-bottom: -8px !important;
    padding-top: 0px !important;
    padding-bottom: 0px !important;
    font-size: 0px !important;
    color: transparent !important;
}
div[data-testid="stSlider"] [data-testid="stWidgetLabel"] {
    display: none !important;
    height: 0px !important;
    margin: 0px !important;
    padding: 0px !important;
}
div[data-testid="stSlider"] div[data-baseweb="slider"] {
    margin-top: 0px !important;
    margin-bottom: 0px !important;
    padding-top: 6px !important;
    padding-bottom: 4px !important;
    font-size: 0px !important;
    color: transparent !important;
}
/* 슬라이더 내부의 모든 텍스트(빨간 막대 오른쪽 끝 위의 빨간 숫자 포함) 투명화 및 크기 0 처리 */
div[data-testid="stSlider"] * {
    font-size: 0px !important;
    color: transparent !important;
    -webkit-text-fill-color: transparent !important;
}
/* 슬라이더 하얀 선(트랙): 검은색 테두리 + 5등분 경계 눈금선 표시 */
div[data-testid="stSlider"] div[data-baseweb="slider"] > div > div {
    height: 10px !important;
    background: repeating-linear-gradient(
        to right,
        #ffffff 0%,
        #ffffff calc(20% - 1.5px),
        #000000 calc(20% - 1.5px),
        #000000 20%
    ) !important;
    border: 2px solid #000000 !important;
    border-radius: 5px !important;
    box-sizing: border-box !important;
}
/* 슬라이더 조절 손잡이(Thumb) */
div[data-testid="stSlider"] [role="slider"] {
    width: 20px !important;
    height: 20px !important;
    background-color: #facc15 !important;
    border: 2.5px solid #000000 !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.6) !important;
}
/* 슬라이더의 숫자 표시용 자식 요소들 완전 숨김 */
div[data-testid="stSlider"] [data-testid="stThumbValue"],
div[data-testid="stSlider"] [data-testid="stTickBar"],
div[data-testid="stSlider"] [data-testid="stTickBarMin"],
div[data-testid="stSlider"] [data-testid="stTickBarMax"],
div[data-testid="stSlider"] [role="slider"] > div,
div[data-testid="stSlider"] [role="slider"] *,
div[data-testid="stSlider"] div[data-baseweb="slider"] > div:nth-child(n+2) {
    display: none !important;
    visibility: hidden !important;
    opacity: 0 !important;
    height: 0px !important;
    width: 0px !important;
    overflow: hidden !important;
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
    color: #1e3c72;
    display: flex;
    align-items: center;
    gap: 8px;
}
.result-badge {
    font-size: 10.5pt;
    font-weight: 800;
    background: #eff6ff;
    color: #1d4ed8;
    padding: 2px 8px;
    border-radius: 10px;
    border: 1px solid #bfdbfe;
}

/* 인물 한글 약력 컨테이너 */
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

/* 계보도 도형 박스: 클릭 가능한 링크 스타일 적용 (클릭 시 새 계보 즉시 생성) */
.tree-diagram-wrapper {
    width: 100% !important;
    box-sizing: border-box !important;
    background: #ffffff;
    border: 2px solid #1e3c72;
    border-radius: 8px;
    padding: 8px 10px;
    margin-top: 2px;
    margin-bottom: 6px;
}
.tree-level-row {
    display: flex;
    align-items: center;
    margin-bottom: 0px !important;
    padding: 4px 6px;
    background: #f8fafc;
    border-radius: 5px;
    border: 1px solid #cbd5e1;
    line-height: 1.35;
}
.tree-connector-line {
    text-align: center;
    color: #1e3c72;
    font-weight: 900;
    font-size: 9.5pt;
    line-height: 1.0;
    margin: 0px !important;
    padding: 0px;
}
.tree-gen-badge {
    min-width: 78px;
    font-weight: 700;
    color: #1e3c72;
    font-size: 9.5pt;
    padding: 2px 5px;
    background: #dbeafe;
    border-radius: 4px;
    text-align: center;
    margin-right: 6px;
}
.tree-nodes-container {
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
    align-items: center;
}
.tree-node-box {
    display: inline-block !important;
    text-decoration: none !important;
    cursor: pointer !important;
    background: #ffffff !important;
    border: 1.5px solid #64748b !important;
    border-radius: 4px !important;
    padding: 2px 7px !important;
    font-size: 9.5pt !important;
    color: #0f172a !important;
    line-height: 1.3 !important;
    transition: all 0.15s ease-in-out !important;
}
.tree-node-box:hover {
    background: #dbeafe !important;
    border-color: #2563eb !important;
    color: #1d4ed8 !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 2px 4px rgba(0,0,0,0.15) !important;
}
.tree-node-target {
    background: #eff6ff !important;
    border: 2px solid #2563eb !important;
    font-weight: 800 !important;
    color: #1e40af !important;
}
.tree-h-line {
    color: #64748b;
    font-weight: 700;
    font-size: 8.5pt;
    margin: 0 1px;
}

/* ⭐ 종중 주요 유적 및 문화재 안내: 바탕색 녹색(#15803d), 헤더 흰색 글씨 */
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

/* 업로드하신 수정app.py의 모바일 좌우 여백 동일화 CSS 그대로 적용 */
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
        padding-left: 12px !important;
        padding-right: 12px !important;
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
# 2. 데이터 로드 및 초고속(0.2초) 고유ID 기반 부모-자녀 매핑 전처리
# =============================================================================
def parse_gid_tuple(gid):
    """고유ID(예: C-529-4-1-37)를 (권, 페이지, 단, 순서, 세수)로 분리"""
    m = re.match(r"^([A-Z])-(\d+)-(\d+)-(\d+)-(\d+)$", str(gid).strip())
    if m:
        return m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4)), int(m.group(5))
    return "", 0, 0, 0, 0

@st.cache_data
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
                return None, "시트에 데이터가 없습니다."
            header = [str(c).strip() if c is not None else f"Col{idx}" for idx, c in enumerate(rows[0])]
            raw_df = pd.DataFrame(rows[1:], columns=header).fillna("")
            status_msg = "성공 (Excel)"
        except Exception as e:
            return None, str(e)

    if raw_df is None or raw_df.empty:
        return None, f"데이터 파일({DATA_FILE})을 찾을 수 없습니다."

    # 초고속 딕셔너리 인덱싱(O(N), 0.2초)으로 동명이인 부친 고유ID(_father_gid) 정확히 매핑
    records = raw_df.to_dict("records")
    by_name = {}
    for r in records:
        vol, page, tier, seq, gen = parse_gid_tuple(r.get("고유ID", ""))
        r["_vol"] = vol
        r["_page"] = page
        r["_tier"] = tier
        r["_seq"] = seq
        r["_gen"] = gen
        nm = str(r.get("인명", "")).strip()
        by_name.setdefault(nm, []).append(r)

    father_gids = []
    for r in records:
        f_name = str(r.get("부친", "")).strip()
        if not f_name or f_name == "覯구" or f_name == "-":
            father_gids.append("")
            continue
        cands_all = by_name.get(f_name, [])
        if not cands_all:
            father_gids.append("")
            continue
        if len(cands_all) == 1:
            father_gids.append(str(cands_all[0]["고유ID"]))
            continue

        # 1차: 세수(본인 세수 - 1) 일치 후보
        c_gen = r["_gen"]
        cands = [c for c in cands_all if c["_gen"] == c_gen - 1] or cands_all
        if len(cands) == 1:
            father_gids.append(str(cands[0]["고유ID"]))
            continue

        # 2차: 후보의 부친 == 본인의 조부 일치 후보
        gf_name = str(r.get("조부", "")).strip()
        if gf_name and gf_name != "-":
            cands_gf = [c for c in cands if str(c.get("부친", "")).strip() == gf_name]
            if len(cands_gf) == 1:
                father_gids.append(str(cands_gf[0]["고유ID"]))
                continue
            elif len(cands_gf) > 1:
                cands = cands_gf

        # 3차: 같은 권(Volume)에서 페이지가 작거나 같으면서 가장 가까운 선조 선택
        c_vol, c_page = r["_vol"], r["_page"]
        same_vol = [c for c in cands if c["_vol"] == c_vol and c["_page"] <= c_page]
        if same_vol:
            same_vol.sort(key=lambda x: (-x["_page"], x["_seq"]))
            father_gids.append(str(same_vol[0]["고유ID"]))
            continue

        # 4차: 생년/배위 상세 기록이 있는 본 기록 우선
        non_empty = [c for c in cands if ("생" in str(c.get("한글 약력 정리", "")) or "배 :" in str(c.get("한글 약력 정리", "")))]
        if len(non_empty) == 1:
            father_gids.append(str(non_empty[0]["고유ID"]))
            continue

        father_gids.append(str(cands[0]["고유ID"]))

    raw_df["_father_gid"] = father_gids
    return raw_df, status_msg

df, load_msg = load_and_preprocess_genealogy()

# =============================================================================
# 3. 보조 유틸리티 함수: 한글한자 병기, 과거 통합 검색어, 한글 약력 문단 생성
# =============================================================================
# ⭐ 과거 관련 통합 검색어 그룹 (이 중 1단어만 입력해도 그룹 전체 단어를 검색)
EXAM_GROUP_KEYWORDS = [
    "과거", "문과", "급제", "사마", "별시", "대과", "무과", "잡과", "장원",
    "진사", "생원", "향시", "회시", "복시", "현랑과", "현량과", "증광시",
    "전시", "식년시", "알성시", "정시", "춘당대시", "갑과", "을과", "병과"
]

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
    ⭐ 검색 결과창 제목의 함자를 '한글한자' 붙여서 병기 (예: '嶒증' -> '증嶒', '景時경시' -> '경시景時')
    """
    s = str(raw_val).strip()
    ko, hj = split_name(s)
    if ko and hj:
        return f"{ko}{hj}"
    return s

def match_name_query(row, q):
    """
    인명 검색 및 계보 검색 시 함자, 성명, 자(字), 초자(初字), 호(號), 일휘(一諱), 초휘(初諱), 초명(初名), 통명(通名) 모두 검색
    """
    q = q.strip()
    if not q:
        return False
    q_alt = q[1:] if q.startswith('안') and len(q) > 1 else q
    
    # 1. 함자, 성명 (인명 컬럼)
    name_str = str(row.get('인명', '')).strip()
    ko_name = re.sub(r'[^가-힣]+', '', name_str)
    cn_name = re.sub(r'[^一-龥]+', '', name_str)
    
    if q in name_str or q_alt in name_str:
        return True
    if ko_name and (q == ko_name or q == '안' + ko_name or q_alt == ko_name):
        return True
    if cn_name and (q == cn_name or q == '安' + cn_name or q_alt == cn_name):
        return True
        
    # 2. 자, 초자, 호, 일휘, 초휘, 초명, 통명 (한글 약력 정리 첫 줄)
    first_line_ko = str(row.get('한글 약력 정리', '')).split('\n')[0]
    ko_patterns = [
        r'자는\s*([^\.,\n\s]+)',
        r'초자는?\s*([^\.,\n\s]+)',
        r'호는\s*([^\.,\n\s]+)',
        r'일휘는?\s*([^\.,\n\s]+)',
        r'초휘는?\s*([^\.,\n\s]+)',
        r'초명은?\s*([^\.,\n\s]+)',
        r'통명은?\s*([^\.,\n\s]+)',
    ]
    for pattern in ko_patterns:
        for m in re.finditer(pattern, first_line_ko):
            val = m.group(1).strip()
            if q == val or q_alt == val or q in val:
                return True

    # 3. 字, 初字, 號, 一諱, 初諱, 初名, 通名 (한문 약력 정리 첫 줄)
    first_line_cn = str(row.get('한문 약력 정리', '')).split('\n')[0]
    cn_patterns = [
        r'字\s*([^\.,\n\s]+)',
        r'初字\s*([^\.,\n\s]+)',
        r'號\s*([^\.,\n\s]+)',
        r'一諱\s*([^\.,\n\s]+)',
        r'初諱\s*([^\.,\n\s]+)',
        r'初名\s*([^\.,\n\s]+)',
        r'通名\s*([^\.,\n\s]+)',
    ]
    for pattern in cn_patterns:
        for m in re.finditer(pattern, first_line_cn):
            val = m.group(1).strip()
            if q == val or q_alt == val or q in val:
                return True
                
    return False

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

def format_candidate_label(row):
    _, gen_str = extract_generation_info(row.get("고유ID", ""))
    person_ko = clean_ko_name(row.get("인명", ""))
    gfather_ko = clean_ko_name(row.get("조부", ""))
    rel = str(row.get("관계", "")).strip()
    adopt_type, father_ko, _ = get_adoption_info(row)

    if adopt_type == "입계":
        return f"[{gen_str}] {person_ko}  {rel} (계부:{father_ko}, 조부:{gfather_ko}) - 고유ID: {row['고유ID']}"
    elif adopt_type == "출계":
        return f"[{gen_str}] {person_ko}  {rel} (출계) (생부:{father_ko}, 조부:{gfather_ko}) - 고유ID: {row['고유ID']}"
    else:
        return f"[{gen_str}] {person_ko}  {rel} (부:{father_ko}, 조부:{gfather_ko}) - 고유ID: {row['고유ID']}"

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
    약력 텍스트에서 문집·유고·실기·유집·유사·서집록·저서 정보를 추출하여 '문집' 항목으로 반환
    """
    for line in str(ko_bio).split("\n"):
        if line.startswith("문집 :"):
            return line.replace("문집 :", "").strip()
        if line.startswith("학문·인품·업적 :"):
            c = line.replace("학문·인품·업적 :", "").strip()
            if "설서 완귀선생 유사" in c:
                return "설서 완귀선생 유사(說書 完龜先生 遺事)"
            if "문집이 있어 세상에 행해지며" in c:
                return "문집이 있어 세상에 행해짐"
            if "실기 상하 두 책이 있다" in c:
                return "실기(實記) 상·하 두 책이 있음"
            if "유고와 화부재강록이 있다" in c:
                return "유고(遺稿)와 화부재강록이 있음"
            if "고종 신축 1901년에 적변으로 유고가 소실되었다" in c:
                return "유고(고종 신축 1901년 적변으로 소실)"
            if "집안에 소장된 유고가 있" in c or "유고가 집안에 소장" in c:
                return "집안에 소장된 유고(遺稿)가 있음"
            if "유집이 있다" in c:
                return "유집(遺集)이 있음"
            if "학술논문집이 있다" in c:
                return "학술논문집이 있음"
            if "문집이 있다" in c:
                return "문집(文集)이 있음"
            if "유고가 있" in c:
                return "유고(遺稿)가 있음"
            if "서집록이 있다" in c:
                return "서집록(書集錄)이 있음"
            m = re.search(r"(저서[^\.]+있다)", c)
            if m:
                return m.group(1).strip()
    return ""

def format_hakmun_sentences_with_linebreaks(text):
    """
    학력·인품·업적 항목에서 문장이 끝난 경우 마침표(.)를 찍고 줄바꿈(<br>) 처리
    """
    if not text:
        return ""
    s = str(text).strip()
    s = re.sub(
        r'(하였다|있었다|되었다|받았다|보임되었다|역임하였다|전하였다|없었다|추중하였다|존중하였다|칭송하였다|마쳤다|올랐다|돌아왔다|소실되었다|세웠다|받들었다|수훈하였다|길렀다|높았다|있도다|빛나리라|격절되었도다|만하도다|아름다움을 이었다|실려 있다|기록이다|것이다|있다)\s+(?!(?:고|며|니|하고|하니|하며|하였으니|하였다|하더라도))',
        r'\1. ',
        s
    )
    parts = [p.strip() for p in re.split(r'\.\s+', s) if p.strip()]
    formatted_parts = [p if p.endswith('.') else p + '.' for p in parts]
    return "<br>".join(formatted_parts)

def parse_person_hangul_details(row):
    """
    인물 한글 약력 항목 구성:
    1) 부, 조부 통합 -> '부친·조부' 항목 (부친 00, 조부 00)
    2) 자, 호 통합 -> '자·호' 항목 (자 00, 호 00 형식)
    3) '학력·인품·업적' 바로 위에 '문집' 항목 배치
    4) '학력·인품·업적'은 문장이 끝날 때마다 마침표를 찍고 줄바꿈 처리
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

    # 부친·조부 통합 표기
    fg_parts = []
    if father_display and father_display != "-":
        fg_parts.append(f"부친 {father_display}")
    if gfather_ko and gfather_ko != "-":
        fg_parts.append(f"조부 {gfather_ko}")
    father_gfather_str = ", ".join(fg_parts)

    ko_bio = str(row.get("한글 약력 정리", ""))
    first_ko = ko_bio.split("\n")[0] if ko_bio else ""

    details = [
        ("세수", gen_str),
    ]
    if father_gfather_str:
        details.append(("부친·조부", father_gfather_str))

    # 자·호 통합 표기 (자 00, 호 00 형식)
    m_ja = re.search(r'자는\s*([^\.,\n]+)', first_ko)
    m_ho = re.search(r'호는\s*([^\.,\n]+)', first_ko)
    jaho_parts = []
    if m_ja:
        jaho_parts.append(f"자 {m_ja.group(1).strip()}")
    if m_ho:
        jaho_parts.append(f"호 {m_ho.group(1).strip()}")
    if jaho_parts:
        details.append(("자·호", ", ".join(jaho_parts)))

    m = re.search(r'초자는?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("초자", m.group(1).strip()))
    m = re.search(r'초휘는?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("초휘", m.group(1).strip()))
    m = re.search(r'일휘는?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("일휘", m.group(1).strip()))
    m = re.search(r'통명은?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("통명", m.group(1).strip()))
    m = re.search(r'초명은?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("초명", m.group(1).strip()))

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
        details.append(("과거·관직", gwageo_val))

    # '학력·인품·업적' 바로 위에 '문집' 항목 삽입
    munjip_val = extract_munjip_info(ko_bio)
    if munjip_val:
        details.append(("문집", munjip_val))

    if hakmun_val:
        details.append(("학력·인품·업적", format_hakmun_sentences_with_linebreaks(hakmun_val)))

    for myo in myo_list:
        details.append(("묘", myo))
    for bae in bae_list:
        details.append(("배위", bae))

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
    ⭐ 검색 결과창 제목 함자: 한글한자 붙여서 병기 (예: 증嶒, 경시景時)
    ⭐ '신상 및 계보 정보' -> '인물 한글 약력'으로 변경
    """
    combined_name = format_ko_hanja_attached(target_row['인명'])
    adopt_type, _, _ = get_adoption_info(target_row)
    rel_badge_text = f"{target_row['관계']} (출계)" if adopt_type == "출계" else str(target_row['관계'])

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
            <div style="font-size:10pt; color:#6b7280; font-family:monospace;">고유ID: {target_row['고유ID']}</div>
        </div>
        
        <div style="font-weight:800; color:#1e3c72; margin-bottom:4px; font-size:11pt;">
            📌 인물 한글 약력
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

def get_siblings_by_father_gid(person_row, all_df):
    """
    동명이인 부친과 혼동되지 않도록 부친의 고유ID(_father_gid)가 동일한 형제들만 정확히 추출
    """
    f_gid = str(person_row.get("_father_gid", "")).strip()
    if f_gid:
        return all_df[all_df["_father_gid"] == f_gid]
    f_name = str(person_row.get("부친", "")).strip()
    if f_name == "覯구":
        return all_df[all_df["부친"] == "覯구"]
    return pd.DataFrame([person_row])

def render_tree_hierarchy(target_row, all_df, max_up=3, max_down=2, key_prefix="tree"):
    """
    고유ID(_father_gid) 검증을 통해 동명이인 선조 오류 없이 정확한 직계 선대·하대 계보도 생성
    """
    target_gid = target_row['고유ID']
    target_name = target_row['인명']
    
    # 1. 선대 조상 탐색 (1 ~ max_up대): 고유ID(_father_gid) 기반으로 정확히 추적
    ancestors = []
    curr = target_row
    for up in range(1, max_up + 1):
        f_gid = str(curr.get('_father_gid', '')).strip()
        if not f_gid:
            break
        f_match = all_df[all_df['고유ID'] == f_gid]
        if f_match.empty:
            break
        f_row = f_match.iloc[0]
        sibs = get_siblings_by_father_gid(f_row, all_df)
        ancestors.append({"level": up, "row": f_row, "siblings": sibs})
        curr = f_row

    # ⭐ 계보 검색 결과창 제목: '계보도' [한글한자 병기 인물명] 표시 후 줄바꿈하여 '(네모 박스 인물 클릭 시 새 계보 생성)' 표시
    person_label = format_ko_hanja_attached(target_name)
    diagram_parts = [
        f'<div id="search-result-anchor"></div>'
        f'<div class="tree-diagram-wrapper">'
        f'<div style="font-weight:800; color:#1e3c72; font-size:12pt; line-height:1.35; margin-bottom:6px;">'
        f'🌳 계보도 [{person_label}]<br>'
        f'<span style="font-size:9.5pt; font-weight:500; color:#475569;">(네모 박스 인물 클릭 시 새 계보 생성)</span>'
        f'</div>'
    ]
    
    for idx_a, anc in enumerate(reversed(ancestors)):
        lvl = anc['level']
        r = anc['row']
        g_num, _ = extract_generation_info(r['고유ID'])
        sibs = anc['siblings']
        
        if idx_a > 0:
            diagram_parts.append('<div class="tree-connector-line">│<br>▼</div>')
            
        if lvl >= 3:
            rel_str = f" ({len(sibs)}자 중 {r['관계']})" if len(sibs) > 1 else f" ({r['관계']})"
            lbl_badge = "증조부" if lvl == 3 else ("고조부" if lvl == 4 else f"{lvl}대조")
            r_gid = r['고유ID']
            diagram_parts.append(f"""
            <div class="tree-level-row">
                <span class="tree-gen-badge">▲ {lbl_badge}({g_num}세)</span>
                <div class="tree-nodes-container">
                    <a href="?tree_gid={r_gid}" target="_self" class="tree-node-box tree-node-target"><b>{r['인명']}</b>{rel_str}</a>
                </div>
            </div>
            """)
        else:
            lbl_badge = "부친(父)" if lvl == 1 else "조부(祖)"
            node_htmls = []
            for s_i, (_, s) in enumerate(sibs.iterrows()):
                is_direct = (s['고유ID'] == r['고유ID'])
                cls = "tree-node-box tree-node-target" if is_direct else "tree-node-box"
                s_name = s['인명']
                s_rel = s['관계']
                s_gid = s['고유ID']
                if s_i > 0:
                    node_htmls.append('<span class="tree-h-line">──</span>')
                node_htmls.append(f'<a href="?tree_gid={s_gid}" target="_self" class="{cls}"><b>{s_name}</b>({s_rel})</a>')
            nodes_joined = ''.join(node_htmls)
            diagram_parts.append(f"""
            <div class="tree-level-row">
                <span class="tree-gen-badge">▲ {lbl_badge}({g_num}세)</span>
                <div class="tree-nodes-container">{nodes_joined}</div>
            </div>
            """)

    if len(ancestors) > 0:
        diagram_parts.append('<div class="tree-connector-line">│<br>▼</div>')

    # 2. 본인 및 형제 행: 부친의 고유ID(_father_gid)가 일치하는 친형제만 표시
    t_sibs = get_siblings_by_father_gid(target_row, all_df)
    tg_num, _ = extract_generation_info(target_row['고유ID'])
    
    t_nodes = []
    for s_i, (_, s) in enumerate(t_sibs.iterrows()):
        is_t = (s['고유ID'] == target_row['고유ID'])
        cls = "tree-node-box tree-node-target" if is_t else "tree-node-box"
        marker = " 🎯[기준]" if is_t else ""
        s_name = s['인명']
        s_rel = s['관계']
        s_gid = s['고유ID']
        if s_i > 0:
            t_nodes.append('<span class="tree-h-line">──</span>')
        t_nodes.append(f'<a href="?tree_gid={s_gid}" target="_self" class="{cls}"><b>{s_name}</b>{marker}({s_rel})</a>')
    t_joined = ''.join(t_nodes)
    diagram_parts.append(f"""
    <div class="tree-level-row" style="background:#eff6ff; border:2px solid #2563eb;">
        <span class="tree-gen-badge" style="background:#1d4ed8; color:white;">● 본인({tg_num}세)</span>
        <div class="tree-nodes-container">{t_joined}</div>
    </div>
    """)

    # 3. 하대 후손 탐색 (1 ~ max_down대): 부모의 고유ID(_father_gid == p_gid)가 일치하는 자녀만 추출
    down_labels = {
        1: "자(子)",
        2: "손(孫)",
        3: "증손(曾孫)",
        4: "현손(玄孫)",
        5: "5대손",
        6: "6대손"
    }
    current_parents = [target_row]
    for down_lvl in range(1, max_down + 1):
        next_gen_rows = []
        node_htmls = []
        for p_row in current_parents:
            p_gid = p_row['고유ID']
            p_name = p_row['인명']
            children = all_df[all_df['_father_gid'] == p_gid]
            for _, ch in children.iterrows():
                next_gen_rows.append(ch)
                ch_name = ch['인명']
                ch_rel = ch['관계']
                ch_gid = ch['고유ID']
                if len(node_htmls) > 0:
                    node_htmls.append('<span class="tree-h-line">──</span>')
                if down_lvl == 1:
                    node_htmls.append(f'<a href="?tree_gid={ch_gid}" target="_self" class="tree-node-box"><b>{ch_name}</b>({ch_rel})</a>')
                else:
                    node_htmls.append(f'<a href="?tree_gid={ch_gid}" target="_self" class="tree-node-box"><b>{ch_name}</b>({p_name}의 {ch_rel})</a>')
        if not next_gen_rows:
            break
        diagram_parts.append('<div class="tree-connector-line">│<br>▼</div>')
        dg_num = (tg_num + down_lvl) if tg_num else (24 + down_lvl)
        d_badge = down_labels.get(down_lvl, f"{down_lvl}대손")
        d_joined = ''.join(node_htmls)
        diagram_parts.append(f"""
        <div class="tree-level-row">
            <span class="tree-gen-badge">▼ {d_badge}({dg_num}세)</span>
            <div class="tree-nodes-container">{d_joined}</div>
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

# 타이틀 상자 출력
render_html(f"""
<div class="title-container">
    <div class="title-left-icon">
        {icon_img_tag}
    </div>
    <div class="title-right-content">
        <div class="title-clan-name">광주안씨 완귀공파 종중</div>
        <div class="title-main-center">족보 검색 시스템</div>
        <div class="title-sub-bottom">- 완귀공(휘 嶒, 24세) ~ 44세<br>총 2,887인(고위 기준) -</div>
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
# 5. 검색 결과 창 (엔터 시 자판 숨김 + 상단 스크롤 + 과거 통합 키워드 검색 적용)
# =============================================================================
if search_type == "인명 검색" and query_val.strip():
    q = query_val.strip()
    matched_df = df[df.apply(lambda r: match_name_query(r, q), axis=1)]
    
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
            candidate_labels = [format_candidate_label(row) for _, row in matched_df.iterrows()]
            selected_label = st.selectbox("선조 선택", candidate_labels, label_visibility="collapsed")
            selected_gid = selected_label.split("고유ID: ")[-1].strip()
            target_row = matched_df[matched_df['고유ID'] == selected_gid].iloc[0]
        else:
            target_row = matched_df.iloc[0]
            
        render_person_result_card(target_row)
        scroll_to_search_top()

elif search_type == "계보 검색":
    q = query_val.strip()
    if q != st.session_state.get('last_lineage_query', ''):
        st.session_state['last_lineage_query'] = q
        st.session_state['active_tree_gid'] = None

    tree_target = None
    if st.session_state.get('active_tree_gid'):
        g_match = df[df['고유ID'] == st.session_state['active_tree_gid']]
        if not g_match.empty:
            tree_target = g_match.iloc[0]

    if tree_target is None and q:
        l_matches = df[df.apply(lambda r: match_name_query(r, q), axis=1)]
        if len(l_matches) == 0:
            render_html(f"""
            <div class="no-result-notice">
                계보를 작성할 선조 인명('{q}')을 찾을 수 없습니다.
            </div>
            """)
            scroll_to_search_top()
        elif len(l_matches) > 1:
            render_html(f"""
            <div class="multi-ancestor-banner">
                📢 동명이인 또는 복수 후보가 총 {len(l_matches)}위 검색되었습니다. 아래에서 계보 기준 인물을 선택하세요.
            </div>
            """)
            cand_opts = [format_candidate_label(r) for _, r in l_matches.iterrows()]
            sel_cand = st.selectbox("계보 기준 인물 선택", cand_opts, label_visibility="collapsed")
            target_gid = sel_cand.split("고유ID: ")[-1].strip()
            tree_target = l_matches[l_matches['고유ID'] == target_gid].iloc[0]
        else:
            tree_target = l_matches.iloc[0]

    if tree_target is not None:
        c_cond1, c_cond2 = st.columns(2)
        with c_cond1:
            render_html("""
            <div class="slider-custom-header">
                <div class="slider-title-text">직계 상대 선조수 선택</div>
                <div class="slider-ticks-row">
                    <span class="slider-tick-item">1대</span>
                    <span class="slider-tick-item">2대</span>
                    <span class="slider-tick-item">3대</span>
                    <span class="slider-tick-item">4대</span>
                    <span class="slider-tick-item">5대</span>
                    <span class="slider-tick-item">6대</span>
                </div>
            </div>
            """)
            up_gens = st.slider("직계 상대 선조수 선택", min_value=1, max_value=6, value=3, format="", label_visibility="collapsed")
        with c_cond2:
            render_html("""
            <div class="slider-custom-header">
                <div class="slider-title-text">직계 하대 선조수 선택</div>
                <div class="slider-ticks-row">
                    <span class="slider-tick-item">1대</span>
                    <span class="slider-tick-item">2대</span>
                    <span class="slider-tick-item">3대</span>
                    <span class="slider-tick-item">4대</span>
                    <span class="slider-tick-item">5대</span>
                    <span class="slider-tick-item">6대</span>
                </div>
            </div>
            """)
            down_gens = st.slider("직계 하대 선조수 선택", min_value=1, max_value=6, value=2, format="", label_visibility="collapsed")

        if st.session_state.get('active_tree_gid'):
            if st.button("🔄 초기 검색 인물로 복귀", use_container_width=True):
                st.session_state['active_tree_gid'] = None
                if hasattr(st, "query_params") and "tree_gid" in st.query_params:
                    del st.query_params["tree_gid"]
                st.rerun()

        render_tree_hierarchy(tree_target, df, max_up=up_gens, max_down=down_gens, key_prefix="main_tree")
        scroll_to_search_top()

elif search_type == "단어 검색" and query_val.strip():
    term = query_val.strip()

    # ⭐ 과거 통합 검색어 중 1단어라도 입력하면 열거된 모든 과거 관련 단어를 통합 검색
    if term in EXAM_GROUP_KEYWORDS:
        search_terms_list = EXAM_GROUP_KEYWORDS
        regex_pattern = "|".join(re.escape(k) for k in EXAM_GROUP_KEYWORDS)
        banner_term_label = f"'{term}' (과거·문과·급제·사마·진사·생원 등 과거 관련 통합 검색)"
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
    word_matches = df[mask]
    
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
        candidate_list = [format_candidate_label(r) for _, r in word_matches.iterrows()]
        chosen = st.selectbox("조회할 선조를 선택하세요", candidate_list, label_visibility="collapsed")
        chosen_gid = chosen.split("고유ID: ")[-1].strip()
        chosen_row = word_matches[word_matches['고유ID'] == chosen_gid].iloc[0]
        
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
            <div class="h-card-name">2. 영천 육우당(六友堂) <span>↗</span></div>
            <div class="h-card-desc">경상북도 영천시 도남동 566 (2012.09.17. 지정)</div>
        </a>
        <a href="https://www.heritage.go.kr/heri/cul/culSelectDetail.do?pageNo=1_1_1_1&sngl=N&ccbaAsno=0007220000000&ccbaCpno=3412603100001" target="_blank" class="heritage-sub-card">
            <div class="h-card-name">3. 청도 만회정(晩悔亭) <span>↗</span></div>
            <div class="h-card-desc">경북 청도군 운문면 지촌리 907 (2026.02.19. 지정)</div>
        </a>
        <div class="heritage-sub-card">
            <div class="h-card-name">4. 삼괴당(三槐堂)</div>
            <div class="h-card-desc">광주안씨 문중 선조 강학 및 추모 유적</div>
        </div>
    </div>
</div>
""")

# =============================================================================
# 7. 하단 푸터 (Footer)
# =============================================================================
render_html("""
<div style="text-align:center; padding:8px 4px 6px 4px; color:#ffffff; font-size:9.5pt; font-weight:600; border-top:1px solid #475569; margin-top:8px; line-height:1.45; text-shadow: 0 1px 2px rgba(0,0,0,0.7);">
    데이터 출처: 2008년 발간 시직공파보 PDF파일 변환후 데이터베이스(완귀공파  고위 기준 총 2,887인)화 &lt;연락처 안동수 010 5097 6403&gt;
</div>
""")
