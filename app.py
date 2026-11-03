import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import sqlite3
import re
import os
import base64

# =============================================================================
# 0. 기본 설정 및 파일 경로
# =============================================================================
ICON_FILE = "icon.png"
DB_FILE = "genealogy.db"
EXCEL_FILE = "genealogy.xlsx"

# 저장소 업로드 배경 이미지 파일 탐색
BG_PATHS = [
    "background.png",
    "background.jpg",
    "bg.png",
    "bg.jpg",
    "artifacts/file_generation/ttl=63d/output/background.png",
    "artifacts/file_generation/ttl=63d/output/bg.png"
]
bg_file = next((p for p in BG_PATHS if os.path.exists(p)), None)

# 아이콘 및 데이터 파일 탐색
if not os.path.exists(ICON_FILE):
    alt_icon = "artifacts/file_generation/ttl=63d/output/icon.png"
    if os.path.exists(alt_icon):
        ICON_FILE = alt_icon

DB_PATHS = [
    "genealogy.db",
    "artifacts/file_generation/ttl=63d/output/genealogy.db"
]
db_file_path = next((p for p in DB_PATHS if os.path.exists(p)), DB_FILE)

EXCEL_PATHS = [
    "genealogy.xlsx",
    "완귀공파 족보 검색 시스템.xlsx",
    "artifacts/file_generation/ttl=63d/output/genealogy.xlsx"
]
excel_file_path = next((p for p in EXCEL_PATHS if os.path.exists(p)), EXCEL_FILE)

# ⭐ st.set_page_config는 항상 최상단에서 가장 먼저 호출
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

def scroll_to_search_top():
    """
    ⭐ 검색 엔터 후:
    1) 모바일 가상 자판(키보드)이 즉시 숨겨지도록 입력창 포커스 해제(blur)
    2) 인명검색 라인이 화면 제일 위에 오도록 상단으로 자동 스크롤
    """
    components.html(
        """
        <script>
            setTimeout(function() {
                try {
                    var parentDoc = window.parent.document;
                    if (parentDoc.activeElement) {
                        parentDoc.activeElement.blur();
                    }
                    var inputs = parentDoc.querySelectorAll('input');
                    inputs.forEach(function(inp) {
                        inp.blur();
                    });
                    var target = parentDoc.getElementById('search-section-anchor');
                    if (target) {
                        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    }
                } catch (e) {}
            }, 80);
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
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700;800;900&family=Nanum+Myeongjo:wght@400;700&display=swap');

html, body, [class*="css"], .stMarkdown, p, div {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', 'Malgun Gothic', sans-serif;
}

/* ⭐ 상단 헤더 및 앱 내부 툴바/푸터/배지 숨김 */
header[data-testid="stHeader"],
div[data-testid="stToolbar"],
div[data-testid="stDecoration"],
div[data-testid="stStatusWidget"],
div[data-testid="stAppDeployButton"],
.stDeployButton,
footer,
#MainMenu,
[class*="viewerBadge"],
[class^=" _profileContainer"],
[class*="_profileContainer"],
[class^=" _container"],
iframe[title="streamlit_badge"] {
    display: none !important;
    visibility: hidden !important;
    height: 0px !important;
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
    gap: 0.25rem !important;
}
div[data-testid="stElementContainer"],
.element-container {
    margin-top: 0px !important;
    margin-bottom: 0px !important;
}

/* 저장소 배경 이미지 적용 및 투명도 제로(100% 완전 불투명) */
.stApp {
    background-color: #162b52;
    background-size: cover !important;
    background-position: center top !important;
    background-attachment: fixed !important;
    opacity: 1 !important;
}

/* ⭐ 제일 위 페이지 타이틀 창: 좌측 초대형 아이콘(118px) + 우측 타이틀 문장 */
.title-container {
    width: 100% !important;
    box-sizing: border-box !important;
    margin-top: 8px !important;
    margin-bottom: 14px !important;
    padding: 10px 14px;
    border-bottom: 3px solid #8c2d19;
    background: #ffffff !important;
    opacity: 1 !important;
    border-radius: 6px;
    box-shadow: 0 3px 8px rgba(0,0,0,0.18);
    display: flex !important;
    flex-direction: row !important;
    align-items: center !important;
    justify-content: space-between !important;
    gap: 12px !important;
}

.title-left-icon {
    flex-shrink: 0 !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
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
    color: #1a202c;
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
    font-weight: 600;
    color: #334155;
    text-align: center;
    width: 100%;
    margin-top: 1px;
    line-height: 1.35;
    letter-spacing: -0.2px;
}

/* ⭐ 위 타이틀 창과 간격을 조금 더 넓힌 안내 문구 라인 */
.compact-search-header {
    width: 100% !important;
    text-align: center !important;
    margin-top: 12px !important;
    margin-bottom: 2px !important;
    padding: 2px 2px !important;
    line-height: 1.3 !important;
}
.compact-search-subtitle {
    font-size: 10.8pt !important;
    font-weight: 700 !important;
    color: #fef08a !important;
    text-align: center !important;
    margin-top: 0px !important;
    margin-bottom: 0px !important;
    line-height: 1.3 !important;
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

/* ⭐ 검색 결과 없을 때 안내 문구: 글자 색상 흰색(#ffffff) */
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

/* ⭐ 직계 상대/하대 선조수 선택 슬라이더:
   1) 문구: '직계 상대 선조수 선택' / '직계 하대 선조수 선택'
   2) 하얀색 대수 글자(1대~6대): 크기 축소(9.2pt), 비볼드체(font-weight: 400)
   3) 빨간 막대 오른쪽 끝부분 위에 표시되는 빨간 숫자(ThumbValue / BaseWeb 툴팁) 완전 삭제 */
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
/* 슬라이더 조절 손잡이(Thumb) 및 내부/상단 빨간 숫자 완전 제거 */
div[data-testid="stSlider"] [role="slider"] {
    width: 20px !important;
    height: 20px !important;
    background-color: #facc15 !important;
    border: 2.5px solid #000000 !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.6) !important;
    font-size: 0px !important;
    color: transparent !important;
}
div[data-testid="stSlider"] [data-testid="stThumbValue"],
div[data-testid="stSlider"] [data-testid="stTickBar"],
div[data-testid="stSlider"] [data-testid="stTickBarMin"],
div[data-testid="stSlider"] [data-testid="stTickBarMax"],
div[data-testid="stSlider"] [role="slider"] > div,
div[data-testid="stSlider"] [role="slider"] *,
div[data-testid="stSlider"] [role="slider"]::before,
div[data-testid="stSlider"] [role="slider"]::after,
div[data-testid="stSlider"] div[data-baseweb="slider"] span,
div[data-testid="stSlider"] div[data-baseweb="slider"] p {
    display: none !important;
    visibility: hidden !important;
    opacity: 0 !important;
    font-size: 0px !important;
    color: transparent !important;
    -webkit-text-fill-color: transparent !important;
    height: 0px !important;
    width: 0px !important;
    overflow: hidden !important;
    pointer-events: none !important;
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

/* 신상 및 계보 정보 */
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

/* 문화재 안내 */
.heritage-sub-container {
    width: 100% !important;
    box-sizing: border-box !important;
    background-color: #ffffff;
    border: 2px solid #cbd5e1;
    border-radius: 8px;
    padding: 9px 12px;
    margin-top: 6px;
    margin-bottom: 6px;
}
.heritage-sub-header {
    font-size: 11pt;
    font-weight: 800;
    color: #1e3c72;
    margin-bottom: 6px;
    line-height: 1.4;
}
.heritage-sub-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 6px;
}
.heritage-sub-card {
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 6px 8px;
    text-decoration: none !important;
    display: block;
}
.heritage-sub-card:hover {
    background: #eff6ff;
    border-color: #2563eb;
}
.h-card-name {
    font-size: 10pt;
    font-weight: 700;
    color: #1e40af;
    margin-bottom: 2px;
    line-height: 1.35;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.h-card-desc {
    font-size: 8.8pt;
    color: #475569;
    line-height: 1.35;
}
</style>
"""

render_html(CUSTOM_CSS)

# =============================================================================
# 2. 데이터 로드 및 고유ID 기반 부모-자녀 정밀 매핑 전처리
# =============================================================================
def parse_gid_tuple(gid):
    """고유ID(예: C-529-4-1-37)를 (권, 페이지, 단, 순서, 세수)로 분리"""
    m = re.match(r"^([A-Z])-(\d+)-(\d+)-(\d+)-(\d+)$", str(gid).strip())
    if m:
        return m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4)), int(m.group(5))
    return "", 0, 0, 0, 0

def resolve_exact_father_gid(child_row, all_df):
    """
    동명이인 선조가 여러 명 있을 때 단순 인명 비교가 아니라
    1) 세수(본인 세수 - 1)
    2) 조부 성함 일치 여부(후보 부친의 '부친' == 본인의 '조부')
    3) 족보 권(A~D)·페이지·순번의 근접성
    을 종합 검증하여 정확한 부친의 고유ID(_father_gid)를 찾아냅니다.
    """
    f_name = str(child_row.get("부친", "")).strip()
    if not f_name or f_name == "覯구" or f_name == "-":
        return ""
    gf_name = str(child_row.get("조부", "")).strip()
    c_vol, c_page, c_tier, c_seq, c_gen = parse_gid_tuple(child_row.get("고유ID", ""))

    # 1차: 인명 일치 + 세수가 정확히 1세대 위(c_gen - 1)인 후보
    cands = all_df[(all_df["인명"] == f_name) & (all_df["_gen"] == c_gen - 1)]
    if cands.empty:
        cands = all_df[all_df["인명"] == f_name]
    if cands.empty:
        return ""
    if len(cands) == 1:
        return str(cands.iloc[0]["고유ID"])

    # 2차: 후보의 부친 성함이 자녀의 조부 성함과 일치하는 후보로 압축
    if gf_name and gf_name != "-":
        cands_gf = cands[cands["부친"] == gf_name]
        if len(cands_gf) == 1:
            return str(cands_gf.iloc[0]["고유ID"])
        elif len(cands_gf) > 1:
            cands = cands_gf

    # 3차: 같은 권(Volume) 내에서 자녀 페이지 이하이면서 가장 가까운 페이지에 수록된 선조 선택
    same_vol = cands[(cands["_vol"] == c_vol) & (cands["_page"] <= c_page)]
    if not same_vol.empty:
        same_vol_sorted = same_vol.sort_values(by=["_page", "_seq"], ascending=[False, True])
        return str(same_vol_sorted.iloc[0]["고유ID"])

    # 4차: 권이 넘어가는 경우(C권 38세 -> D권 39세 등) 본 기록(생년·배위 수록) 우선 선택
    non_empty_bio = cands[cands["한글 약력 정리"].astype(str).str.contains("생|배 :", na=False)]
    if len(non_empty_bio) == 1:
        return str(non_empty_bio.iloc[0]["고유ID"])

    return str(cands.iloc[0]["고유ID"])

@st.cache_data
def load_and_preprocess_genealogy():
    raw_df = None
    status_msg = ""

    if os.path.exists(db_file_path):
        try:
            conn = sqlite3.connect(f"file:{db_file_path}?mode=ro", uri=True)
            raw_df = pd.read_sql("SELECT * FROM genealogy", conn).fillna("")
            conn.close()
            if not raw_df.empty:
                status_msg = "성공 (SQLite DB)"
        except Exception:
            try:
                conn = sqlite3.connect(db_file_path)
                raw_df = pd.read_sql("SELECT * FROM genealogy", conn).fillna("")
                conn.close()
                if not raw_df.empty:
                    status_msg = "성공 (SQLite DB)"
            except Exception:
                pass

    if (raw_df is None or raw_df.empty) and os.path.exists(excel_file_path):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(excel_file_path, read_only=True)
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
        return None, f"데이터 파일을 찾을 수 없습니다. (확인 경로: {excel_file_path})"

    # 고유ID 파싱 및 부친 고유ID(_father_gid) 사전 매핑
    parsed = raw_df["고유ID"].apply(parse_gid_tuple)
    raw_df["_vol"] = [p[0] for p in parsed]
    raw_df["_page"] = [p[1] for p in parsed]
    raw_df["_tier"] = [p[2] for p in parsed]
    raw_df["_seq"] = [p[3] for p in parsed]
    raw_df["_gen"] = [p[4] for p in parsed]

    father_gids = [resolve_exact_father_gid(row, raw_df) for _, row in raw_df.iterrows()]
    raw_df["_father_gid"] = father_gids

    return raw_df, status_msg

df, load_msg = load_and_preprocess_genealogy()

# =============================================================================
# 3. 보조 유틸리티 함수: 세수 표기, 입계/출계 판별, 문집 추출, 한글 신상 문단 생성
# =============================================================================
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

def match_name_query(row, q):
    """
    인명 검색 및 계보 검색 시 안내문은 그대로 유지하되,
    실제 검색은 함자, 성명, 자(字), 초자(初字), 호(號), 일휘(一諱), 초휘(初諱), 초명(初名), 통명(通名)을 모두 검색
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
    if search_term and search_term.strip():
        sterm = search_term.strip()
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
    ⭐ 학력·인품·업적 항목에서 문장이 끝난 경우 마침표(.)를 찍고 줄바꿈(<br>) 처리
    """
    if not text:
        return ""
    s = str(text).strip()
    # 마침표 없이 종결어미로 끝난 뒤 다음 문장이 이어지는 곳에 마침표(.) 자동 부여
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
    신상 및 계보 정보 항목 구성:
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
    h_name = clean_ko_name(target_row['인명'])
    adopt_type, _, _ = get_adoption_info(target_row)
    rel_badge_text = f"{target_row['관계']} (출계)" if adopt_type == "출계" else str(target_row['관계'])

    meta_paragraphs_html = format_hangul_only_paragraph_html(target_row, search_term=search_term)
    cn_raw = highlight_corrections_html(target_row.get('한문 약력 원문', ''), search_term=search_term)

    render_html(f"""
    <div id="search-result-anchor"></div>
    <div class="result-card">
        <div class="result-card-header">
            <div class="result-name-big">
                <span>{h_name if h_name else target_row['인명']}</span>
                <span class="result-badge">{rel_badge_text}</span>
            </div>
            <div style="font-size:10pt; color:#6b7280; font-family:monospace;">고유ID: {target_row['고유ID']}</div>
        </div>
        
        <div style="font-weight:800; color:#1e3c72; margin-bottom:4px; font-size:11pt;">
            📌 신상 및 계보 정보
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

    diagram_parts = [f'<div id="search-result-anchor"></div><div class="tree-diagram-wrapper"><div style="font-weight:800; color:#1e3c72; font-size:11.5pt; margin-bottom:4px;">🌳 [{clean_ko_name(target_name)}] 기준 혈통 계보도 <span style="font-size:9pt; font-weight:400; color:#64748b;">(네모 박스 클릭 시 해당 인물로 즉시 새 계보 생성)</span></div>']
    
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
                is_direct
