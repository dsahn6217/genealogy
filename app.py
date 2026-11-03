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

/* ⭐ 직계 상대/하대 선조수 설정 슬라이더:
   1) 문구: '직계 상대 선조수 설정' / '직계 하대 선조수 설정'
   2) 하얀색 대수 글자(1대~6대): 크기 축소(9.2pt), 비볼드체(font-weight: 400)
   3) 하얀색 라인 위 빨간색 기본 숫자(ThumbValue 등) 완전 삭제 */
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
/* 슬라이더 조절 손잡이(Thumb) */
div[data-testid="stSlider"] [role="slider"] {
    width: 20px !important;
    height: 20px !important;
    background-color: #facc15 !important;
    border: 2.5px solid #000000 !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.6) !important;
}
/* ⭐ 하얀색 라인 위에 뜨는 빨간색 숫자(ThumbValue) 및 기본 TickBar 완전 삭제 */
div[data-testid="stSlider"] [data-testid="stThumbValue"],
div[data-testid="stSlider"] [data-testid="stTickBar"],
div[data-testid="stSlider"] [data-testid="stTickBarMin"],
div[data-testid="stSlider"] [data-testid="stTickBarMax"],
div[data-testid="stSlider"] div[role="slider"] > div,
div[data-testid="stSlider"] div[role="slider"] * {
    display: none !important;
    visibility: hidden !important;
    opacity: 0 !important;
    font-size: 0px !important;
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
# 2. 데이터 로드 (SQLite DB 우선, 없으면 엑셀 로드)
# =============================================================================
@st.cache_data
def load_and_preprocess_genealogy():
    if os.path.exists(db_file_path):
        try:
            conn = sqlite3.connect(f"file:{db_file_path}?mode=ro", uri=True)
            df = pd.read_sql("SELECT * FROM genealogy", conn).fillna("")
            conn.close()
            if not df.empty:
                return df, "성공 (SQLite DB)"
        except Exception:
            try:
                conn = sqlite3.connect(db_file_path)
                df = pd.read_sql("SELECT * FROM genealogy", conn).fillna("")
                conn.close()
                if not df.empty:
                    return df, "성공 (SQLite DB)"
            except Exception:
                pass

    if os.path.exists(excel_file_path):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(excel_file_path, read_only=True)
            sheet_name = "족보전사결과" if "족보전사결과" in wb.sheetnames else wb.sheetnames[0]
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return None, "시트에 데이터가 없습니다."
            header = [str(c).strip() if c is not None else f"Col{idx}" for idx, c in enumerate(rows[0])]
            df = pd.DataFrame(rows[1:], columns=header).fillna("")
            return df, "성공 (Excel)"
        except Exception as e:
            return None, str(e)

    return None, f"데이터 파일을 찾을 수 없습니다. (확인 경로: {excel_file_path})"

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
    인명 검색 및 계보 검색 시 함자, 성명, 호, 초명, 통명만 정확하게 검색
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
        
    # 2. 호, 초명, 통명 (한글 약력 정리 첫 줄)
    first_line_ko = str(row.get('한글 약력 정리', '')).split('\n')[0]
    for pattern in [r'호는\s*([^\.,\n\s]+)', r'초명은?\s*([^\.,\n\s]+)', r'통명은?\s*([^\.,\n\s]+)']:
        m = re.search(pattern, first_line_ko)
        if m:
            val = m.group(1).strip()
            if q == val or q_alt == val or q in val:
                return True

    # 3. 호, 초명, 통명 (한문 약력 정리 첫 줄)
    first_line_cn = str(row.get('한문 약력 정리', '')).split('\n')[0]
    for pattern in [r'號\s*([^\.,\n\s]+)', r'初名\s*([^\.,\n\s]+)', r'通名\s*([^\.,\n\s]+)']:
        m = re.search(pattern, first_line_cn)
        if m:
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

def parse_person_hangul_details(row):
    """
    ⭐ 신상 및 계보 정보 항목 구성:
    1) 부, 조부 통합 -> '부친·조부' 항목 (부친 00, 조부 00)
    2) 자, 호 통합 -> '자·호' 항목 (자 00, 호 00 형식)
    3) '학력·인품·업적' 바로 위에 '문집' 항목 배치
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

    # ⭐ '학력·인품·업적' 바로 위에 '문집' 항목 삽입
    munjip_val = extract_munjip_info(ko_bio)
    if munjip_val:
        details.append(("문집", munjip_val))

    if hakmun_val:
        details.append(("학력·인품·업적", hakmun_val))

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

def render_tree_hierarchy(target_row, all_df, max_up=3, max_down=2, key_prefix="tree"):
    """
    계보도 도형 박스 클릭 시 새 계보 즉시 생성 (선대 1~6대, 하대 1~6대 지원)
    """
    target_gid = target_row['고유ID']
    target_name = target_row['인명']
    
    # 1. 선대 조상 탐색 (1 ~ max_up대)
    ancestors = []
    curr = target_row
    for up in range(1, max_up + 1):
        f_name = str(curr.get('부친', '')).strip()
        if not f_name or f_name == '覯구':
            break
        f_match = all_df[all_df['인명'] == f_name]
        if f_match.empty:
            break
        f_row = f_match.iloc[0]
        ff_name = str(f_row.get('부친', '')).strip()
        sibs = all_df[all_df['부친'] == ff_name] if ff_name else pd.DataFrame([f_row])
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

    # 2. 본인 및 형제 행
    t_f = str(target_row.get('부친', '')).strip()
    t_sibs = all_df[all_df['부친'] == t_f] if t_f else pd.DataFrame([target_row])
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

    # 3. 하대 후손 탐색 (1 ~ max_down대까지 동적 지원)
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
            p_name = p_row['인명']
            children = all_df[all_df['부친'] == p_name]
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
# 4. 최상단 헤더(좌측 초대형 아이콘 + 우측 타이틀) 및 검색 인터페이스
# =============================================================================
# ⭐ 아이콘 크기 추가 확대 (96px -> 118px)
icon_img_tag = f'<img src="{icon_b64}" style="width:118px; height:118px; object-fit:contain; display:block;">' if icon_b64 else '<span style="font-size:68px;">📜</span>'

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
    st.info(f"데이터베이스 파일(`genealogy.db` 또는 `genealogy.xlsx`)이 저장소에 존재하는지 확인해 주시기 바랍니다.")
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

# 검색어 안내문 표시 및 입력창 (테두리 흰색, 세로 중앙 정렬)
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
# 5. 검색 결과 창 (검색 실패 안내문 흰색 글자 적용)
# =============================================================================
if search_type == "인명 검색" and query_val.strip():
    q = query_val.strip()
    matched_df = df[df.apply(lambda r: match_name_query(r, q), axis=1)]
    
    if len(matched_df) == 0:
        render_html(f"""
        <div class="no-result-notice">
            입력하신 검색어 '{q}'에 해당하는 인물을 찾을 수 없습니다. (함자, 성명, 호, 초명, 통명)
        </div>
        """)
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
                계보를 작성할 선조 인명('{q}')을 찾을 수 없습니다. (함자, 성명, 호, 초명, 통명)
            </div>
            """)
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
                <div class="slider-title-text">직계 상대 선조수 설정</div>
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
            up_gens = st.slider("직계 상대 선조수 설정", min_value=1, max_value=6, value=3, label_visibility="collapsed")
        with c_cond2:
            render_html("""
            <div class="slider-custom-header">
                <div class="slider-title-text">직계 하대 선조수 설정</div>
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
            down_gens = st.slider("직계 하대 선조수 설정", min_value=1, max_value=6, value=2, label_visibility="collapsed")

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
    mask = (
        df['한글 약력 정리'].astype(str).str.contains(term, case=False, na=False) |
        df['한문 약력 정리'].astype(str).str.contains(term, case=False, na=False) |
        df['한글 약력 원문'].astype(str).str.contains(term, case=False, na=False) |
        df['한문 약력 원문'].astype(str).str.contains(term, case=False, na=False)
    )
    word_matches = df[mask]
    
    if len(word_matches) == 0:
        render_html(f"""
        <div class="no-result-notice">
            입력하신 단어 '{term}'을(를) 포함하고 있는 약력 기록이 없습니다.
        </div>
        """)
    else:
        render_html(f"""
        <div class="multi-ancestor-banner">
            📢 검색어 <span class="search-word-red">'{term}'</span>을(를) 포함하는 약력을 가진 선조가 총 {len(word_matches):,}위(다수) 검색되었습니다. 아래 목록에서 선조를 선택하세요.
        </div>
        """)
        candidate_list = [format_candidate_label(r) for _, r in word_matches.iterrows()]
        chosen = st.selectbox("조회할 선조를 선택하세요", candidate_list, label_visibility="collapsed")
        chosen_gid = chosen.split("고유ID: ")[-1].strip()
        chosen_row = word_matches[word_matches['고유ID'] == chosen_gid].iloc[0]
        
        render_person_result_card(chosen_row, search_term=term)
        scroll_to_search_top()

# =============================================================================
# 6. 종중 주요 유적 및 문화재 안내 (항상 검색 결과 창 아래에 위치)
# =============================================================================
render_html("""
<div class="heritage-sub-container">
    <div class="heritage-sub-header">🏛️ 종중 주요 유적 및 문화재 안내 (클릭 시 국가유산 포털로 이동)</div>
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
