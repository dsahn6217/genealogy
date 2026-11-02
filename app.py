import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import re
import os
import base64

# =============================================================================
# 0. 기본 설정 및 아이콘 / 데이터 / 배경 파일 경로
# =============================================================================
ICON_FILE = "icon.png"
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

if not os.path.exists(DATA_FILE):
    alt_data = "genealogy.xlsx"
    if os.path.exists(alt_data):
        DATA_FILE = alt_data
    elif os.path.exists("완귀공파 족보 검색 시스템.xlsx"):
        DATA_FILE = "완귀공파 족보 검색 시스템.xlsx"
    elif os.path.exists("artifacts/file_generation/ttl=63d/output/genealogy.xlsx"):
        DATA_FILE = "artifacts/file_generation/ttl=63d/output/genealogy.xlsx"

@st.cache_data
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

def scroll_to_result():
    """검색 완료 시 화면이 즉시 검색 결과창 위치로 1회 스크롤되도록 처리 (모바일 부하 없음)"""
    components.html(
        """
        <script>
            setTimeout(function() {
                try {
                    var target = window.parent.document.getElementById('search-result-anchor');
                    if (target) {
                        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    }
                } catch (e) {}
            }, 100);
        </script>
        """,
        height=0
    )

# =============================================================================
# 1. 페이지 설정 및 디자인 스타일 CSS (타이틀 좌측 대형 아이콘, 줄 겹침 해소, 검은색 테두리)
# =============================================================================
st.set_page_config(
    page_title="광주안씨 완귀공파 종중 족보 검색 시스템",
    page_icon=ICON_FILE if os.path.exists(ICON_FILE) else "📜",
    layout="wide",
    initial_sidebar_state="collapsed"
)

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700;800;900&family=Nanum+Myeongjo:wght@400;700&display=swap');

html, body, [class*="css"], .stMarkdown, p, div {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', 'Malgun Gothic', sans-serif;
}

/* ⭐ 모바일 부하를 유발하는 JS 루프 없이 순수 CSS로만 상단 헤더·푸터·툴바 숨김 */
header[data-testid="stHeader"],
#MainMenu,
footer,
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
[data-testid="manage-app-button"],
.stAppDeployButton,
.viewerBadge_container__r5tak,
.viewerBadge_link__qRIco,
.styles_viewerBadge__CvC9N {
    display: none !important;
    visibility: hidden !important;
}

/* ⭐ 화면의 모든 창 가로 정렬은 중앙에 배치 */
.main .block-container, 
[data-testid="stAppViewBlockContainer"], 
.block-container {
    max-width: 100% !important;
    width: 100% !important;
    margin-left: auto !important;
    margin-right: auto !important;
    padding-top: 1.1rem !important;
    padding-bottom: 1.2rem !important;
    padding-left: 0.65rem !important;
    padding-right: 0.65rem !important;
    box-sizing: border-box !important;
}

div[data-testid="stVerticalBlock"] {
    gap: 0.42rem !important;
    align-items: center !important;
    width: 100% !important;
}
div[data-testid="stElementContainer"],
.element-container,
.stMarkdown {
    width: 100% !important;
    margin-left: auto !important;
    margin-right: auto !important;
}

/* 저장소 배경 이미지 적용 및 투명도 제로(100% 완전 불투명) */
.stApp {
    background-color: #162b52;
    background-size: cover !important;
    background-position: center top !important;
    background-attachment: scroll !important;
    opacity: 1 !important;
}

/* ⭐ 제일 위 페이지 타이틀 창: 왼쪽에 아이콘을 세로 중앙으로 크게 배치하고 타이틀 문장은 오른쪽에 배치 */
.title-container {
    width: 100% !important;
    box-sizing: border-box !important;
    margin-top: 10px !important;
    margin-bottom: 6px !important;
    margin-left: auto !important;
    margin-right: auto !important;
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

/* ⭐ '다음 중 하나를 선택하고 검색어를 한글로 입력'과 아래 '인명 검색' 라인이 겹치지 않도록 적당한 간격 확보 */
.compact-search-header {
    width: 100% !important;
    text-align: center !important;
    margin-top: 6px !important;
    margin-bottom: 6px !important;
    padding: 2px 2px !important;
    line-height: 1.45 !important;
}
.compact-search-subtitle {
    font-size: 12pt !important;
    font-weight: 800 !important;
    color: #fef08a !important;
    text-align: center !important;
    margin-top: 0px !important;
    margin-bottom: 2px !important;
    line-height: 1.45 !important;
    letter-spacing: -0.4px !important;
    text-shadow: 0 1px 3px rgba(0,0,0,0.65);
}

/* ⭐ 검색 유형 선택 라인: 겹침 없이 적당한 상·하단 간격 유지 + 가로 중앙 정렬 */
div[data-testid="stRadio"] {
    width: 100% !important;
    margin-top: 4px !important;
    margin-bottom: 4px !important;
    padding-top: 2px !important;
    padding-bottom: 2px !important;
    display: flex !important;
    justify-content: center !important;
}
div[data-testid="stRadio"] > div {
    width: 100% !important;
    display: flex !important;
    justify-content: center !important;
    margin: 0px !important;
    padding: 0px !important;
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
    padding: 2px 2px !important;
    overflow-x: auto !important;
}
div[role="radiogroup"] > label,
div[role="radiogroup"] label[data-baseweb="radio"] {
    white-space: nowrap !important;
    flex-shrink: 0 !important;
    margin-right: 2px !important;
    margin-bottom: 0px !important;
    padding: 2px 2px !important;
    line-height: 1.4 !important;
}
div[role="radiogroup"] label p,
div[role="radiogroup"] label span {
    font-size: 12.5pt !important;
    font-weight: 800 !important;
    color: #ffffff !important;
    white-space: nowrap !important;
    line-height: 1.4 !important;
    letter-spacing: -0.3px !important;
    text-shadow: 0 1px 3px rgba(0,0,0,0.65);
}

/* ⭐ 입력 검색 안내 창: 아래에 있는 검색어 입력 창과의 간격을 넓힘 */
.search-guide-box {
    width: 100% !important;
    box-sizing: border-box !important;
    margin-left: auto !important;
    margin-right: auto !important;
    font-size: 10pt !important;
    font-weight: 400 !important;
    color: #0f172a !important;
    background-color: #fef9c3;
    padding: 7px 10px;
    border-radius: 6px;
    border: 2px solid #eab308;
    line-height: 1.48 !important;
    margin-top: 2px !important;
    margin-bottom: 14px !important;
}

/* ⭐ 검색어 입력 창: 테두리를 까만색(#000000)으로 변경하고 위쪽 안내 창과의 간격 확보 */
div[data-testid="stTextInput"] {
    width: 100% !important;
    margin-left: auto !important;
    margin-right: auto !important;
    margin-top: 6px !important;
    margin-bottom: 6px !important;
}
div[data-testid="stTextInput"] div[data-baseweb="input"] {
    border: 3.5px solid #000000 !important;
    border-radius: 8px !important;
    background-color: #ffffff !important;
    box-shadow: 0 2px 6px rgba(0,0,0,0.25) !important;
    box-sizing: border-box !important;
}
div[data-testid="stTextInput"] div[data-baseweb="base-input"] {
    border: none !important;
    background-color: transparent !important;
}
div[data-testid="stTextInput"] input {
    color: #4b5563 !important;
    font-weight: 700 !important;
    font-size: 12pt !important;
    padding: 7px 10px !important;
    line-height: 1.4 !important;
    background-color: #ffffff !important;
}
div[data-testid="stTextInput"] input::placeholder {
    color: #888888 !important;
    font-weight: 400 !important;
    opacity: 1 !important;
}
div[data-testid="stTextInput"] input::-webkit-input-placeholder {
    color: #888888 !important;
    font-weight: 400 !important;
    opacity: 1 !important;
}

/* ⭐ 검색어 입력창 하단에 생성되는 결과창 안내문: 파란색 글씨 + 비볼딕체(가늘게) + 가로 중앙 정렬 */
.multi-ancestor-banner {
    width: 100% !important;
    box-sizing: border-box !important;
    margin-left: auto !important;
    margin-right: auto !important;
    background-color: #eff6ff !important;
    border: 1.5px solid #3b82f6 !important;
    color: #1d4ed8 !important;
    padding: 7px 10px;
    border-radius: 6px;
    font-size: 11pt !important;
    font-weight: 400 !important;
    margin-top: 4px !important;
    margin-bottom: 8px !important;
    line-height: 1.48 !important;
}

/* ⭐ 검색창 명부(선택 드롭다운): 글자체를 볼딕체(굵은 글씨)로 지정 및 가로 중앙 정렬 */
div[data-testid="stSelectbox"] {
    width: 100% !important;
    margin-left: auto !important;
    margin-right: auto !important;
    margin-top: 2px !important;
    margin-bottom: 6px !important;
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

/* 검색 결과 인물 카드 (가로 중앙 정렬 및 행간 확보) */
.result-card {
    width: 100% !important;
    box-sizing: border-box !important;
    margin-left: auto !important;
    margin-right: auto !important;
    background-color: #ffffff;
    border: 2px solid #1e3c72;
    border-top: 4px solid #1e3c72;
    border-radius: 8px;
    padding: 10px 12px;
    margin-top: 6px !important;
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

/* 신상 및 계보 정보 (문장이 겹치지 않도록 행간 확보) */
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

/* ⭐ 한문 약력 원문 전용 컨테이너 */
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

/* 단어 검색어 빨간색 강조 */
.search-word-red {
    color: #ff0000 !important;
    font-weight: 700 !important;
    background-color: #fee2e2;
    padding: 0px 4px;
    border-radius: 3px;
    border: 1px solid #f87171;
}

/* 계보도 도형 간 간격 및 연결선 표시 */
.tree-diagram-wrapper {
    width: 100% !important;
    box-sizing: border-box !important;
    margin-left: auto !important;
    margin-right: auto !important;
    background: #ffffff;
    border: 2px solid #1e3c72;
    border-radius: 8px;
    padding: 8px 10px;
    margin-top: 6px;
    margin-bottom: 8px;
}
.tree-level-row {
    display: flex;
    align-items: center;
    margin-bottom: 0px !important;
    padding: 5px 6px;
    background: #f8fafc;
    border-radius: 5px;
    border: 1px solid #cbd5e1;
    line-height: 1.4;
}
.tree-connector-line {
    text-align: center;
    color: #1e3c72;
    font-weight: 900;
    font-size: 10pt;
    line-height: 1.05;
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
    background: #ffffff;
    border: 1.5px solid #64748b;
    border-radius: 4px;
    padding: 2px 7px;
    font-size: 9.5pt;
    color: #0f172a;
    line-height: 1.3;
}
.tree-node-target {
    background: #eff6ff;
    border: 2px solid #2563eb;
    font-weight: 800;
    color: #1e40af;
}
.tree-h-line {
    color: #64748b;
    font-weight: 700;
    font-size: 8.5pt;
    margin: 0 1px;
}

/* 종중 주요 유적 및 문화재 안내 (항상 검색 결과창 아래에 위치, 가로 중앙 정렬) */
.heritage-sub-container {
    width: 100% !important;
    box-sizing: border-box !important;
    margin-left: auto !important;
    margin-right: auto !important;
    background-color: #ffffff;
    border: 2px solid #cbd5e1;
    border-radius: 8px;
    padding: 9px 12px;
    margin-top: 8px;
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
# 2. 데이터 로드 및 전처리 캐싱
# =============================================================================
@st.cache_data
def load_and_preprocess_genealogy():
    if not os.path.exists(DATA_FILE):
        return None, f"데이터 파일({DATA_FILE})을 찾을 수 없습니다."
    try:
        wb = openpyxl.load_workbook(DATA_FILE, read_only=True)
        sheet_name = "족보전사결과" if "족보전사결과" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return None, "시트에 데이터가 없습니다."
        header = [str(c).strip() if c is not None else f"Col{idx}" for idx, c in enumerate(rows[0])]
        df = pd.DataFrame(rows[1:], columns=header).fillna("")
        return df, "성공"
    except Exception as e:
        return None, str(e)

df, load_msg = load_and_preprocess_genealogy()

# =============================================================================
# 3. 보조 유틸리티 함수: 세수 표기, 입계/출계 판별, 한글 신상 문단 생성, 계보도
# =============================================================================
def extract_generation_info(gid):
    """
    ⭐ '[0세  완귀공의 00대손]' 형식으로 표기하기 위한 세수 문자열 생성
    """
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
    """한문을 제거하고 순수 한글 인명만 추출"""
    s = str(raw_val).strip()
    ko, _ = split_name(s)
    if ko:
        return ko
    cleaned = re.sub(r"[一-龥]+", "", s).strip()
    return cleaned if cleaned else s

def get_adoption_info(row):
    """
    ⭐ 입계자 / 출계자 / 일반 판별 및 계부·생부·양부 한글 성명 추출
    - 입계자: 관계 == '계자' 또는 한글 약력 정리에 '생부 00의' 포함 -> ('입계', 계부(부친컬럼), 생부)
    - 출계자: 한글 약력 정리에 '양부 00의 계자' 또는 원문에 '출계'/'出系' 포함 -> ('출계', 생부(부친컬럼), 양부)
    - 일반: ('일반', 부친, '')
    """
    rel = str(row.get("관계", "")).strip()
    ko_bio = str(row.get("한글 약력 정리", "")).strip()
    ko_raw = str(row.get("한글 약력 원문", "")).strip()
    cn_raw = str(row.get("한문 약력 원문", "")).strip()
    father_ko = clean_ko_name(row.get("부친", ""))

    # 1. 입계자 판별
    if rel == "계자" or "생부 " in ko_bio:
        m = re.search(r"생부\s*([^\s의\.]+)의?", ko_bio)
        saengbu_ko = clean_ko_name(m.group(1)) if m else ""
        return "입계", father_ko, saengbu_ko

    # 2. 출계자 판별
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
    """
    ⭐ 검색 결과 창 명부(선택 목록) 표기 형식:
    - '[0세  완귀공의 00대손]' 표기
    - 입계자인 경우: '[0세  완귀공의 00대손] 함자  계자 (계부:00, 조부:00)'
    - 출계자인 경우: '[0세  완귀공의 00대손] 함자  0자 (출계) (생부:00, 조부:00)'
    - 일반인 경우:   '[0세  완귀공의 00대손] 함자  0자 (부:00, 조부:00)'
    """
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
        "1775년", "1775", "一七七五", "1898년", "1898", "一八九八",
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

def parse_person_hangul_details(row):
    """
    ⭐ '신상 및 계보 정보' 한글 전용 항목 생성:
    - 입계자인 경우: 부 : 00(생부 00) 로 띄어쓰기 표시
    - 출계자인 경우: 부 : 00(양부 00) 로 띄어쓰기 표시
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

    ko_bio = str(row.get("한글 약력 정리", ""))
    first_ko = ko_bio.split("\n")[0] if ko_bio else ""

    details = [
        ("세수", gen_str),
        ("부", father_display),
        ("조부", gfather_ko),
    ]

    m = re.search(r'자는\s*([^\.,\n]+)', first_ko)
    if m: details.append(("자", m.group(1).strip()))
    m = re.search(r'초자는?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("초자", m.group(1).strip()))
    m = re.search(r'호는\s*([^\.,\n]+)', first_ko)
    if m: details.append(("호", m.group(1).strip()))
    m = re.search(r'초휘는?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("초휘", m.group(1).strip()))
    m = re.search(r'일휘는?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("일휘", m.group(1).strip()))
    m = re.search(r'통명은?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("통명", m.group(1).strip()))
    m = re.search(r'초명은?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("초명", m.group(1).strip()))

    for line in ko_bio.split("\n"):
        if line.startswith("과거·관직 :"):
            details.append(("과거·관직", line.replace("과거·관직 :", "").strip()))
        elif line.startswith("학문·인품·업적 :"):
            details.append(("학력·인품·업적", line.replace("학문·인품·업적 :", "").strip()))
        elif line.startswith("묘 :"):
            details.append(("묘", line.replace("묘 :", "").strip()))
        elif line.startswith("배 :") or line.startswith("배위 :"):
            details.append(("배위", re.sub(r'^(배|배위)\s*:\s*', '', line).strip()))

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
    ⭐ 인물 검색 결과 카드 통합 출력:
    1. 상단 타이틀: 출계자인 경우 '함자 00  0자 (출계)'로 표시
    2. '신상 및 계보 정보': 입계자인 경우 '부 : 00(생부 00)', 출계자인 경우 '부 : 00(양부 00)' 표시
    3. 상세·한글·한문 약력 창을 모두 없애고 '한문 약력 원문'만 표시
    """
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
    target_gid = target_row['고유ID']
    target_name = target_row['인명']
    
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
        sibs = all
