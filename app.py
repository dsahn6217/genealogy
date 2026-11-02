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
    """검색 완료 시 화면이 즉시 검색 결과창 위치로 자동 스크롤되도록 처리"""
    components.html(
        """
        <script>
            setTimeout(function() {
                var target = window.parent.document.getElementById('search-result-anchor');
                if (target) {
                    target.scrollIntoView({ behavior: 'smooth', block: 'start' });
                }
            }, 120);
        </script>
        """,
        height=0
    )

# =============================================================================
# 1. 페이지 설정 및 디자인 스타일 CSS (여백 초밀착, 한 줄 라디오, 입력창 굵은 테두리)
# =============================================================================
st.set_page_config(
    page_title="광주안씨 완귀공파 종중 족보 검색 시스템",
    page_icon=ICON_FILE if os.path.exists(ICON_FILE) else "📜",
    layout="wide",
    initial_sidebar_state="collapsed"
)

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&family=Nanum+Myeongjo:wght@400;700&display=swap');

html, body, [class*="css"], .stMarkdown, p, div {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', 'Malgun Gothic', sans-serif;
}

/* 최상단 밀착 배치를 위한 상단 여백 제거 */
header[data-testid="stHeader"] {
    display: none !important;
}
.main .block-container, 
[data-testid="stAppViewBlockContainer"], 
.block-container {
    max-width: 100% !important;
    width: 100% !important;
    padding-top: 0rem !important;
    padding-bottom: 1rem !important;
    padding-left: 0.5rem !important;
    padding-right: 0.5rem !important;
}

/* 창 별로 간격이 벌어져 있는 것을 최대한 좁힘 */
div[data-testid="stVerticalBlock"] > div {
    gap: 0.1rem !important;
    margin-bottom: 0px !important;
}
.element-container {
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

/* 제일 위 족보검색 시스템 네모 상자 (상단 여백없이 밀착, 하단 간격 최소화) */
.title-container {
    width: 100% !important;
    box-sizing: border-box !important;
    margin-top: 0px !important;
    margin-bottom: 2px !important;
    padding: 6px 10px 6px 10px;
    border-bottom: 3px solid #8c2d19;
    background: #ffffff !important;
    opacity: 1 !important;
    border-radius: 6px;
    box-shadow: 0 3px 8px rgba(0,0,0,0.18);
}

/* 좌측 최상단에 아이콘 + '광주안씨 완귀공파 종중' */
.title-top-left {
    display: flex;
    align-items: center;
    justify-content: flex-start;
    gap: 6px;
    margin: 0px !important;
    padding: 0px !important;
    line-height: 1.05;
}
.title-clan-name {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
    font-size: 13.5pt;
    font-weight: 800;
    color: #1e3c72;
    letter-spacing: -0.3px;
}

/* 가로 중앙에 '족보 검색 시스템' */
.title-main-center {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
    font-size: 16.5pt;
    font-weight: 800;
    color: #1a202c;
    text-align: center;
    width: 100%;
    margin-top: 1px;
    margin-bottom: 1px;
    line-height: 1.1;
    letter-spacing: -0.5px;
}

/* 하단: '- 완귀공(휘 嶒, 24세) ~ 44세'<br>'총 2,887인(고위 기준) -' */
.title-sub-bottom {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
    font-size: 10.5pt;
    font-weight: 600;
    color: #334155;
    text-align: center;
    width: 100%;
    margin-top: 1px;
    line-height: 1.2;
    letter-spacing: -0.2px;
}

/* ⭐ '족보 검색' 단어 삭제 후 안내문을 상단 창 바로 아래에 밀착 표시 */
.compact-search-header {
    margin-top: 2px !important;
    margin-bottom: 1px !important;
    padding: 0px 2px !important;
    line-height: 1.2 !important;
}
.compact-search-subtitle {
    font-size: 12pt !important;
    font-weight: 800 !important;
    color: #fef08a !important;
    margin-top: 0px !important;
    margin-bottom: 1px !important;
    letter-spacing: -0.4px !important;
    text-shadow: 0 1px 3px rgba(0,0,0,0.65);
}

/* ⭐ 검색 방식(인명 검색, 계보 검색, 단어 검색) 모바일에서도 무조건 한 줄로 표시 */
div[role="radiogroup"] {
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: nowrap !important;
    align-items: center !important;
    justify-content: flex-start !important;
    gap: 10px !important;
    width: 100% !important;
    margin-top: 0px !important;
    margin-bottom: 2px !important;
    padding: 2px 2px !important;
    overflow-x: auto !important;
}
div[role="radiogroup"] > label,
div[role="radiogroup"] label[data-baseweb="radio"] {
    white-space: nowrap !important;
    flex-shrink: 0 !important;
    margin-right: 4px !important;
    padding: 2px 2px !important;
}
div[role="radiogroup"] label p,
div[role="radiogroup"] label span {
    font-size: 12.5pt !important;
    font-weight: 800 !important;
    color: #ffffff !important;
    white-space: nowrap !important;
    letter-spacing: -0.3px !important;
    text-shadow: 0 1px 3px rgba(0,0,0,0.65);
}

/* ⭐ 검색어 입력창 테두리를 굵게 표시 및 예시 글씨 색상 회색 지정 */
div[data-testid="stTextInput"] {
    margin-top: 0px !important;
    margin-bottom: 2px !important;
}
div[data-testid="stTextInputRootElement"],
div[data-testid="stTextInput"] > div > div,
div[data-testid="stTextInput"] div[data-baseweb="base-input"],
div[data-testid="stTextInput"] div[data-baseweb="input"] {
    border: 3.5px solid #f59e0b !important;
    border-radius: 8px !important;
    background-color: #ffffff !important;
    box-shadow: 0 0 0 2px #0f172a, 0 3px 8px rgba(0,0,0,0.3) !important;
    box-sizing: border-box !important;
}
div[data-testid="stTextInput"] input {
    color: #4b5563 !important;
    font-weight: 700 !important;
    font-size: 12pt !important;
    padding: 7px 10px !important;
    background-color: #ffffff !important;
}
div[data-testid="stTextInput"] input::placeholder {
    color: #888888 !important;
    font-weight: 500 !important;
    opacity: 1 !important;
}
div[data-testid="stTextInput"] input::-webkit-input-placeholder {
    color: #888888 !important;
    font-weight: 500 !important;
    opacity: 1 !important;
}

/* 검색 방식별 동적 안내문 상자 (간격 최소화) */
.search-guide-right-text {
    font-size: 10pt;
    color: #0f172a;
    background-color: #fef9c3;
    padding: 5px 10px;
    border-radius: 6px;
    border: 2px solid #eab308;
    line-height: 1.25;
    font-weight: 600;
    margin-top: 0px !important;
    margin-bottom: 2px !important;
}

/* 검색 결과 인물 카드 (간격 최소화) */
.result-card {
    background-color: #ffffff;
    border: 2px solid #1e3c72;
    border-top: 4px solid #1e3c72;
    border-radius: 8px;
    padding: 10px 12px;
    margin-top: 4px !important;
    margin-bottom: 6px !important;
    box-shadow: 0 4px 12px rgba(0,0,0,0.2);
}
.result-card-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 2px solid #e2e8f0;
    padding-bottom: 4px;
    margin-bottom: 6px;
}
.result-name-big {
    font-size: 16pt;
    font-weight: 800;
    color: #1e3c72;
    display: flex;
    align-items: center;
    gap: 8px;
}
.result-badge {
    font-size: 10pt;
    font-weight: 700;
    background: #eff6ff;
    color: #1d4ed8;
    padding: 2px 8px;
    border-radius: 10px;
    border: 1px solid #bfdbfe;
}

/* 인물 기본 신상 및 계보 정보 (한글 전용 문단 형식, 간격 최소화) */
.meta-para-container {
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 8px 10px;
    margin-bottom: 8px;
    line-height: 1.38;
    font-size: 10.5pt;
    color: #0f172a;
}
.meta-para-item {
    margin-bottom: 2px;
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

/* ⭐ 상세 한글·한문 약력 박스 가로(상하 풀와이드) 배치: 한글 약력 먼저, 그 뒤에 정리된 한문 약력 배치 */
.bio-horizontal-stack {
    display: flex;
    flex-direction: column;
    gap: 6px;
    width: 100%;
}
.bio-detail-container {
    width: 100% !important;
    box-sizing: border-box !important;
    background-color: #fcfcfd;
    border: 1px solid #cbd5e1;
    border-left: 4px solid #1e3c72;
    border-radius: 6px;
    padding: 8px 12px;
    margin-bottom: 0px !important;
    line-height: 1.45;
    font-size: 10.5pt;
    color: #1e293b;
}
.bio-section-title {
    font-size: 11pt;
    font-weight: 800;
    color: #1e3c72;
    margin-bottom: 4px;
    padding-bottom: 2px;
    border-bottom: 1px solid #cbd5e1;
}
.bio-raw-subbox {
    margin-top: 6px;
    padding-top: 5px;
    border-top: 1px dashed #94a3b8;
    font-size: 10pt;
    color: #334155;
    line-height: 1.4;
    background-color: #f8fafc;
    padding: 6px 8px;
    border-radius: 4px;
}
.bio-tag {
    font-weight: 700;
    color: #1e40af;
    background: #eff6ff;
    padding: 1px 5px;
    border-radius: 4px;
    margin-right: 4px;
    font-size: 9.5pt;
    display: inline-block;
}
.red-correction {
    color: #dc2626 !important;
    font-weight: 700;
    background-color: #fef2f2;
    padding: 0px 3px;
    border-radius: 3px;
    border-bottom: 1px dashed #f87171;
}

/* 단어 검색어 빨간색 강조 및 다수 선조 안내 배너 */
.search-word-red {
    color: #ff0000 !important;
    font-weight: 800 !important;
    background-color: #fee2e2;
    padding: 0px 4px;
    border-radius: 3px;
    border: 1px solid #f87171;
}
.multi-ancestor-banner {
    background-color: #fff1f2;
    border: 2px solid #e11d48;
    color: #9f1239;
    padding: 6px 10px;
    border-radius: 6px;
    font-size: 10.5pt;
    font-weight: 700;
    margin-top: 3px;
    margin-bottom: 4px;
    line-height: 1.3;
}

/* 계보도 도형 간 간격 최소화 및 도형 간 연결선 표시 */
.tree-diagram-wrapper {
    background: #ffffff;
    border: 2px solid #1e3c72;
    border-radius: 8px;
    padding: 8px 10px;
    margin-top: 4px;
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
}
.tree-connector-line {
    text-align: center;
    color: #1e3c72;
    font-weight: 900;
    font-size: 10pt;
    line-height: 0.95;
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
    gap: 3px;
    align-items: center;
}
.tree-node-box {
    background: #ffffff;
    border: 1.5px solid #64748b;
    border-radius: 4px;
    padding: 2px 7px;
    font-size: 9.5pt;
    color: #0f172a;
    line-height: 1.2;
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

/* 종중 주요 유적 및 문화재 안내 (항상 검색 결과창 아래에 위치) */
.heritage-sub-container {
    background-color: #ffffff;
    border: 2px solid #cbd5e1;
    border-radius: 8px;
    padding: 8px 12px;
    margin-top: 6px;
    margin-bottom: 6px;
}
.heritage-sub-header {
    font-size: 11pt;
    font-weight: 800;
    color: #1e3c72;
    margin-bottom: 6px;
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
    margin-bottom: 1px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.h-card-desc {
    font-size: 8.8pt;
    color: #475569;
    line-height: 1.25;
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
# 3. 보조 유틸리티 함수: 세수, 인명 분리, 한글 전용 신상 문단 생성, 계보도
# =============================================================================
def extract_generation_info(gid):
    m = re.search(r"-(\d+)$", str(gid).strip())
    if m:
        gen = int(m.group(1))
        wangwi_desc_gen = gen - 24
        desc_str = "완귀공 본인" if wangwi_desc_gen == 0 else f"완귀공의 {wangwi_desc_gen}대손"
        return gen, f"{gen}세손 ({desc_str})"
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
    ⭐ 인물 기본 신상 및 계보 정보에서 한글로만 표시하고 한문은 기재하지 않음
    """
    gid = row.get("고유ID", "")
    gen_num, gen_str = extract_generation_info(gid)
    
    father_ko, _ = split_name(row.get("부친", ""))
    gfather_ko, _ = split_name(row.get("조부", ""))
    
    ko_bio = str(row.get("한글 약력 정리", ""))
    first_ko = ko_bio.split("\n")[0] if ko_bio else ""
    
    details = [
        ("세수", gen_str),
        ("부", father_ko if father_ko else re.sub(r"[一-龥]+", "", str(row.get("부친", ""))).strip()),
        ("조부", gfather_ko if gfather_ko else re.sub(r"[一-龥]+", "", str(row.get("조부", ""))).strip()),
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
        elif line.startswith("배 :") or line.startswith("배위 :"):
            details.append(("배위", re.sub(r'^(배|배위)\s*:\s*', '', line).strip()))
            
    return details

def format_hangul_only_paragraph_html(row, search_term=None):
    """
    ⭐ 인물 기본 신상 및 계보 정보를 한글로만 표시 (내용이 없는 항목은 생략)
    """
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

def format_structured_bio(bio_text, is_hanmun=False, search_term=None):
    if not bio_text:
        return '<div style="color:#64748b; font-style:italic;">수록된 정리 약력이 없습니다.</div>'
    lines = [ln.strip() for ln in str(bio_text).split("\n") if ln.strip()]
    html_items = []
    for ln in lines:
        hl = highlight_corrections_html(ln, search_term=search_term)
        if ln.startswith("묘 :") or ln.startswith("墓 :"):
            lbl = "墓" if is_hanmun else "묘"
            html_items.append(f'<div><span class="bio-tag">{lbl}</span> {hl[3:].strip()}</div>')
        elif ln.startswith("과거·관직 :") or ln.startswith("科·官 :"):
            lbl = "科·官" if is_hanmun else "과거·관직"
            cut = 5 if is_hanmun else 6
            html_items.append(f'<div><span class="bio-tag">{lbl}</span> {hl[cut:].strip()}</div>')
        elif ln.startswith("학문·인품·업적 :") or ln.startswith("學·行·業 :"):
            lbl = "學·行·業" if is_hanmun else "학문·인품·업적"
            cut = 6 if is_hanmun else 9
            html_items.append(f'<div><span class="bio-tag">{lbl}</span> {hl[cut:].strip()}</div>')
        elif ln.startswith("배 :") or ln.startswith("配 :"):
            lbl = "配" if is_hanmun else "배"
            html_items.append(f'<div><span class="bio-tag">{lbl}</span> {hl[3:].strip()}</div>')
        else:
            lbl = "身常" if is_hanmun else "신상"
            html_items.append(f'<div><span class="bio-tag">{lbl}</span> {hl}</div>')
    return "".join(html_items)

def render_person_result_card(target_row, search_term=None):
    """
    ⭐ 인물 검색 결과 카드 통합 출력:
    1. 인물 기본 신상 및 계보 정보: 한글로만 표시 (한문 미기재)
    2. 상세 한글·한문 약력란: 박스를 가로형(위아래 순차 배치)으로 하여
       한글 약력(정리+원문)을 먼저 배치하고, 그 뒤에 한문 약력 창은 '정리된 한문 약력'만 기재
    """
    h_name, _ = split_name(target_row['인명'])
    meta_paragraphs_html = format_hangul_only_paragraph_html(target_row, search_term=search_term)
    
    ko_raw = highlight_corrections_html(target_row.get('한글 약력 원문', ''), search_term=search_term)
    
    render_html(f"""
    <div id="search-result-anchor"></div>
    <div class="result-card">
        <div class="result-card-header">
            <div class="result-name-big">
                <span>{h_name if h_name else target_row['인명']}</span>
                <span class="result-badge">{target_row['관계']}</span>
            </div>
            <div style="font-size:10pt; color:#6b7280; font-family:monospace;">고유ID: {target_row['고유ID']}</div>
        </div>
        
        <div style="font-weight:800; color:#1e3c72; margin-bottom:3px; font-size:11pt;">
            📌 인물 기본 신상 및 계보 정보
        </div>
        <div class="meta-para-container">
            {meta_paragraphs_html}
        </div>
        
        <div style="font-weight:800; color:#1e3c72; margin-top:6px; margin-bottom:4px; font-size:11pt;">
            📜 상세 한글·한문 약력
        </div>
        
        <div class="bio-horizontal-stack">
            <div class="bio-detail-container">
                <div class="bio-section-title">1. 한글 약력 (정리 및 한글 원문)</div>
                {format_structured_bio(target_row.get('한글 약력 정리', ''), is_hanmun=False, search_term=search_term)}
                <div class="bio-raw-subbox">
                    <b style="color:#1e3c72;">[한글 약력 원문]</b><br>{ko_raw if ko_raw else '수록된 원문이 없습니다.'}
                </div>
            </div>
            <div class="bio-detail-container">
                <div class="bio-section-title">2. 한문 약력 (정리된 한문 약력)</div>
                {format_structured_bio(target_row.get('한문 약력 정리', ''), is_hanmun=True, search_term=search_term)}
            </div>
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
        sibs = all_df[all_df['부친'] == ff_name] if ff_name else pd.DataFrame([f_row])
        ancestors.append({"level": up, "row": f_row, "siblings": sibs})
        curr = f_row

    diagram_parts = [f'<div id="search-result-anchor"></div><div class="tree-diagram-wrapper"><div style="font-weight:800; color:#1e3c72; font-size:11.5pt; margin-bottom:4px;">🌳 [{target_name}] 기준 혈통 계보도</div>']
    
    for idx_a, anc in enumerate(reversed(ancestors)):
        lvl = anc['level']
        r = anc['row']
        g_num, g_str = extract_generation_info(r['고유ID'])
        sibs = anc['siblings']
        
        if idx_a > 0:
            diagram_parts.append('<div class="tree-connector-line">│<br>▼</div>')
            
        if lvl >= 3:
            rel_str = f" ({len(sibs)}자 중 {r['관계']})" if len(sibs) > 1 else f" ({r['관계']})"
            lbl_badge = "증조부" if lvl == 3 else f"{lvl}대조"
            diagram_parts.append(f"""
            <div class="tree-level-row">
                <span class="tree-gen-badge">▲ {lbl_badge}({g_num}세)</span>
                <div class="tree-nodes-container">
                    <span class="tree-node-box tree-node-target"><b>{r['인명']}</b>{rel_str}</span>
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
                if s_i > 0:
                    node_htmls.append('<span class="tree-h-line">──</span>')
                node_htmls.append(f'<span class="{cls}"><b>{s_name}</b>({s_rel})</span>')
            nodes_joined = ''.join(node_htmls)
            diagram_parts.append(f"""
            <div class="tree-level-row">
                <span class="tree-gen-badge">▲ {lbl_badge}({g_num}세)</span>
                <div class="tree-nodes-container">{nodes_joined}</div>
            </div>
            """)

    if len(ancestors) > 0:
        diagram_parts.append('<div class="tree-connector-line">│<br>▼</div>')

    t_f = str(target_row.get('부친', '')).strip()
    t_sibs = all_df[all_df['부친'] == t_f] if t_f else pd.DataFrame([target_row])
    tg_num, tg_str = extract_generation_info(target_row['고유ID'])
    
    t_nodes = []
    for s_i, (_, s) in enumerate(t_sibs.iterrows()):
        is_t = (s['고유ID'] == target_row['고유ID'])
        cls = "tree-node-box tree-node-target" if is_t else "tree-node-box"
        marker = " 🎯[기준]" if is_t else ""
        s_name = s['인명']
        s_rel = s['관계']
        if s_i > 0:
            t_nodes.append('<span class="tree-h-line">──</span>')
        t_nodes.append(f'<span class="{cls}"><b>{s_name}</b>{marker}({s_rel})</span>')
    t_joined = ''.join(t_nodes)
    diagram_parts.append(f"""
    <div class="tree-level-row" style="background:#eff6ff; border:2px solid #2563eb;">
        <span class="tree-gen-badge" style="background:#1d4ed8; color:white;">● 본인({tg_num}세)</span>
        <div class="tree-nodes-container">{t_joined}</div>
    </div>
    """)

    sons = all_df[all_df['부친'] == target_name]
    grandsons_rows = []
    if len(sons) > 0:
        diagram_parts.append('<div class="tree-connector-line">│<br>▼</div>')
        s_nodes = []
        for s_i, (_, s) in enumerate(sons.iterrows()):
            s_name = s['인명']
            s_rel = s['관계']
            if s_i > 0:
                s_nodes.append('<span class="tree-h-line">──</span>')
            s_nodes.append(f'<span class="tree-node-box"><b>{s_name}</b>({s_rel})</span>')
        s_joined = ''.join(s_nodes)
        sg_num = tg_num + 1 if tg_num else 25
        diagram_parts.append(f"""
        <div class="tree-level-row">
            <span class="tree-gen-badge">▼ 자(子)({sg_num}세)</span>
            <div class="tree-nodes-container">{s_joined}</div>
        </div>
        """)
        
        if max_down >= 2:
            grandsons = []
            for _, s in sons.iterrows():
                gs = all_df[all_df['부친'] == s['인명']]
                for _, g in gs.iterrows():
                    grandsons_rows.append(g)
                    g_name = g['인명']
                    s_name = s['인명']
                    g_rel = g['관계']
                    if len(grandsons) > 0:
                        grandsons.append('<span class="tree-h-line">──</span>')
                    grandsons.append(f'<span class="tree-node-box"><b>{g_name}</b>({s_name}의 {g_rel})</span>')
            if grandsons:
                diagram_parts.append('<div class="tree-connector-line">│<br>▼</div>')
                g_joined = ''.join(grandsons)
                diagram_parts.append(f"""
                <div class="tree-level-row">
                    <span class="tree-gen-badge">▼ 손(孫)({sg_num+1}세)</span>
                    <div class="tree-nodes-container">{g_joined}</div>
                </div>
                """)

    diagram_parts.append('</div>')
    render_html("".join(diagram_parts))

    # 계보도 도형 클릭 시 즉시 새 계보 생성
    st.markdown("##### 👆 계보도 인물 클릭 (클릭 시 검색창 입력 없이 즉시 새 계보 생성)")
    clickable_Pool = []
    seen_gids = set()
    for anc in reversed(ancestors):
        for _, s in anc['siblings'].iterrows():
            if s['고유ID'] not in seen_gids:
                clickable_Pool.append(s)
                seen_gids.add(s['고유ID'])
    for _, s in t_sibs.iterrows():
        if s['고유ID'] not in seen_gids:
            clickable_Pool.append(s)
            seen_gids.add(s['고유ID'])
    for _, s in sons.iterrows():
        if s['고유ID'] not in seen_gids:
            clickable_Pool.append(s)
            seen_gids.add(s['고유ID'])
    for g in grandsons_rows:
        if g['고유ID'] not in seen_gids:
            clickable_Pool.append(g)
            seen_gids.add(g['고유ID'])

    if clickable_Pool:
        cols_per_row = 6
        for i in range(0, len(clickable_Pool), cols_per_row):
            chunk = clickable_Pool[i:i+cols_per_row]
            b_cols = st.columns(len(chunk))
            for c_i, p_row in enumerate(chunk):
                p_gid = p_row['고유ID']
                is_curr = (p_gid == target_gid)
                btn_lbl = f"{'🎯 ' if is_curr else ''}{p_row['인명']}({p_row['관계']})"
                if b_cols[c_i].button(btn_lbl, key=f"{key_prefix}_node_{p_gid}", use_container_width=True):
                    st.session_state['active_tree_gid'] = p_gid
                    st.rerun()

    render_person_result_card(target_row)

# =============================================================================
# 4. 최상단 헤더 및 직결된 족보 검색 인터페이스 ('족보 검색' 단어 삭제)
# =============================================================================
icon_img_tag = f'<img src="{icon_b64}" style="width:26px; height:26px; border-radius:50%; vertical-align:middle;">' if icon_b64 else '📜'

# 저장소 배경 이미지가 있을 경우 주입 (투명율 0)
bg_style_tag = f"<style>.stApp {{ background-image: url('{bg_b64}') !important; background-size: cover !important; background-attachment: fixed !important; opacity: 1 !important; }}</style>" if bg_b64 else ""
if bg_style_tag:
    render_html(bg_style_tag)

# ⭐ 1. 최상단 족보 검색 시스템 상자 + 바로 아래 안내문 표시 ('족보 검색' 단어 삭제)
render_html(f"""
<div class="title-container">
    <div class="title-top-left">
        {icon_img_tag}
        <span class="title-clan-name">광주안씨 완귀공파 종중</span>
    </div>
    <div class="title-main-center">
        족보 검색 시스템
    </div>
    <div class="title-sub-bottom">
        - 완귀공(휘 嶒, 24세) ~ 44세<br>총 2,887인(고위 기준) -
    </div>
</div>
<div class="compact-search-header">
    <div class="compact-search-subtitle">아래 방식 중 하나를 선택하고 검색어를 한글로 입력하세요</div>
</div>
""")

if df is None:
    st.error(f"데이터 로드 오류: {load_msg}")
    st.info("파일 이름이 `genealogy.xlsx`인지 확인해 주시기 바랍니다.")
    st.stop()

# ⭐ 2. 검색 방식 선택 (인명 검색, 계보 검색, 단어 검색 - 한 줄로 표시)
search_type = st.radio(
    "검색 종류 선택",
    ["인명 검색", "계보 검색", "단어 검색"],
    horizontal=True,
    label_visibility="collapsed"
)

# ⭐ 3. 굵은 테두리 검색어 입력창 (예시 중 한문 삭제 및 회색 글씨) 및 동적 안내문
c_input, c_guide = st.columns([1.1, 2.4])

with c_input:
    if search_type == "인명 검색":
        query_val = st.text_input("검색어 입력", placeholder="예: 경시, 안경시, 증, 완귀", key="name_query", label_visibility="collapsed")
    elif search_type == "계보 검색":
        query_val = st.text_input("검색어 입력", placeholder="예: 증, 명한, 안경시", key="lineage_query", label_visibility="collapsed")
    else:
        query_val = st.text_input("검색어 입력", placeholder="예: 시강원, 헌릉참봉, 1699", key="word_query", label_visibility="collapsed")

with c_guide:
    if search_type == "인명 검색":
        render_html("""
        <div class="search-guide-right-text">
            💡 <b>[인명 검색 안내]</b> 인물의 함자(성씨 안을 넣어도 무방), 자(字), 초자(初字), 호(號), 초휘(初諱), 일휘(一諱) 통명(通名), 초명(初名) 등을 한글로 입력
        </div>
        """)
    elif search_type == "계보 검색":
        render_html("""
        <div class="search-guide-right-text">
            🌳 <b>[계보 검색 안내]</b> 계보를 조회할 인물의 함자(성씨 안 포함 가능), 자(字), 호(號), 초휘(初諱) 등을 한글로 입력 (생성된 계보도 도형 클릭 시 즉시 새 계보 생성)
        </div>
        """)
    else:
        render_html("""
        <div class="search-guide-right-text">
            🔍 <b>[단어 검색 안내]</b> 약력에 포함된 관직명, 과거 급제, 지명, 묘소 좌향, 시호, 연대 등의 단어를 한글로 입력 (검색어는 <span style="color:#dc2626;">빨간색</span>으로 표시됨)
        </div>
        """)

# =============================================================================
# ⭐ 4. 검색 결과 창 (항상 '종중 주요 유적 및 문화재 안내' 창 위에 뜨도록 고정 + 즉시 스크롤)
# =============================================================================
if search_type == "인명 검색" and query_val.strip():
    q = query_val.strip()
    q_alt = q[1:] if q.startswith("안") and len(q) > 1 else q
    
    mask = (
        df['인명'].astype(str).str.contains(q, case=False, na=False) |
        df['인명'].astype(str).str.contains(q_alt, case=False, na=False) |
        df['한글 약력 정리'].apply(lambda x: q in str(x).split('\n')[0] or q_alt in str(x).split('\n')[0]) |
        df['한문 약력 정리'].apply(lambda x: q in str(x).split('\n')[0] or q_alt in str(x).split('\n')[0])
    )
    matched_df = df[mask]
    
    if len(matched_df) == 0:
        st.warning(f"입력하신 검색어 '{q}'에 해당하는 인물을 찾을 수 없습니다.")
    else:
        if len(matched_df) > 1:
            render_html(f"""
            <div class="multi-ancestor-banner">
                📢 동명이인 또는 복수 후보가 총 {len(matched_df)}위 검색되었습니다. 아래에서 조회할 인물을 선택하세요.
            </div>
            """)
            candidate_labels = []
            for idx, row in matched_df.iterrows():
                gen_num, gen_str = extract_generation_info(row['고유ID'])
                lbl = f"[{gen_str}] {row['인명']} (부: {row['부친']}, 조부: {row['조부']}, 관계: {row['관계']}) - 고유ID: {row['고유ID']}"
                candidate_labels.append(lbl)
            
            selected_label = st.selectbox("선조 선택", candidate_labels, label_visibility="collapsed")
            selected_gid = selected_label.split("고유ID: ")[-1].strip()
            target_row = matched_df[matched_df['고유ID'] == selected_gid].iloc[0]
        else:
            target_row = matched_df.iloc[0]
            
        render_person_result_card(target_row)
        scroll_to_result()

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
        q_alt = q[1:] if q.startswith("안") and len(q) > 1 else q
        mask = (
            df['인명'].astype(str).str.contains(q, case=False, na=False) |
            df['인명'].astype(str).str.contains(q_alt, case=False, na=False) |
            df['한글 약력 정리'].apply(lambda x: q in str(x).split('\n')[0] or q_alt in str(x).split('\n')[0])
        )
        l_matches = df[mask]
        if len(l_matches) == 0:
            st.warning("계보를 작성할 선조 인명을 찾을 수 없습니다.")
        elif len(l_matches) > 1:
            cand_opts = [f"{r['인명']} (세수: {extract_generation_info(r['고유ID'])[1]}, 부: {r['부친']}) - {r['고유ID']}" for _, r in l_matches.iterrows()]
            sel_cand = st.selectbox("계보 기준 인물 선택", cand_opts, label_visibility="collapsed")
            target_gid = sel_cand.split(" - ")[-1].strip()
            tree_target = l_matches[l_matches['고유ID'] == target_gid].iloc[0]
        else:
            tree_target = l_matches.iloc[0]

    if tree_target is not None:
        c_cond1, c_cond2, c_cond3 = st.columns([1.2, 1.2, 1])
        with c_cond1:
            up_gens = st.slider("상향 세대수 (선대 조상)", min_value=1, max_value=5, value=3)
        with c_cond2:
            down_gens = st.slider("하향 세대수 (후손 자손)", min_value=1, max_value=4, value=2)
        with c_cond3:
            if st.session_state.get('active_tree_gid'):
                if st.button("🔄 초기 검색 인물로 복귀", use_container_width=True):
                    st.session_state['active_tree_gid'] = None
                    st.rerun()
        render_tree_hierarchy(tree_target, df, max_up=up_gens, max_down=down_gens, key_prefix="main_tree")
        scroll_to_result()

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
        st.warning(f"입력하신 단어 '{term}'을(를) 포함하고 있는 약력 기록이 없습니다.")
    else:
        render_html(f"""
        <div class="multi-ancestor-banner">
            📢 검색어 <span class="search-word-red">'{term}'</span>을(를) 포함하는 약력을 가진 선조가 총 <b>{len(word_matches):,}위(다수)</b> 검색되었습니다. 아래 목록에서 선조를 선택하세요.
        </div>
        """)
        candidate_list = []
        for _, r in word_matches.iterrows():
            gen_num, gen_str = extract_generation_info(r['고유ID'])
            candidate_list.append(f"[{gen_str}] {r['인명']} (부: {r['부친']}, 관계: {r['관계']}) - 고유ID: {r['고유ID']}")
            
        chosen = st.selectbox("조회할 선조를 선택하세요", candidate_list, label_visibility="collapsed")
        chosen_gid = chosen.split("고유ID: ")[-1].strip()
        chosen_row = word_matches[word_matches['고유ID'] == chosen_gid].iloc[0]
        
        render_person_result_card(chosen_row, search_term=term)
        scroll_to_result()

# =============================================================================
# ⭐ 5. 종중 주요 유적 및 문화재 안내 (항상 검색 결과 창 아래에 위치)
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
# 6. 하단 푸터 (Footer)
# =============================================================================
render_html("""
<div style="text-align:center; padding:8px 0 4px 0; color:#cbd5e1; font-size:9pt; border-top:1px solid #334155; margin-top:8px;">
    광주안씨 완귀공파 종중 회의소 | 데이터 출처: 보정 족보 1~2권 통합본 (총 2,887위) | 연동 파일: genealogy.xlsx
</div>
""")
