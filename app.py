import streamlit as st
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

# =============================================================================
# 1. 페이지 설정 및 디자인 스타일 CSS (제4장~제7장 요령 반영)
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

/* 제5장 1항: 최상단 밀착 배치를 위한 상단 여백 제거 */
header[data-testid="stHeader"] {
    display: none !important;
}
.main .block-container, 
[data-testid="stAppViewBlockContainer"], 
.block-container {
    max-width: 100% !important;
    width: 100% !important;
    padding-top: 0rem !important;  /* 최상단 여백 없이 완전 밀착 */
    padding-bottom: 1.5rem !important;
    padding-left: 0.8rem !important;
    padding-right: 0.8rem !important;
}

/* 제5장 2항 & 제4장 4항 ④: 짙은 청색 배경 대비 및 투명도 제로(100% 완전 불투명) */
.stApp {
    background-color: #162b52;
    background-size: cover !important;
    background-position: center top !important;
    background-attachment: fixed !important;
    opacity: 1 !important;         /* 투명율 0 (완전 불투명) */
}

/* 검색창의 모든 기본 글자는 배경이 짙은 청색이므로 밝은 흰색/금백색 계열로 가독성 확보 */
[data-testid="stMarkdownContainer"] h3,
[data-testid="stMarkdownContainer"] h4,
[data-testid="stMarkdownContainer"] h5 {
    color: #ffffff !important;
    margin-top: 4px !important;
    margin-bottom: 4px !important;
    line-height: 1.25 !important;
}
div[role="radiogroup"] label p,
div[role="radiogroup"] label span,
.stCaption, [data-testid="stCaptionContainer"] {
    color: #f8fafc !important;
    font-weight: 600 !important;
}

/* 제4장 4항 ④: 도형과 도형 사이, 글자 행과 행 사이의 간격을 최대한 좁힘 */
div[data-testid="stVerticalBlock"] > div {
    gap: 0.25rem !important;
}

/* ⭐ 제4장 4항 ⑤: 입력하는 검색창은 가장 중요한 부분이므로 짙은 테두리를 적용 */
div[data-testid="stTextInput"] div[data-baseweb="input"] {
    border: 3px solid #0f172a !important;
    border-radius: 6px !important;
    background-color: #ffffff !important;
    box-shadow: 0 0 0 2px #f59e0b, 0 4px 10px rgba(0,0,0,0.25) !important;
}
div[data-testid="stTextInput"] input {
    color: #0f172a !important;
    font-weight: 700 !important;
    font-size: 11.5pt !important;
    padding: 6px 10px !important;
}

/* 네비게이션 탭 글자 가독성 */
button[data-baseweb="tab"] p {
    color: #e2e8f0 !important;
    font-weight: 700 !important;
    font-size: 11pt !important;
}
button[data-baseweb="tab"][aria-selected="true"] p {
    color: #fbbf24 !important;
}

/* 제5장 1항: 제일 위 족보검색 시스템 네모 상자 (상단 여백없이 밀착, 불투명 백색) */
.title-container {
    width: 100% !important;
    box-sizing: border-box !important;
    margin-top: 0px !important;    /* 상단 여백없이 밀착 */
    margin-bottom: 6px !important;
    padding: 6px 12px 8px 12px;
    border-bottom: 3px solid #8c2d19;
    background: #ffffff !important;
    opacity: 1 !important;
    border-radius: 6px;
    box-shadow: 0 3px 10px rgba(0,0,0,0.18);
}

/* 여백없이 좌측 최상단에 아이콘 + '광주안씨 완귀공파 종중' */
.title-top-left {
    display: flex;
    align-items: center;
    justify-content: flex-start;
    gap: 6px;
    margin: 0px !important;
    padding: 0px !important;
    line-height: 1.1;
}
.title-clan-name {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
    font-size: 14pt;
    font-weight: 800;
    color: #1e3c72;
    letter-spacing: -0.3px;
}

/* 그 오른쪽 한행 아래 단에 '족보 검색 시스템'을 가로 중앙에 배치 */
.title-main-center {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
    font-size: 17pt;
    font-weight: 800;
    color: #1a202c;
    text-align: center;
    width: 100%;
    margin-top: 1px;
    margin-bottom: 2px;
    line-height: 1.15;
    letter-spacing: -0.5px;
}

/* 그 아래에 '- 완귀공(휘 嶒, 24세) ~ 44세'<br>'총 2,887인(고위 기준) -' 표시 */
.title-sub-bottom {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
    font-size: 11pt;
    font-weight: 600;
    color: #334155;
    text-align: center;
    width: 100%;
    margin-top: 2px;
    line-height: 1.25;
    letter-spacing: -0.2px;
}

/* 제4장 4항 ⑥: 검색 방식별 동적 안내문 상자 (행간 및 여백 최소화) */
.search-guide-right-text {
    font-size: 10pt;
    color: #0f172a;
    background-color: #fef9c3;
    padding: 6px 10px;
    border-radius: 6px;
    border: 2px solid #eab308;
    line-height: 1.3;
    font-weight: 600;
}

/* 제4장 5항: 검색 결과 인물 카드 (문단 형식, 행간/도형 간격 최소화) */
.result-card {
    background-color: #ffffff;
    border: 2px solid #1e3c72;
    border-top: 4px solid #1e3c72;
    border-radius: 8px;
    padding: 14px 16px;
    margin-top: 8px;
    margin-bottom: 12px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
}
.result-card-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 2px solid #e2e8f0;
    padding-bottom: 6px;
    margin-bottom: 8px;
}
.result-name-big {
    font-size: 17pt;
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
    padding: 2px 10px;
    border-radius: 12px;
    border: 1px solid #bfdbfe;
}

/* ⭐ 제4장 5항 ②, ④: 표를 사용하지 않고 일반 문단 형식으로 작성 (행간 축소) */
.meta-para-container {
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 10px 12px;
    margin-bottom: 10px;
    line-height: 1.45;
    font-size: 10.5pt;
    color: #0f172a;
}
.meta-para-item {
    margin-bottom: 3px;
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

/* 5단락 상세 약력 박스 (제목에 kr, cn 사용금지) */
.bio-detail-container {
    background-color: #fcfcfd;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 10px 14px;
    margin-bottom: 10px;
    line-height: 1.5;
    font-size: 10.5pt;
    color: #1e293b;
}
.bio-section-title {
    font-size: 11.5pt;
    font-weight: 800;
    color: #1e3c72;
    margin-bottom: 6px;
    padding-bottom: 3px;
    border-bottom: 1px solid #cbd5e1;
}
.bio-tag {
    font-weight: 700;
    color: #1e40af;
    background: #eff6ff;
    padding: 1px 6px;
    border-radius: 4px;
    margin-right: 5px;
    font-size: 10pt;
    display: inline-block;
}
.red-correction {
    color: #dc2626 !important;
    font-weight: 700;
    background-color: #fef2f2;
    padding: 0px 4px;
    border-radius: 3px;
    border-bottom: 1px dashed #f87171;
}

/* ⭐ 제4장 5항 ⑦: 단어 검색어 빨간색 강조 및 다수 선조 안내 배너 */
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
    padding: 8px 12px;
    border-radius: 6px;
    font-size: 11pt;
    font-weight: 700;
    margin-top: 6px;
    margin-bottom: 8px;
    line-height: 1.35;
}

/* ⭐ 제4장 5항 ⑤, ⑥: 계보도 도형 간 간격 최소화 및 도형 간 연결선 표시 */
.tree-diagram-wrapper {
    background: #ffffff;
    border: 2px solid #1e3c72;
    border-radius: 8px;
    padding: 10px 12px;
    margin-top: 6px;
    margin-bottom: 10px;
}
.tree-level-row {
    display: flex;
    align-items: center;
    margin-bottom: 0px !important;
    padding: 5px 8px;
    background: #f8fafc;
    border-radius: 6px;
    border: 1px solid #cbd5e1;
}
.tree-connector-line {
    text-align: center;
    color: #1e3c72;
    font-weight: 900;
    font-size: 11pt;
    line-height: 1.0;
    margin: 0px !important;
    padding: 1px 0px;
}
.tree-gen-badge {
    min-width: 82px;
    font-weight: 700;
    color: #1e3c72;
    font-size: 10pt;
    padding: 3px 6px;
    background: #dbeafe;
    border-radius: 4px;
    text-align: center;
    margin-right: 8px;
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
    padding: 3px 9px;
    font-size: 10pt;
    color: #0f172a;
    line-height: 1.25;
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
    font-size: 9pt;
    margin: 0 1px;
}

/* 제6장: 검색창 하단 부 메뉴 (종중 유적 문화재 안내) */
.heritage-sub-container {
    background-color: #ffffff;
    border: 2px solid #cbd5e1;
    border-radius: 8px;
    padding: 10px 14px;
    margin-top: 12px;
    margin-bottom: 10px;
}
.heritage-sub-header {
    font-size: 11.5pt;
    font-weight: 800;
    color: #1e3c72;
    margin-bottom: 8px;
}
.heritage-sub-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
    gap: 8px;
}
.heritage-sub-card {
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 7px 10px;
    text-decoration: none !important;
    display: block;
}
.heritage-sub-card:hover {
    background: #eff6ff;
    border-color: #2563eb;
}
.h-card-name {
    font-size: 10.5pt;
    font-weight: 700;
    color: #1e40af;
    margin-bottom: 2px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.h-card-desc {
    font-size: 9pt;
    color: #475569;
    line-height: 1.3;
}

/* 아코디언(원문 보기) 스타일 */
[data-testid="stExpander"] {
    border: 1px solid #cbd5e1 !important;
    border-radius: 6px !important;
    margin-bottom: 8px !important;
    background: #ffffff !important;
}
[data-testid="stExpander"] summary p,
[data-testid="stExpander"] summary span {
    font-size: 11pt !important;
    font-weight: 700 !important;
    color: #1e3c72 !important;
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
# 3. 보조 유틸리티 함수: 세수, 인명 분리, 고증 빨간색 하이라이트, 문단 생성, 계보도
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
    # 제4장 5항 ⑦: 단어 검색어는 빨간색으로 강조 표시
    if search_term and search_term.strip():
        sterm = search_term.strip()
        formatted = re.sub(
            f"({re.escape(sterm)})",
            r'<span class="search-word-red">\1</span>',
            formatted,
            flags=re.IGNORECASE
        )
    return formatted

def parse_person_bilingual_details(row):
    gid = row.get("고유ID", "")
    gen_num, gen_str = extract_generation_info(gid)
    
    father_ko, father_cn = split_name(row.get("부친", ""))
    gfather_ko, gfather_cn = split_name(row.get("조부", ""))
    
    ko_bio = str(row.get("한글 약력 정리", ""))
    cn_bio = str(row.get("한문 약력 정리", ""))
    
    first_ko = ko_bio.split("\n")[0] if ko_bio else ""
    first_cn = cn_bio.split("\n")[0] if cn_bio else ""
    
    details = {
        "세수": {"label": "세수(世代)", "ko": gen_str, "cn": f"{gen_num}世 (玩龜公 {gen_num-24}代孫)" if gen_num else ""},
        "부": {"label": "부(父)", "ko": father_ko if father_ko else row.get("부친", ""), "cn": father_cn},
        "조부": {"label": "조부(祖父)", "ko": gfather_ko if gfather_ko else row.get("조부", ""), "cn": gfather_cn},
        "자": {"label": "자(字)", "ko": "", "cn": ""},
        "초자": {"label": "초자(初字)", "ko": "", "cn": ""},
        "호": {"label": "호(號)", "ko": "", "cn": ""},
        "초휘": {"label": "초휘(初諱)", "ko": "", "cn": ""},
        "일휘": {"label": "일휘(一諱)", "ko": "", "cn": ""},
        "통명": {"label": "통명(通名)", "ko": "", "cn": ""},
        "초명": {"label": "초명(初名)", "ko": "", "cn": ""},
        "과거_관직": {"label": "과거·관직(科·官)", "ko": "", "cn": ""},
        "학문_업적": {"label": "학력·인품·업적(學·行·業)", "ko": "", "cn": ""},
        "배위": {"label": "배위(配位)", "ko": "", "cn": ""}
    }
    
    m = re.search(r'자는\s*([^\.,\n]+)', first_ko)
    if m: details["자"]["ko"] = m.group(1).strip()
    m = re.search(r'초자는?\s*([^\.,\n]+)', first_ko)
    if m: details["초자"]["ko"] = m.group(1).strip()
    m = re.search(r'호는\s*([^\.,\n]+)', first_ko)
    if m: details["호"]["ko"] = m.group(1).strip()
    m = re.search(r'초휘는?\s*([^\.,\n]+)', first_ko)
    if m: details["초휘"]["ko"] = m.group(1).strip()
    m = re.search(r'일휘는?\s*([^\.,\n]+)', first_ko)
    if m: details["일휘"]["ko"] = m.group(1).strip()
    m = re.search(r'통명은?\s*([^\.,\n]+)', first_ko)
    if m: details["통명"]["ko"] = m.group(1).strip()
    m = re.search(r'초명은?\s*([^\.,\n]+)', first_ko)
    if m: details["초명"]["ko"] = m.group(1).strip()
    
    m = re.search(r'字\s*([^\.,\n]+)', first_cn)
    if m: details["자"]["cn"] = m.group(1).strip()
    m = re.search(r'初字\s*([^\.,\n]+)', first_cn)
    if m: details["초자"]["cn"] = m.group(1).strip()
    m = re.search(r'號\s*([^\.,\n]+)', first_cn)
    if m: details["호"]["cn"] = m.group(1).strip()
    m = re.search(r'初諱\s*([^\.,\n]+)', first_cn)
    if m: details["초휘"]["cn"] = m.group(1).strip()
    m = re.search(r'一諱\s*([^\.,\n]+)', first_cn)
    if m: details["일휘"]["cn"] = m.group(1).strip()
    m = re.search(r'通名\s*([^\.,\n]+)', first_cn)
    if m: details["통명"]["cn"] = m.group(1).strip()
    m = re.search(r'初名\s*([^\.,\n]+)', first_cn)
    if m: details["초명"]["cn"] = m.group(1).strip()
    
    for line in ko_bio.split("\n"):
        if line.startswith("과거·관직 :"):
            details["과거_관직"]["ko"] = line.replace("과거·관직 :", "").strip()
        elif line.startswith("학문·인품·업적 :"):
            details["학문_업적"]["ko"] = line.replace("학문·인품·업적 :", "").strip()
        elif line.startswith("배 :") or line.startswith("배위 :"):
            details["배위"]["ko"] = re.sub(r'^(배|배위)\s*:\s*', '', line).strip()
            
    for line in cn_bio.split("\n"):
        if line.startswith("科·官 :"):
            details["과거_관직"]["cn"] = line.replace("科·官 :", "").strip()
        elif line.startswith("學·行·業 :"):
            details["학문_업적"]["cn"] = line.replace("學·行·業 :", "").strip()
        elif line.startswith("配 :") or line.startswith("配位 :"):
            details["배위"]["cn"] = re.sub(r'^(配|配位)\s*:\s*', '', line).strip()
            
    return details

def format_bilingual_paragraph_html(row, search_term=None):
    """
    제4장 5항 ①, ②, ④: 표를 사용하지 않고 일반 문단 형식으로 작성하며,
    내용이 없는 항목은 제목도 표시하지 않음.
    """
    det = parse_person_bilingual_details(row)
    lines_html = []
    
    for key, item in det.items():
        ko_val = str(item.get("ko", "")).strip()
        cn_val = str(item.get("cn", "")).strip()
        
        if not ko_val and not cn_val:
            continue
        if ko_val == "-" and not cn_val:
            continue
            
        if search_term:
            ko_val = highlight_corrections_html(ko_val, search_term)
            cn_val = highlight_corrections_html(cn_val, search_term)
            
        lbl = item["label"]
        if ko_val and cn_val:
            bilingual_str = f"{ko_val} ({cn_val})"
        elif ko_val:
            bilingual_str = ko_val
        else:
            bilingual_str = cn_val
            
        lines_html.append(
            f"<div class='meta-para-item'><span class='meta-para-label'>• {lbl} :</span> <span class='meta-para-value'>{bilingual_str}</span></div>"
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

def render_tree_hierarchy(target_row, all_df, max_up=3, max_down=2, key_prefix="tree"):
    """
    제4장 5항 ⑤, ⑥:
    - 계보도 도형 간 간격 최소화 및 도형 간 연결선(│, ▼, ──) 표시
    - 계보도 도형을 클릭하면 검색창 입력없이 바로 새로운 계보 생성
    """
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

    # 연결선이 포함된 통합 계보도 다이어그램 조립
    diagram_parts = [f'<div class="tree-diagram-wrapper"><div style="font-weight:800; color:#1e3c72; font-size:12pt; margin-bottom:6px;">🌳 [{target_name}] 기준 혈통 계보도 (연결선 표시)</div>']
    
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

    # 본인 및 형제
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

    # 자(Sons) 및 손(Grandsons)
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

    # ⭐ 제4장 5항 ⑤: 계보도 도형을 클릭하면 검색창 입력없이 바로 새로운 계보 생성토록 할 것
    st.markdown("##### 👆 계보도 도형 클릭 (클릭 시 검색창 입력 없이 즉시 해당 인물의 새 계보 생성 및 약력 표시)")
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

    # 선택된 기준 인물의 핵심 약력 문단 표시 (제목에 kr, cn 사용 금지)
    meta_para = format_bilingual_paragraph_html(target_row)
    render_html(f"""
    <div class="result-card">
        <div class="result-card-header">
            <div class="result-name-big">
                <span>{target_row['인명']}</span>
                <span class="result-badge">{target_row['관계']}</span>
            </div>
            <div style="font-size:10.5pt; color:#64748b;">고유ID: {target_row['고유ID']}</div>
        </div>
        <div class="meta-para-container">
            {meta_para}
        </div>
        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px;">
            <div class="bio-detail-container">
                <div class="bio-section-title">[한글 약력 정리]</div>
                {format_structured_bio(target_row.get('한글 약력 정리', ''), is_hanmun=False)}
            </div>
            <div class="bio-detail-container">
                <div class="bio-section-title">[한문 약력 정리]</div>
                {format_structured_bio(target_row.get('한문 약력 정리', ''), is_hanmun=True)}
            </div>
        </div>
    </div>
    """)

# =============================================================================
# 4. 상단 헤더 (제5장 최상단 밀착 & 제7장 영천 완귀정 소개 삭제 반영)
# =============================================================================
icon_img_tag = f'<img src="{icon_b64}" style="width:28px; height:28px; border-radius:50%; vertical-align:middle;">' if icon_b64 else '📜'

# 저장소 배경 이미지가 있을 경우 주입 (투명율 0)
bg_style_tag = f"<style>.stApp {{ background-image: url('{bg_b64}') !important; background-size: cover !important; background-attachment: fixed !important; opacity: 1 !important; }}</style>" if bg_b64 else ""
if bg_style_tag:
    render_html(bg_style_tag)

# ⭐ 제5장 1항: 제일 위에 족보검색 시스템 상자를 상단 여백없이 밀착
# 좌측 최상단: 아이콘 + '광주안씨 완귀공파 종중'
# 그 오른쪽 한행 아래 단에 '족보 검색 시스템'을 가로 중앙에 배치
# 그 아래에 '- 완귀공(휘 嶒, 24세) ~ 44세'<br>'총 2,887인(고위 기준) -' 표시
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
""")

# (제7장 1항: 영천 완귀정 소개 상단 영역 삭제 완료)

# =============================================================================
# 5. 메인 네비게이션 메뉴 (제1장 초기 총괄 메뉴 & 제5장 묘제일 메뉴)
# =============================================================================
nav_tab1, nav_tab2, nav_tab3 = st.tabs(["🔍 족보 검색 (인명·계보·단어)", "🌳 계보도 전용 조회", "🕯️ 묘제일 안내 (향사 및 시제)"])

if df is None:
    st.error(f"데이터 로드 오류: {load_msg}")
    st.info("파일 이름이 `genealogy.xlsx`인지 확인해 주시기 바랍니다.")
    st.stop()

# =============================================================================
# TAB 1: 족보 검색 (제1장, 제4장 4~5항, 제6장 하단 부 메뉴)
# =============================================================================
with nav_tab1:
    st.markdown("### 족보 검색")
    render_html("<div style='font-size:11pt; color:#e2e8f0; margin-bottom:4px; font-weight:600;'>아래 검색 방식 중 하나를 선택하세요</div>")
    
    # 제4장 4항 ②: 검색 방식 : 인명, 계보, 단어 로 구분
    search_type = st.radio("검색 종류 선택", ["인명 검색", "계보 검색", "단어 검색"], horizontal=True, label_visibility="collapsed")
    
    # 제4장 4항 ③, ⑤, ⑥: 입력 검색창 너비 최소화(짙은 테두리) + 우측에 검색 방식별 동적 안내문 표시
    c_input, c_guide = st.columns([1.0, 2.5])
    
    with c_input:
        if search_type == "인명 검색":
            query_val = st.text_input("검색어 입력", placeholder="예: 경시, 안경시, 嶒, 완귀", key="name_query", label_visibility="collapsed")
        elif search_type == "계보 검색":
            query_val = st.text_input("검색어 입력", placeholder="예: 嶒, 명한, 안경시", key="lineage_query", label_visibility="collapsed")
        else:
            query_val = st.text_input("검색어 입력", placeholder="예: 시강원, 헌릉참봉, 1699", key="word_query", label_visibility="collapsed")
            
    with c_guide:
        # ⭐ 제4장 4항 ⑥: 인물, 계보, 단어를 선택할 때마다 새로운 안내문이 나타나도록 구현
        if search_type == "인명 검색":
            render_html("""
            <div class="search-guide-right-text">
                💡 <b>[인명 검색 안내]</b> 인물의 함자(성씨 안을 넣어도 무방), 자(字), 초자(初字), 호(號), 초휘(初諱), 일휘(一諱) 통명(通名), 초명(初名) 등을 한글로 입력
            </div>
            """)
        elif search_type == "계보 검색":
            render_html("""
            <div class="search-guide-right-text">
                🌳 <b>[계보 검색 안내]</b> 계보를 조회할 인물의 함자(성씨 안 포함 가능), 자(字), 호(號), 초휘(初諱) 등을 입력 (생성된 계보도 도형 클릭 시 즉시 새 계보 생성)
            </div>
            """)
        else:
            render_html("""
            <div class="search-guide-right-text">
                🔍 <b>[단어 검색 안내]</b> 약력에 포함된 관직명, 과거 급제, 지명, 묘소 좌향, 시호, 연대 등의 단어를 입력 (검색어는 <span style="color:#dc2626;">빨간색</span>으로 표시됨)
            </div>
            """)

    # =========================================================================
    # ⭐ 제4장 5항 ③: 검색 결과 창은 항상 검색어 입력 안내문 바로 다음에 생성시킬 것!
    # =========================================================================
    
    # -------------------------------------------------------------------------
    # A. 인명 검색 결과창 (제4장 5항 ①, ②, ③, ④, ⑧, ⑨)
    # -------------------------------------------------------------------------
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
                
            h_name, c_name = split_name(target_row['인명'])
            meta_paragraphs_html = format_bilingual_paragraph_html(target_row)
            
            render_html(f"""
            <div class="result-card">
                <div class="result-card-header">
                    <div class="result-name-big">
                        <span>{h_name if h_name else target_row['인명']}</span>
                        <span style="font-family:'Nanum Myeongjo',serif; font-size:16pt; color:#4b5563;">{c_name}</span>
                        <span class="result-badge">{target_row['관계']}</span>
                    </div>
                    <div style="font-size:10.5pt; color:#6b7280; font-family:monospace;">고유ID: {target_row['고유ID']}</div>
                </div>
                
                <div style="font-weight:700; color:#1e3c72; margin-bottom:4px; font-size:11.5pt;">
                    📌 인물 기본 신상 및 계통 정보 (한글·한자 병기 문단)
                </div>
                <div class="meta-para-container">
                    {meta_paragraphs_html}
                </div>
                
                <div style="font-weight:700; color:#1e3c72; margin-top:10px; margin-bottom:6px; font-size:11.5pt;">
                    📜 상세 한문·한글 약력
                </div>
                
                <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px;">
                    <div class="bio-detail-container">
                        <div class="bio-section-title">[한글 약력 정리]</div>
                        {format_structured_bio(target_row.get('한글 약력 정리', ''), is_hanmun=False)}
                    </div>
                    <div class="bio-detail-container">
                        <div class="bio-section-title">[한문 약력 정리]</div>
                        {format_structured_bio(target_row.get('한문 약력 정리', ''), is_hanmun=True)}
                    </div>
                </div>
            </div>
            """)
            
            with st.expander("📖 [족보 수록 원문 보기 (한글 원문 / 한문 원문 대조)]", expanded=False):
                c_won1, c_won2 = st.columns(2)
                with c_won1:
                    render_html("<b style='color:#1e3c72;'>[한글 약력 원문]</b>")
                    st.text_area("한글 원문", target_row.get('한글 약력 원문', ''), height=150, disabled=True, label_visibility="collapsed")
                with c_won2:
                    render_html("<b style='color:#1e3c72;'>[한문 약력 원문]</b>")
                    st.text_area("한문 원문", target_row.get('한문 약력 원문', ''), height=150, disabled=True, label_visibility="collapsed")

    # -------------------------------------------------------------------------
    # B. 계보 검색 결과창 (제2장 & 제4장 5항 ⑤, ⑥)
    # -------------------------------------------------------------------------
    elif search_type == "계보 검색":
        # 검색어가 새로 입력되면 검색어 기준, 도형 클릭 시 클릭된 인물 기준 즉시 계보 생성
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
                sel_cand = st.selectbox("계보 기준 인물 선택", cand_opts)
                target_gid = sel_cand.split(" - ")[-1].strip()
                tree_target = l_matches[l_matches['고유ID'] == target_gid].iloc[0]
            else:
                tree_target = l_matches.iloc[0]

        if tree_target is not None:
            st.markdown("#### ⚙️ 계보 작성 조건 창")
            c_cond1, c_cond2, c_cond3 = st.columns([1.2, 1.2, 1])
            with c_cond1:
                up_gens = st.slider("상향 세대수 (선대 조상)", min_value=1, max_value=5, value=3)
            with c_cond2:
                down_gens = st.slider("하향 세대수 (후손 자손)", min_value=1, max_value=4, value=2)
            with c_cond3:
                if st.session_state.get('active_tree_gid'):
                    st.write("")
                    if st.button("🔄 초기 검색 인물로 복귀", use_container_width=True):
                        st.session_state['active_tree_gid'] = None
                        st.rerun()
            render_tree_hierarchy(tree_target, df, max_up=up_gens, max_down=down_gens, key_prefix="tab1_tree")

    # -------------------------------------------------------------------------
    # C. 단어 검색 결과창 (제4장 5항 ⑦, ⑧, ⑨)
    # -------------------------------------------------------------------------
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
            # ⭐ 제4장 5항 ⑦: 검색어는 빨간색으로 표시하고 해당 단어가 있는 약력 선조가 다수 있음을 확실하게 표시
            render_html(f"""
            <div class="multi-ancestor-banner">
                📢 검색어 <span class="search-word-red">'{term}'</span>을(를) 포함하는 약력을 가진 선조가 총 <b>{len(word_matches):,}위(다수)</b> 검색되었습니다. 아래 목록에서 선조를 선택하시면 상세 약력이 표시됩니다.
            </div>
            """)
            candidate_list = []
            for _, r in word_matches.iterrows():
                gen_num, gen_str = extract_generation_info(r['고유ID'])
                candidate_list.append(f"[{gen_str}] {r['인명']} (부: {r['부친']}, 관계: {r['관계']}) - 고유ID: {r['고유ID']}")
                
            chosen = st.selectbox("조회할 선조를 선택하세요", candidate_list, label_visibility="collapsed")
            chosen_gid = chosen.split("고유ID: ")[-1].strip()
            chosen_row = word_matches[word_matches['고유ID'] == chosen_gid].iloc[0]
            
            meta_paragraphs_html = format_bilingual_paragraph_html(chosen_row, search_term=term)
            
            render_html(f"""
            <div class="result-card">
                <div class="result-card-header">
                    <div class="result-name-big">
                        <span>{chosen_row['인명']}</span>
                        <span class="result-badge">{chosen_row['관계']}</span>
                    </div>
                    <div style="font-size:10.5pt; color:#6b7280;">고유ID: {chosen_row['고유ID']}</div>
                </div>
                <div class="meta-para-container">
                    {meta_paragraphs_html}
                </div>
                <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px;">
                    <div class="bio-detail-container">
                        <div class="bio-section-title">[한글 약력 정리]</div>
                        {format_structured_bio(chosen_row.get('한글 약력 정리', ''), is_hanmun=False, search_term=term)}
                    </div>
                    <div class="bio-detail-container">
                        <div class="bio-section-title">[한문 약력 정리]</div>
                        {format_structured_bio(chosen_row.get('한문 약력 정리', ''), is_hanmun=True, search_term=term)}
                    </div>
                </div>
            </div>
            """)
            
            with st.expander("📖 [족보 수록 원문 보기 (한글 원문 / 한문 원문 대조)]"):
                c_w1, c_w2 = st.columns(2)
                with c_w1:
                    render_html("<b style='color:#1e3c72;'>[한글 약력 원문]</b>")
                    st.text_area("한글 원문", chosen_row.get('한글 약력 원문', ''), height=150, disabled=True, label_visibility="collapsed")
                with c_w2:
                    render_html("<b style='color:#1e3c72;'>[한문 약력 원문]</b>")
                    st.text_area("한문 원문", chosen_row.get('한문 약력 원문', ''), height=150, disabled=True, label_visibility="collapsed")

    # =========================================================================
    # ⭐ 제6장: 검색 창 하단 부 메뉴 생성 (검색 결과창 아래 하단에 배치)
    # =========================================================================
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
# TAB 2: 계보도 전용 조회 (제2장)
# =============================================================================
with nav_tab2:
    st.markdown("### 🌳 완귀공파 계보도 직접 탐색")
    quick_names = ["嶒증", "宗慶종경", "大海대해", "琠전", "鳴漢명한", "汝履여리", "景時경시"]
    sel_q = st.selectbox("대표 선조 계보 바로보기", quick_names)
    
    if sel_q != st.session_state.get('last_tab2_sel', ''):
        st.session_state['last_tab2_sel'] = sel_q
        st.session_state['active_tab2_gid'] = None

    tab2_target = None
    if st.session_state.get('active_tab2_gid'):
        t2_m = df[df['고유ID'] == st.session_state['active_tab2_gid']]
        if not t2_m.empty:
            tab2_target = t2_m.iloc[0]

    if tab2_target is None:
        q_match = df[df['인명'].astype(str).str.contains(str(sel_q), na=False)]
        if not q_match.empty:
            tab2_target = q_match.iloc[0]

    if tab2_target is not None:
        render_tree_hierarchy(tab2_target, df, key_prefix="tab2_tree")

# =============================================================================
# TAB 3: 묘제일 메뉴 생성 (제5장 묘제일 안내)
# =============================================================================
with nav_tab3:
    st.markdown("### 🕯️ 광주안씨 완귀공파 묘제일(향사 및 시제) 상세 안내")
    
    render_html("""
    <div class="result-card" style="border-top-color: #8c2d19;">
        <div style="font-size:14pt; font-weight:800; color:#8c2d19; margin-bottom:8px;">📅 완귀공 종중 정례 제향 및 시제 일정</div>
        <div class="meta-para-container" style="background:#ffffff; border:1px solid #cbd5e1;">
            <div class="meta-para-item"><b>1. 구암리 상로재(霜露齋) 춘향대제 :</b> 매년 양력 4월 2째주 일요일 (경북 영천시 구암리 상로재, 완귀공 후손 종원 일동)</div>
            <div class="meta-para-item"><b>2. 청제내 선영 묘제 :</b> 매년 춘향대제 봉행 직후 (영천시 금호읍 청제내 화전등, 완귀공 및 배위 영양최씨 묘소)</div>
            <div class="meta-para-item"><b>3. 작산 선영 시향(묘제) :</b> 매년 음력 10월 상정일 (경북 영천시 작산 묘역 안장 각 선조위)</div>
            <div class="meta-para-item"><b>4. 각 지파별 시제 :</b> 음력 10월 중 각 지파별 길일 (효식·효빈·효욱·효갑·효방 등 각 지파 재실 및 선영)</div>
        </div>
        <div class="bio-detail-container" style="background:#ffffff; margin-top:10px;">
            <div class="bio-section-title">🏛️ 전통 제례 절차 및 종원 수칙 안내</div>
            <ol style="line-height:1.75; padding-left:20px; font-size:10.5pt; color:#1e293b; margin:0;">
                <li><b>강신(降神) :</b> 제주가 묘소 또는 신위 앞에 나아가 향을 피우고 잔에 술을 부어 퇴주기에 세 번 나누어 붓고 뇌주하여 지하의 신을 모십니다.</li>
                <li><b>참신(參神) :</b> 신주를 뵙는 절차로 참사자 전원이 재배(再拜)합니다.</li>
                <li><b>초헌(初獻) :</b> 종손(또는 유사)이 첫 번째 잔을 올리고 메(밥)의 뚜껑을 엽니다.</li>
                <li><b>독축(讀祝) :</b> 축관이 꿇어앉아 정성껏 준비한 축문을 엄숙하게 낭독합니다. (낭독 중 참사자 부복)</li>
                <li><b>아헌(亞獻) :</b> 두 번째 잔을 올리는 절차로 종중의 차례 맡은 원로가 잔을 올리고 재배합니다.</li>
                <li><b>종헌(終獻) :</b> 세 번째 잔을 올리는 절차로 잔을 가득 붓지 않고 조금 남겨 올립니다.</li>
                <li><b>유식(侑食) :</b> 신위께서 음식을 흠향하시도록 숟가락을 메에 꽂고 젓가락을 고기나 떡 위에 올려놓습니다.</li>
                <li><b>사신(辭神) :</b> 향사를 마치고 신위를 전송하는 절차로 참사자 전원이 재배합니다.</li>
                <li><b>분축 및 음복(焚祝·飮福) :</b> 축문을 불사르고 선조께서 내려주신 제수와 제주를 나누어 먹으며 종중 간의 화목과 번영을 기원합니다.</li>
            </ol>
        </div>
    </div>
    """)

# =============================================================================
# 6. 하단 푸터 (Footer)
# =============================================================================
render_html("""
<div style="text-align:center; padding:12px 0 6px 0; color:#cbd5e1; font-size:9.5pt; border-top:1px solid #334155; margin-top:20px;">
    광주안씨 완귀공파 종중 회의소 | 데이터 출처: 보정 족보 1~2권 통합본 (총 2,887위) | 연동 파일: genealogy.xlsx
</div>
""")
