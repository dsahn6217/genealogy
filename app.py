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
# 1. 페이지 설정 및 전통 서책 스타일 CSS (제3장 디자인 & 제5장 최상단 밀착/투명도0)
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
    padding-top: 0rem !important;  /* 👈 최상단 여백 없이 밀착 */
    padding-bottom: 2rem !important;
    padding-left: 1rem !important;
    padding-right: 1rem !important;
}

/* 제5장 2항: 저장소 배경 이미지 적용 및 투명도 제로(100% 완전 불투명) */
.stApp {
    background-color: #faf8f5;
    background-image: radial-gradient(#e5dec9 0.75px, transparent 0.75px);
    background-size: 24px 24px;
    opacity: 1 !important;         /* 👈 투명율 0 (완전 불투명) */
}

/* 제5장 1항: 제일 위 족보검색 시스템 네모 상자 (최상단 밀착, 불투명 백색) */
.title-container {
    width: 100% !important;
    box-sizing: border-box !important;
    margin-top: 0px !important;    /* 👈 상단 여백없이 밀착 */
    margin-bottom: 8px !important;
    padding: 10px 14px 10px 14px;
    border-bottom: 2px solid #8c2d19;
    background: #ffffff !important;
    opacity: 1 !important;         /* 👈 투명도 제로 */
    border-radius: 8px;
    box-shadow: 0 3px 10px rgba(0,0,0,0.05);
}

/* 여백없이 좌측 최상단에 아이콘 + '광주안씨 완귀공파 종중' */
.title-top-left {
    display: flex;
    align-items: center;
    justify-content: flex-start;
    gap: 8px;
    margin-bottom: 2px;
}
.title-clan-name {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
    font-size: 13.5pt;
    font-weight: 700;
    color: #1e3c72;
    letter-spacing: -0.3px;
}

/* 그 오른쪽 한행 아래 단에 '족보 검색 시스템'을 가로 중앙에 배치 */
.title-main-center {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
    font-size: 18pt;
    font-weight: 800;
    color: #1a202c;
    text-align: center;
    width: 100%;
    margin-top: 2px;
    margin-bottom: 4px;
    letter-spacing: -0.5px;
}

/* 그 아래에 '- 완귀공(휘 嶒, 24세) ~ 44세' '총 2,887인(고위 기준) -' 표시 */
.title-sub-bottom {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
    font-size: 11pt;
    font-weight: 500;
    color: #4b5563;
    text-align: center;
    width: 100%;
    margin-top: 2px;
    line-height: 1.4;
    letter-spacing: -0.2px;
}

/* 완귀정 소개 펼치기(Expander) 토글 버튼 */
[data-testid="stExpander"] {
    border: 1px solid #cbd5e1 !important;
    border-radius: 8px !important;
    margin-bottom: 12px !important;
    background: #ffffff !important;
    opacity: 1 !important;
}
[data-testid="stExpander"] summary {
    font-size: 12pt !important;
    font-weight: 700 !important;
    color: #1e3c72 !important;
    padding: 8px 14px !important;
}
[data-testid="stExpander"] summary p,
[data-testid="stExpander"] summary span {
    font-size: 12pt !important;
    font-weight: 700 !important;
    color: #1e3c72 !important;
}

/* 완귀정 소개 카드 상자 */
.wangwi-intro-box {
    width: 100% !important;
    box-sizing: border-box !important;
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-left: 5px solid #1e3c72;
    border-radius: 8px;
    padding: 14px 18px;
    margin-bottom: 12px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.03);
    line-height: 1.6;
}
.wangwi-intro-title {
    font-size: 14pt;
    font-weight: 800;
    color: #1e3c72;
    margin-bottom: 10px;
    display: flex;
    align-items: center;
    gap: 8px;
}
.wangwi-intro-item {
    font-size: 11pt;
    color: #2d3748;
    margin-bottom: 6px;
    line-height: 1.65;
    word-break: keep-all;
}
.wangwi-intro-label {
    font-size: 11pt;
    font-weight: 700;
    color: #1e3c72;
    display: inline-block;
    min-width: 65px;
}

/* 제6장: 검색창 하단 부 메뉴 (종중 유적 문화재 안내) 스타일 */
.heritage-sub-container {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    padding: 14px 16px;
    margin-top: 15px;
    margin-bottom: 15px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.03);
}
.heritage-sub-header {
    font-size: 12pt;
    font-weight: 700;
    color: #1e3c72;
    margin-bottom: 10px;
    display: flex;
    align-items: center;
    gap: 6px;
}
.heritage-sub-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
    gap: 10px;
}
.heritage-sub-card {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    padding: 10px 12px;
    text-decoration: none !important;
    display: block;
    transition: all 0.2s ease;
}
.heritage-sub-card:hover {
    background: #eff6ff;
    border-color: #93c5fd;
    transform: translateY(-1px);
}
.h-card-name {
    font-size: 11pt;
    font-weight: 700;
    color: #1e40af;
    margin-bottom: 3px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.h-card-desc {
    font-size: 9.5pt;
    color: #475569;
    line-height: 1.4;
}

/* 검색 가이드 문구 */
.search-guide-right-text {
    font-size: 10pt;
    color: #0369a1;
    background-color: #f0f9ff;
    padding: 6px 12px;
    border-radius: 6px;
    border: 1px solid #bae6fd;
    line-height: 1.45;
}

/* 제4장 5항: 검색 결과 인물 카드 */
.result-card {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-top: 4px solid #1e3c72;
    border-radius: 10px;
    padding: 22px;
    margin-top: 20px;
    margin-bottom: 25px;
    box-shadow: 0 4px 15px rgba(0,0,0,0.05);
}
.result-card-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 2px solid #edf2f7;
    padding-bottom: 12px;
    margin-bottom: 15px;
}
.result-name-big {
    font-size: 20pt;
    font-weight: 800;
    color: #1e3c72;
    display: flex;
    align-items: center;
    gap: 10px;
}
.result-badge {
    font-size: 11pt;
    font-weight: 600;
    background: #eff6ff;
    color: #1d4ed8;
    padding: 3px 12px;
    border-radius: 15px;
    border: 1px solid #bfdbfe;
}

/* ⭐ 제4장 5항 ②: 결과 창은 여백을 줄이기 위해 표로 작성하지 않고 문단으로 표시 */
.meta-para-container {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 14px 16px;
    margin-bottom: 18px;
    line-height: 1.8;
    font-size: 11pt;
}
.meta-para-item {
    margin-bottom: 6px;
    word-break: keep-all;
}
.meta-para-label {
    font-weight: 700;
    color: #1e3c72;
    margin-right: 6px;
}
.meta-para-value {
    color: #1e293b;
}

/* 5단락 상세 약력 박스 */
.bio-detail-container {
    background-color: #fcfcfd;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 16px 20px;
    margin-bottom: 18px;
    line-height: 1.85;
    font-size: 11.5pt;
    color: #2d3748;
}
/* ⭐ 제4장 5항 ③: 제목에 kr, cn 사용금지 */
.bio-section-title {
    font-size: 12.5pt;
    font-weight: 700;
    color: #1e3c72;
    margin-bottom: 10px;
    padding-bottom: 4px;
    border-bottom: 1px solid #e2e8f0;
    display: flex;
    align-items: center;
    gap: 6px;
}
.bio-tag {
    font-weight: 700;
    color: #1e40af;
    background: #eff6ff;
    padding: 2px 8px;
    border-radius: 4px;
    margin-right: 6px;
    font-size: 10.5pt;
    display: inline-block;
}
.red-correction {
    color: #dc2626 !important;
    font-weight: 700;
    background-color: #fef2f2;
    padding: 1px 5px;
    border-radius: 3px;
    border-bottom: 1px dashed #f87171;
}

/* 계보도 지그재그 행 */
.tree-level-row {
    display: flex;
    align-items: center;
    margin-bottom: 12px;
    padding: 8px 12px;
    background: #ffffff;
    border-radius: 8px;
    border: 1px solid #e2e8f0;
}
.tree-gen-badge {
    min-width: 90px;
    font-weight: 700;
    color: #1e3c72;
    font-size: 11pt;
    padding: 4px 8px;
    background: #eff6ff;
    border-radius: 4px;
    text-align: center;
    margin-right: 14px;
}
.tree-nodes-container {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    align-items: center;
}
.tree-node-box {
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 6px 14px;
    font-size: 11pt;
    color: #1e293b;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}
.tree-node-target {
    background: #eff6ff;
    border: 2px solid #2563eb;
    font-weight: 700;
    color: #1e40af;
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
# 3. 보조 유틸리티 함수: 세수, 인명 분리, 고증 빨간색 하이라이트, 계보도 렌더링
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

def highlight_corrections_html(text):
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
    
    # 한글 파싱
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
    
    # 한자 파싱
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
    
    # 과거·관직 및 학문·업적, 배위
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

def format_bilingual_paragraph_html(row):
    """
    ⭐ 제4장 5항 ①, ② : 표 대신 문단으로 표시, 내용이 없는 항목은 완전히 생략
    """
    det = parse_person_bilingual_details(row)
    lines_html = []
    
    for key, item in det.items():
        ko_val = str(item.get("ko", "")).strip()
        cn_val = str(item.get("cn", "")).strip()
        
        # 항목의 내용이 없는 경우 항목도 표시하지 않는다
        if not ko_val and not cn_val:
            continue
        if ko_val == "-" and not cn_val:
            continue
            
        lbl = item["label"]
        if ko_val and cn_val:
            bilingual_str = f"{ko_val} &nbsp;(<span style='color:#475569;'>{cn_val}</span>)"
        elif ko_val:
            bilingual_str = ko_val
        else:
            bilingual_str = cn_val
            
        lines_html.append(
            f"<div class='meta-para-item'><span class='meta-para-label'>• {lbl} :</span> <span class='meta-para-value'>{bilingual_str}</span></div>"
        )
        
    return "".join(lines_html)

def format_structured_bio(bio_text, is_hanmun=False):
    if not bio_text:
        return '<div style="color:#94a3b8; font-style:italic;">수록된 정리 약력이 없습니다.</div>'
    lines = [ln.strip() for ln in str(bio_text).split("\n") if ln.strip()]
    html_items = []
    for ln in lines:
        hl = highlight_corrections_html(ln)
        if ln.startswith("묘 :") or ln.startswith("墓 :"):
            lbl = "墓" if is_hanmun else "묘"
            html_items.append(f'<div><span class="bio-tag">🪦 {lbl}</span> {hl[3:].strip()}</div>')
        elif ln.startswith("과거·관직 :") or ln.startswith("科·官 :"):
            lbl = "科·官" if is_hanmun else "과거·관직"
            cut = 5 if is_hanmun else 6
            html_items.append(f'<div><span class="bio-tag">🏛️ {lbl}</span> {hl[cut:].strip()}</div>')
        elif ln.startswith("학문·인품·업적 :") or ln.startswith("學·行·業 :"):
            lbl = "學·行·業" if is_hanmun else "학문·인품·업적"
            cut = 6 if is_hanmun else 9
            html_items.append(f'<div><span class="bio-tag">📖 {lbl}</span> {hl[cut:].strip()}</div>')
        elif ln.startswith("배 :") or ln.startswith("配 :"):
            lbl = "配" if is_hanmun else "배"
            html_items.append(f'<div><span class="bio-tag">💍 {lbl}</span> {hl[3:].strip()}</div>')
        else:
            lbl = "身常" if is_hanmun else "신상"
            html_items.append(f'<div><span class="bio-tag">👤 {lbl}</span> {hl}</div>')
    return "".join(html_items)

def render_tree_hierarchy(target_row, all_df, max_up=3, max_down=2):
    target_gid = target_row['고유ID']
    target_name = target_row['인명']
    
    st.markdown(f"#### 🌳 [{target_name}] 기준 수평 지그재그 계보도")
    
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
        
    for anc in reversed(ancestors):
        lvl = anc['level']
        r = anc['row']
        g_num, g_str = extract_generation_info(r['고유ID'])
        sibs = anc['siblings']
        
        if lvl >= 3:
            rel_str = f" ({len(sibs)}자 중 {r['관계']})" if len(sibs) > 1 else f" ({r['관계']})"
            lbl_badge = "증조부" if lvl == 3 else f"{lvl}대조"
            render_html(f"""
            <div class="tree-level-row">
                <span class="tree-gen-badge">▲ {lbl_badge} ({g_num}세)</span>
                <div class="tree-nodes-container">
                    <span class="tree-node-box"><b>{r['인명']}</b>{rel_str}</span>
                </div>
            </div>
            """)
        else:
            lbl_badge = "부친(父)" if lvl == 1 else "조부(祖)"
            node_htmls = []
            for _, s in sibs.iterrows():
                is_direct = (s['고유ID'] == r['고유ID'])
                cls = "tree-node-box tree-node-target" if is_direct else "tree-node-box"
                s_name = s['인명']
                s_rel = s['관계']
                node_htmls.append(f'<span class="{cls}"><b>{s_name}</b> ({s_rel})</span>')
            nodes_joined = ''.join(node_htmls)
            render_html(f"""
            <div class="tree-level-row">
                <span class="tree-gen-badge">▲ {lbl_badge} ({g_num}세)</span>
                <div class="tree-nodes-container">{nodes_joined}</div>
            </div>
            """)
            
    t_f = str(target_row.get('부친', '')).strip()
    t_sibs = all_df[all_df['부친'] == t_f] if t_f else pd.DataFrame([target_row])
    tg_num, tg_str = extract_generation_info(target_row['고유ID'])
    
    t_nodes = []
    for _, s in t_sibs.iterrows():
        is_t = (s['고유ID'] == target_row['고유ID'])
        cls = "tree-node-box tree-node-target" if is_t else "tree-node-box"
        marker = " 🎯 [기준]" if is_t else ""
        s_name = s['인명']
        s_rel = s['관계']
        t_nodes.append(f'<span class="{cls}"><b>{s_name}</b>{marker} ({s_rel})</span>')
    t_joined = ''.join(t_nodes)
    render_html(f"""
    <div class="tree-level-row" style="background:#eff6ff; border:2px solid #3b82f6;">
        <span class="tree-gen-badge" style="background:#2563eb; color:white;">● 본인 ({tg_num}세)</span>
        <div class="tree-nodes-container">{t_joined}</div>
    </div>
    """)
    
    sons = all_df[all_df['부친'] == target_name]
    if len(sons) > 0:
        s_nodes = []
        for _, s in sons.iterrows():
            s_name = s['인명']
            s_rel = s['관계']
            s_nodes.append(f'<span class="tree-node-box"><b>{s_name}</b> ({s_rel})</span>')
        s_joined = ''.join(s_nodes)
        sg_num = tg_num + 1 if tg_num else 25
        render_html(f"""
        <div class="tree-level-row">
            <span class="tree-gen-badge">▼ 자(子) ({sg_num}세)</span>
            <div class="tree-nodes-container">{s_joined}</div>
        </div>
        """)
        
        if max_down >= 2:
            grandsons = []
            for _, s in sons.iterrows():
                gs = all_df[all_df['부친'] == s['인명']]
                for _, g in gs.iterrows():
                    g_name = g['인명']
                    s_name = s['인명']
                    g_rel = g['관계']
                    grandsons.append(f'<span class="tree-node-box"><b>{g_name}</b> ({s_name}의 {g_rel})</span>')
            if grandsons:
                g_joined = ''.join(grandsons)
                render_html(f"""
                <div class="tree-level-row">
                    <span class="tree-gen-badge">▼ 손(孫) ({sg_num+1}세)</span>
                    <div class="tree-nodes-container">{g_joined}</div>
                </div>
                """)
                
    st.write("")
    st.markdown("##### 🔍 계보 인물 약력 바로보기")
    related_people = [target_row] + [anc['row'] for anc in ancestors] + [s for _, s in sons.iterrows()]
    p_opts = {f"{p['인명']} ({p['관계']}, ID: {p['고유ID']})": p for p in related_people}
    selected_p_key = st.selectbox("약력을 열람할 계보 선조 선택", list(p_opts.keys()))
    sel_p_row = p_opts[selected_p_key]
    
    render_html(f"""
    <div class="bio-detail-container" style="margin-top:10px;">
        <div class="bio-section-title">👤 [{sel_p_row['인명']}] 선조 핵심 약력</div>
        {format_structured_bio(sel_p_row.get('한글 약력 정리', ''), is_hanmun=False)}
    </div>
    """)

# =============================================================================
# 4. 상단 헤더 및 완귀정 문화재 소개 (제5장 최상단 밀착 & 제3장 동영상)
# =============================================================================
icon_img_tag = f'<img src="{icon_b64}" style="width:30px; height:30px; border-radius:50%; vertical-align:middle;">' if icon_b64 else '📜'

# 저장소 배경 이미지가 있을 경우 주입
bg_style_tag = f"<style>.stApp {{ background-image: url('{bg_b64}') !important; background-size: cover !important; background-attachment: fixed !important; }}</style>" if bg_b64 else ""
if bg_style_tag:
    render_html(bg_style_tag)

# ⭐ 제5장 1항: 제일 위에 족보검색 시스템 상자를 상단 여백없이 밀착
# 여백없이 좌측 최상단에 아이콘 배치 + '광주안씨 완귀공파 종중'
# 그 오른쪽 한행 아래 단에 '족보 검색 시스템'을 가로 중앙에 배치
# 그 아래에 '- 완귀공(휘 嶒, 24세) ~ 44세' '총 2,887인(고위 기준) -' 표시
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
        - 완귀공(휘 嶒, 24세) ~ 44세 &nbsp; 총 2,887인(고위 기준) -
    </div>
</div>
""")

# 완귀정 소개 펼치기 영역 (내용 수정 없음)
with st.expander("🏯 [영천 완귀정(경상북도 민속문화재 제20호) 소개 및 문화재 영상 보기]", expanded=False):
    col_intro_1, col_intro_2 = st.columns([1.3, 1])
    with col_intro_1:
        render_html("""
        <div class="wangwi-intro-box">
            <div class="wangwi-intro-title">영천 완귀정(경상북도 민속문화재 제20호)</div>
            <div class="wangwi-intro-item"><span class="wangwi-intro-label">건립자 :</span> 완귀 안증(1494~1553, 玩龜 安嶒) 선생</div>
            <div class="wangwi-intro-item"><span class="wangwi-intro-label">위 치 :</span> 경상북도 영천시 도남동 (虎溪 정자)</div>
            <div class="wangwi-intro-item"><span class="wangwi-intro-label">유 래 :</span> 을사사화(1545)의 기미를 내다보고 벼슬에서 물러나 향리 호계(虎溪) 위에 정자를 짓고, 거북이 여섯 지체를 감추어 난을 피한다는 뜻을 취하여 '완귀(玩龜)'라 편액하였습니다.</div>
            <div class="wangwi-intro-item"><span class="wangwi-intro-label">제 향 :</span> <b>매년 양력 4월 2째주 일요일</b> 구암리 상로재(霜露齋)에서 봉행.</div>
        </div>
        """)
    with col_intro_2:
        st.caption("🎥 영천 완귀정 문화재 탐방 영상 (YouTube)")
        st.video("https://www.youtube.com/watch?v=XyfHLpZrrlo")

# =============================================================================
# 5. 메인 네비게이션 메뉴 (제1장 초기 총괄 메뉴 & 제5장 묘제일 메뉴)
# =============================================================================
nav_tab1, nav_tab2, nav_tab3 = st.tabs(["🔍 족보 검색 (인명·단어·계보)", "🌳 계보도 전용 조회", "🕯️ 묘제일 안내 (향사 및 시제)"])

if df is None:
    st.error(f"데이터 로드 오류: {load_msg}")
    st.info("파일 이름이 `genealogy.xlsx`인지 확인해 주시기 바랍니다.")
    st.stop()

# =============================================================================
# TAB 1: 족보 검색 (제1장, 제4장 4, 5항 & 제6장 하단 부 메뉴)
# =============================================================================
with nav_tab1:
    st.markdown("### 족보 검색")
    render_html("<div style='font-size:12pt; color:#4b5563; margin-bottom:10px;'>아래 검색 방식 중 하나를 선택하세요</div>")
    
    # 제4장 4항 ②: 검색 방식 (인명, 계보, 단어)
    search_type = st.radio("검색 종류 선택", ["인명 검색", "계보 검색", "단어 검색"], horizontal=True, label_visibility="collapsed")
    
    # 제4장 4항 ③: 검색 방식 아래 입력 검색창 너비 최소화 및 우측 안내문 표시
    st.write("")
    c_input, c_guide = st.columns([1.1, 2.4])
    
    with c_input:
        if search_type == "인명 검색":
            query_val = st.text_input("검색어 입력", placeholder="예: 경시, 안경시, 嶒, 완귀, 만원", key="name_query", label_visibility="collapsed")
        elif search_type == "계보 검색":
            query_val = st.text_input("검색어 입력", placeholder="예: 嶒, 명한, 안경시", key="lineage_query", label_visibility="collapsed")
        else:
            query_val = st.text_input("검색어 입력", placeholder="예: 시강원, 헌릉참봉, 1699", key="word_query", label_visibility="collapsed")
            
    with c_guide:
        render_html("""
        <div class="search-guide-right-text">
            💡 <b>검색어 입력 안내 :</b><br>
            인물의 함자(성씨 안을 넣어도 무방), 자(字), 초자(初字), 호(號), 초휘(初諱), 일휘(一諱) 통명(通名), 초명(初名) 등을 한글로 입력하세요.
        </div>
        """)

    # ⭐ 제6장: 검색 창 하단 부 메뉴 생성 (종중 주요 유적 및 문화재 링크 안내)
    render_html("""
    <div class="heritage-sub-container">
        <div class="heritage-sub-header">🏛️ 종중 주요 유적 및 문화재 안내 (클릭 시 국가유산 포털로 이동)</div>
        <div class="heritage-sub-grid">
            <a href="https://www.heritage.go.kr/heri/cul/culSelectDetail.do?pageNo=1_1_1_1&sngl=Y&ccbaCpno=2443700200000#audioDiv" target="_blank" class="heritage-sub-card">
                <div class="h-card-name">1. 영천 완귀정(玩龜亭) <span>↗</span></div>
                <div class="h-card-desc">경상북도 민속문화재 제20호 (도남동 호계)</div>
            </a>
            <a href="https://www.heritage.go.kr/heri/cul/culSelectDetail.do?pageNo=1_1_1_1&sngl=Y&ccbaCpno=3413705940000" target="_blank" class="heritage-sub-card">
                <div class="h-card-name">2. 영천 육우당(六友堂) <span>↗</span></div>
                <div class="h-card-desc">경상북도 영천시 도남동 566 (2012.09.17 지정)</div>
            </a>
            <a href="https://www.heritage.go.kr/heri/cul/culSelectDetail.do?pageNo=1_1_1_1&sngl=N&ccbaAsno=0007220000000&ccbaCpno=3412603100001" target="_blank" class="heritage-sub-card">
                <div class="h-card-name">3. 청도 만회정(晩悔亭) <span>↗</span></div>
                <div class="h-card-desc">경북 청도군 운문면 지촌리 907 (2026.02.19 지정)</div>
            </a>
            <div class="heritage-sub-card">
                <div class="h-card-name">4. 삼괴당(三槐堂)</div>
                <div class="h-card-desc">광주안씨 문중 선조 강학 및 추모 유적</div>
            </div>
        </div>
    </div>
    """)

    # -------------------------------------------------------------------------
    # A. 인명 검색 처리 (제1장 ①, ② & 제4장 5항)
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
            st.warning(f"입력하신 검색어 '{q}'에 해당하는 인물을 찾을 수 없습니다. (성씨 '안'을 제외하거나 한자로 입력해 보세요.)")
        else:
            if len(matched_df) > 1:
                st.info(f"검색 결과 총 {len(matched_df)}위의 인물이 검색되었습니다. 아래 목록에서 조회할 인물을 선택하세요.")
                
                candidate_labels = []
                for idx, row in matched_df.iterrows():
                    gen_num, gen_str = extract_generation_info(row['고유ID'])
                    h_name, c_name = split_name(row['인명'])
                    lbl = f"[{gen_str}] {row['인명']} (부: {row['부친']}, 조부: {row['조부']}, 관계: {row['관계']}) - 고유ID: {row['고유ID']}"
                    candidate_labels.append(lbl)
                
                selected_label = st.selectbox("선조 선택 (동명이인 또는 복수 후보)", candidate_labels)
                selected_gid = selected_label.split("고유ID: ")[-1].strip()
                target_row = matched_df[matched_df['고유ID'] == selected_gid].iloc[0]
            else:
                target_row = matched_df.iloc[0]
                
            h_name, c_name = split_name(target_row['인명'])
            
            # ⭐ 제4장 5항 ①, ② : 표 대신 여백을 줄인 문단 형식, 내용 없는 항목 자동 생략
            meta_paragraphs_html = format_bilingual_paragraph_html(target_row)
            
            render_html(f"""
            <div class="result-card">
                <div class="result-card-header">
                    <div class="result-name-big">
                        <span>{h_name if h_name else target_row['인명']}</span>
                        <span style="font-family:'Nanum Myeongjo',serif; font-size:18pt; color:#4b5563;">{c_name}</span>
                        <span class="result-badge">{target_row['관계']}</span>
                    </div>
                    <div style="font-size:11pt; color:#6b7280; font-family:monospace;">고유ID: {target_row['고유ID']}</div>
                </div>
                
                <!-- 문단 형식 신상 및 계통 정보 -->
                <div style="font-weight:700; color:#1e3c72; margin-bottom:8px; font-size:12pt;">
                    📌 인물 기본 신상 및 계통 정보 (한글·한자 병기)
                </div>
                <div class="meta-para-container">
                    {meta_paragraphs_html}
                </div>
                
                <!-- ⭐ 제4장 5항 ③: 제목에 kr cn 사용금지 -->
                <div style="font-weight:700; color:#1e3c72; margin-top:16px; margin-bottom:10px; font-size:12pt;">
                    📜 상세 표준 약력 (5단락 체계화 및 고증 보정)
                </div>
                
                <div style="display:grid; grid-template-columns: 1fr 1fr; gap:16px;">
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
            
            # ④ 마지막에 족보 수록 원문 보기 생성
            with st.expander("📖 [족보 수록 원문 보기 (한글 원문 / 한문 원문 대조)]", expanded=False):
                c_won1, c_won2 = st.columns(2)
                with c_won1:
                    render_html("<b>[한글 약력 원문]</b>")
                    st.text_area("한글 원문", target_row.get('한글 약력 원문', ''), height=180, disabled=True, label_visibility="collapsed")
                with c_won2:
                    render_html("<b>[한문 약력 원문 (漢文 原文)]</b>")
                    st.text_area("한문 원문", target_row.get('한문 약력 원문', ''), height=180, disabled=True, label_visibility="collapsed")

    # -------------------------------------------------------------------------
    # B. 계보 검색 연동 (제2장 계보 작성)
    # -------------------------------------------------------------------------
    elif search_type == "계보 검색" and query_val.strip():
        q = query_val.strip()
        q_alt = q[1:] if q.startswith("안") and len(q) > 1 else q
        mask = (
            df['인명'].astype(str).str.contains(q, case=False, na=False) |
            df['인명'].astype(str).str.contains(q_alt, case=False, na=False) |
            df['한글 약력 정리'].apply(lambda x: q in str(x).split('\n')[0] or q_alt in str(x).split('\n')[0])
        )
        l_matches = df[mask]
        if len(l_matches) == 0:
            st.warning("계보를 작성할 선조 인명을 찾을 수 없습니다.")
        else:
            if len(l_matches) > 1:
                cand_opts = [f"{r['인명']} (세수: {extract_generation_info(r['고유ID'])[1]}, 부: {r['부친']}) - {r['고유ID']}" for _, r in l_matches.iterrows()]
                sel_cand = st.selectbox("계보 기준 인물 선택", cand_opts)
                target_gid = sel_cand.split(" - ")[-1].strip()
                tree_target = l_matches[l_matches['고유ID'] == target_gid].iloc[0]
            else:
                tree_target = l_matches.iloc[0]
            
            st.markdown("#### ⚙️ 계보 작성 조건 창")
            c_cond1, c_cond2 = st.columns(2)
            with c_cond1:
                up_gens = st.slider("상향 세대수 (선대 조상)", min_value=1, max_value=5, value=3)
            with c_cond2:
                down_gens = st.slider("하향 세대수 (후손 자손)", min_value=1, max_value=4, value=2)
                
            st.caption("※ 규칙 적용: 본인·부·조부 및 자·손자는 형제 전원 표기, 증조부 이상 및 증손자 이하는 'O자 중 O자' 축소 수평 지그재그 표기")
            
            render_tree_hierarchy(tree_target, df, max_up=up_gens, max_down=down_gens)

    # -------------------------------------------------------------------------
    # C. 단어 검색 처리 (제1장 ③)
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
        
        st.markdown(f"**단어 검색 결과 : '{term}'을(를) 포함하는 약력을 가진 선조 총 {len(word_matches):,}위**")
        
        if len(word_matches) == 0:
            st.info("해당 단어를 포함하고 있는 약력 기록이 없습니다.")
        else:
            candidate_list = []
            for _, r in word_matches.iterrows():
                gen_num, gen_str = extract_generation_info(r['고유ID'])
                candidate_list.append(f"[{gen_str}] {r['인명']} (부: {r['부친']}, 관계: {r['관계']}) - 고유ID: {r['고유ID']}")
                
            chosen = st.selectbox("조회할 선조를 선택하세요 (선택 시 핵심 약력 표시)", candidate_list)
            chosen_gid = chosen.split("고유ID: ")[-1].strip()
            chosen_row = word_matches[word_matches['고유ID'] == chosen_gid].iloc[0]
            
            render_html(f"""
            <div class="result-card">
                <div class="result-card-header">
                    <div class="result-name-big">
                        <span>{chosen_row['인명']}</span>
                        <span class="result-badge">{chosen_row['관계']}</span>
                    </div>
                    <div style="font-size:11pt; color:#6b7280;">고유ID: {chosen_row['고유ID']}</div>
                </div>
                <div class="bio-detail-container">
                    <div class="bio-section-title">💡 핵심 약력 요약</div>
                    {format_structured_bio(chosen_row.get('한글 약력 정리', ''), is_hanmun=False)}
                </div>
            </div>
            """)
            
            with st.expander("📖 [상세 원문 및 한문 약력 대조 보기]"):
                render_html("<b>[한문 약력 정리]</b>")
                st.write(chosen_row.get('한문 약력 정리', ''))
                render_html("<b>[한글 약력 원문]</b>")
                st.write(chosen_row.get('한글 약력 원문', ''))

# =============================================================================
# TAB 2: 계보도 전용 조회 (제2장)
# =============================================================================
with nav_tab2:
    st.markdown("### 🌳 완귀공파 계보도 직접 탐색")
    st.caption("특정 선조를 선택하여 직계 선대 및 후손의 혈통 계보를 계통적으로 열람할 수 있습니다.")
    quick_names = ["嶒증", "宗慶종경", "大海대해", "琠전", "鳴漢명한", "汝履여리", "景時경시"]
    sel_q = st.selectbox("대표 선조 계보 바로보기", quick_names)
    q_match = df[df['인명'].astype(str).str.contains(str(sel_q), na=False)]
    if not q_match.empty:
        render_tree_hierarchy(q_match.iloc[0], df)

# =============================================================================
# TAB 3: 묘제일 메뉴 생성 (제5장 묘제일 안내)
# =============================================================================
with nav_tab3:
    st.markdown("### 🕯️ 광주안씨 완귀공파 묘제일(향사 및 시제) 상세 안내")
    
    render_html("""
    <div class="wangwi-intro-box" style="border-left-color: #8c2d19;">
        <div class="wangwi-intro-title" style="color:#8c2d19;">📅 완귀공 종중 정례 제향 및 시제 일정표</div>
        <div class="meta-para-container" style="background:#ffffff; border:1px solid #cbd5e1; margin-top:10px;">
            <div class="meta-para-item"><b>1. 구암리 상로재(霜露齋) 춘향대제 :</b> 매년 양력 4월 2째주 일요일 (경북 영천시 구암리 상로재, 완귀공 후손 종원 일동)</div>
            <div class="meta-para-item"><b>2. 청제내 선영 묘제 :</b> 매년 춘향대제 봉행 직후 (영천시 금호읍 청제내 화전등, 완귀공 및 배위 영양최씨 묘소)</div>
            <div class="meta-para-item"><b>3. 작산 선영 시향(묘제) :</b> 매년 음력 10월 상정일 (경북 영천시 작산 묘역 안장 각 선조위)</div>
            <div class="meta-para-item"><b>4. 각 지파별 시제 :</b> 음력 10월 중 각 지파별 길일 (효식·효빈·효욱·효갑·효방 등 각 지파 재실 및 선영)</div>
        </div>
    </div>
    
    <div class="bio-detail-container" style="background:#ffffff;">
        <div class="bio-section-title">🏛️ 전통 제례 절차 및 종원 수칙 안내</div>
        <ol style="line-height:2.0; padding-left:20px; font-size:11pt; color:#374151;">
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
    """)

# =============================================================================
# 6. 하단 푸터 (Footer)
# =============================================================================
render_html("""
<div style="text-align:center; padding:25px 0 10px 0; color:#9ca3af; font-size:10pt; border-top:1px solid #e5e7eb; margin-top:35px;">
    광주안씨 완귀공파 종중 회의소 | 데이터 출처: 보정 족보 1~2권 통합본 (총 2,887위) | 연동 파일: genealogy.xlsx
</div>
""")
