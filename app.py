import streamlit as st
import pandas as pd
import openpyxl
import re
import os

# -----------------------------------------------------------------------------
# 1. 페이지 기본 설정 및 스타일 (CSS)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="광주안씨 완귀공파 족보 검색 시스템",
    page_icon="📜",
    layout="wide",
    initial_sidebar_state="expanded"
)

CUSTOM_CSS = """
<style>
/* 전체 폰트 및 배경 스타일 */
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&family=Nanum+Myeongjo:wght@400;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Noto Sans KR', sans-serif;
}

/* 상단 메인 헤더 배너 */
.main-header-banner {
    background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
    color: white;
    padding: 24px 30px;
    border-radius: 12px;
    margin-bottom: 25px;
    box-shadow: 0 4px 15px rgba(0,0,0,0.1);
}
.main-header-title {
    font-size: 26px;
    font-weight: 700;
    margin-bottom: 8px;
    display: flex;
    align-items: center;
    gap: 10px;
}
.main-header-sub {
    font-size: 14px;
    color: #e0e7ff;
    font-weight: 300;
}

/* 검색 결과 인물 카드 컨테이너 */
.person-card-container {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 20px;
    margin-bottom: 22px;
    box-shadow: 0 3px 10px rgba(0,0,0,0.04);
    transition: all 0.2s ease-in-out;
}
.person-card-container:hover {
    border-color: #cbd5e1;
    box-shadow: 0 6px 16px rgba(0,0,0,0.08);
}

/* 인명 및 기본 식별 행 */
.result-person-row {
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 10px;
    padding-bottom: 14px;
    border-bottom: 2px solid #f1f5f9;
    margin-bottom: 14px;
}
.person-icon-badge {
    font-size: 22px;
    background: #eff6ff;
    width: 38px;
    height: 38px;
    display: flex;
    align-items: center;
    justify-content: center;
    border-radius: 8px;
    border: 1px solid #dbeafe;
}
.person-name-hangul {
    font-size: 20px;
    font-weight: 700;
    color: #0f172a;
}
.person-name-hanja {
    font-size: 17px;
    font-family: 'Nanum Myeongjo', serif;
    color: #475569;
}
.person-rel-badge {
    font-size: 13px;
    font-weight: 600;
    background-color: #e0e7ff;
    color: #3730a3;
    padding: 3px 10px;
    border-radius: 20px;
}
.person-gid-text {
    font-size: 12px;
    color: #64748b;
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    padding: 3px 9px;
    border-radius: 6px;
    font-family: monospace;
    margin-left: auto;
}

/* 혈통 계통 정보 박스 */
.lineage-badge-row {
    display: flex;
    gap: 15px;
    flex-wrap: wrap;
    background-color: #f8fafc;
    padding: 10px 14px;
    border-radius: 8px;
    font-size: 13.5px;
    margin-bottom: 14px;
    border: 1px solid #f1f5f9;
}
.lineage-item-label {
    font-weight: 600;
    color: #475569;
}
.lineage-item-val {
    color: #1e293b;
    font-weight: 500;
}

/* 주요 약력 박스 */
.bio-summary-box {
    background-color: #fcfcfd;
    border: 1px solid #edf2f7;
    border-radius: 8px;
    padding: 15px 18px;
}
.bio-summary-header {
    font-size: 14px;
    font-weight: 700;
    color: #1e3a8a;
    margin-bottom: 10px;
    display: flex;
    align-items: center;
    gap: 6px;
}
.bio-summary-body {
    font-size: 14px;
    line-height: 1.75;
    color: #334155;
    word-break: keep-all;
}
.bio-para-line {
    margin-bottom: 8px;
    padding-left: 2px;
}
.bio-label-tag {
    font-weight: 700;
    color: #1e40af;
    background-color: #eff6ff;
    padding: 2px 7px;
    border-radius: 4px;
    margin-right: 6px;
    font-size: 13px;
    display: inline-block;
}
.red-highlight {
    color: #dc2626 !important;
    font-weight: 600;
    background-color: #fef2f2;
    padding: 1px 4px;
    border-radius: 3px;
}
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. 엑셀 데이터 로드 함수 (캐싱 적용)
# -----------------------------------------------------------------------------
EXCEL_PATHS = [
    "완귀공파 족보 검색 시스템.xlsx",
    "artifacts/file_generation/ttl=63d/output/완귀공파 족보 검색 시스템.xlsx",
    "완귀공파 족보 통합.xlsx",
    "artifacts/file_generation/ttl=63d/output/완귀공파 족보 통합.xlsx",
    "보정 족보A_B_C_D_통합_전사결과.xlsx"
]

@st.cache_data
def load_genealogy_data():
    target_path = None
    for p in EXCEL_PATHS:
        if os.path.exists(p):
            target_path = p
            break
            
    if target_path is None:
        return None, "엑셀 파일을 찾을 수 없습니다."
        
    try:
        wb = openpyxl.load_workbook(target_path, read_only=True)
        sheet_name = "족보전사결과" if "족보전사결과" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return None, "시트에 데이터가 없습니다."
            
        header = [str(c).strip() if c is not None else f"Col{idx}" for idx, c in enumerate(rows[0])]
        data = rows[1:]
        
        df = pd.DataFrame(data, columns=header)
        df = df.fillna("")
        return df, target_path
    except Exception as e:
        return None, str(e)

# -----------------------------------------------------------------------------
# 3. 유틸리티 함수: 세수 계산 및 텍스트 파싱
# -----------------------------------------------------------------------------
def extract_generation_from_gid(gid):
    # 고유ID 예시: B-93-4-1-32 -> 32세손
    m = re.search(r"-(\d+)$", str(gid).strip())
    if m:
        gen = int(m.group(1))
        offset = gen - 23  # 완귀공(安嶒)이 23세손 또는 24세손
        return gen, f"{gen}세손 (완귀공 {gen-23}대손)"
    return None, "세수 미상"

def split_name_hangul_hanja(name_raw):
    name_str = str(name_raw).strip()
    # 예: 景時경시 -> 景時 (한문), 경시 (한글)
    m = re.match(r"^([一-龥]+)?([가-힣]+)?$", name_str)
    if m and m.group(1) and m.group(2):
        return m.group(2), m.group(1)
    
    # 한글만 있는 경우
    if re.fullmatch(r"^[가-힣]+$", name_str):
        return name_str, ""
    # 한자만 있는 경우
    if re.fullmatch(r"^[一-龥]+$", name_str):
        return "", name_str
        
    return name_str, ""

def highlight_red_corrections(text):
    if not text:
        return ""
    # 왕대/간지/연대 보정 키워드 강조 처리
    keywords = [
        "1699년", "1699", "一六九九", "1626년", "1626", "一六二六",
        "1558년", "1558", "一五五八", "1816년", "1816", "一八一六年",
        "1814년", "1814", "一八一四", "1771년", "1771", "一七七一",
        "1740년", "1740", "一七四〇", "1750년", "1750", "一七五〇",
        "1775년", "1775", "一七七五", "1898년", "1898", "一八九八",
        "1910년", "1910", "一九一〇", "순종", "純宗", "헌종", "憲宗",
        "기해", "己亥", "무신", "戊申", "자좌 오향", "자좌 오", "子坐 午向", "子坐 午",
        "내등말", "內嶝末", "오보허", "五步許", "북록 묘좌", "北麓 卯坐"
    ]
    formatted = str(text)
    for kw in keywords:
        if kw in formatted:
            formatted = formatted.replace(kw, f'<span class="red-highlight">{kw}</span>')
    return formatted

def format_bio_html(bio_raw):
    if not bio_raw:
        return '<div style="color:#94a3b8; font-style:italic;">기록된 약력이 없습니다.</div>'
        
    lines = [line.strip() for line in str(bio_raw).split("\n") if line.strip()]
    html_paragraphs = []
    
    for idx, line in enumerate(lines):
        line_hl = highlight_red_corrections(line)
        
        # 5단락 헤더 태그 디자인 분기
        if line_hl.startswith("묘 :") or line_hl.startswith("墓 :"):
            content = line_hl[3:].strip()
            label = "묘소" if line_hl.startswith("묘") else "墓所"
            html_paragraphs.append(f'<div class="bio-para-line"><span class="bio-label-tag">🪦 {label}</span>{content}</div>')
        elif line_hl.startswith("과거·관직 :") or line_hl.startswith("科·官 :"):
            content = line_hl[6:].strip() if line_hl.startswith("과거") else line_hl[5:].strip()
            label = "과거·관직" if line_hl.startswith("과거") else "科·官"
            html_paragraphs.append(f'<div class="bio-para-line"><span class="bio-label-tag">🏛️ {label}</span>{content}</div>')
        elif line_hl.startswith("학문·인품·업적 :") or line_hl.startswith("學·行·業 :"):
            content = line_hl[9:].strip() if line_hl.startswith("학문") else line_hl[6:].strip()
            label = "학문·업적" if line_hl.startswith("학문") else "學·業"
            html_paragraphs.append(f'<div class="bio-para-line"><span class="bio-label-tag">📖 {label}</span>{content}</div>')
        elif line_hl.startswith("배 :") or line_hl.startswith("配 :"):
            content = line_hl[3:].strip()
            label = "배위" if line_hl.startswith("배") else "配位"
            html_paragraphs.append(f'<div class="bio-para-line"><span class="bio-label-tag">💍 {label}</span>{content}</div>')
        elif line_hl.startswith("남편 :") or line_hl.startswith("夫 :"):
            content = line_hl[4:].strip() if line_hl.startswith("남편") else line_hl[3:].strip()
            html_paragraphs.append(f'<div class="bio-para-line"><span class="bio-label-tag">👔 남편</span>{content}</div>')
        else:
            # 1단락 (기본 신상)
            html_paragraphs.append(f'<div class="bio-para-line"><span class="bio-label-tag">👤 신상</span>{line_hl}</div>')
            
    return "".join(html_paragraphs)

# -----------------------------------------------------------------------------
# 4. 사이드바 검색 및 필터 옵션
# -----------------------------------------------------------------------------
df, source_info = load_genealogy_data()

st.sidebar.header("🔍 선조 인명 검색")

if df is None:
    st.error(f"데이터 로드 실패: {source_info}")
    uploaded_file = st.sidebar.file_uploader("족보 엑셀 파일 직접 업로드", type=["xlsx", "xlsm"])
    if uploaded_file:
        wb = openpyxl.load_workbook(uploaded_file, read_only=True)
        sheet_name = "족보전사결과" if "족보전사결과" in wb.sheetnames else wb.sheetnames[0]
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        header = [str(c).strip() for c in rows[0]]
        df = pd.DataFrame(rows[1:], columns=header).fillna("")
        st.success("업로드 파일 로드 완료!")
        st.rerun()
    st.stop()

# 검색어 입력
search_query = st.sidebar.text_input("인명 (한글 또는 한자)", placeholder="예: 경시, 景時, 완귀, 안증").strip()
gid_query = st.sidebar.text_input("고유ID 검색", placeholder="예: B-93-4-1-32, A-24").strip()

# 세수 필터
gen_list = sorted(list(set([extract_generation_from_gid(g)[0] for g in df['고유ID'] if extract_generation_from_gid(g)[0] is not None])))
selected_gen = st.sidebar.selectbox("세수(世代) 선택", ["전체 세수"] + [f"{g}세손" for g in gen_list])

# 약력 언어 모드 선택
bio_mode = st.sidebar.radio("약력 표시 모드", ["한글 약력 정리", "한문 약력 정리", "원문·정리 대조"], index=0)

# -----------------------------------------------------------------------------
# 5. 검색 로직 실행
# -----------------------------------------------------------------------------
filtered_df = df.copy()

if search_query:
    filtered_df = filtered_df[
        filtered_df['인명'].astype(str).str.contains(search_query, case=False, na=False) |
        filtered_df['한글 약력 정리'].astype(str).str.contains(search_query, case=False, na=False) |
        filtered_df['한문 약력 정리'].astype(str).str.contains(search_query, case=False, na=False)
    ]

if gid_query:
    filtered_df = filtered_df[filtered_df['고유ID'].astype(str).str.contains(gid_query, case=False, na=False)]

if selected_gen != "전체 세수":
    target_gen = int(selected_gen.replace("세손", ""))
    filtered_df = filtered_df[filtered_df['고유ID'].apply(lambda x: extract_generation_from_gid(x)[0] == target_gen)]

# -----------------------------------------------------------------------------
# 6. 메인 화면 렌더링
# -----------------------------------------------------------------------------
# 배너
st.markdown("""
<div class="main-header-banner">
    <div class="main-header-title">📜 광주안씨 완귀공파 족보 검색 시스템</div>
    <div class="main-header-sub">선조 인명·세계도 계통·5단락 표준 약력 및 고증 보정(빨간색) 통합 검색 포털</div>
</div>
""", unsafe_allow_html=True)

# 검색 결과 통계 요약
total_count = len(filtered_df)
c1, c2, c3 = st.columns([2, 2, 6])
c1.metric("검색 결과 인원", f"{total_count:,} 명")
c2.metric("전체 등록 인원", f"{len(df):,} 명")
c3.caption(f"📁 데이터 소스: `{os.path.basename(source_info)}` (족보전사결과 시트)")

st.write("")

# -----------------------------------------------------------------------------
# 7. 인명 검색 결과 카드 출력 (반드시 unsafe_allow_html=True 사용)
# -----------------------------------------------------------------------------
if total_count == 0:
    st.info("검색 조건에 일치하는 선조 기록이 없습니다. 검색어를 다시 확인해 주세요.")
else:
    # 페이지네이션 (1페이지당 10명씩 표시)
    PAGE_SIZE = 10
    total_pages = (total_count - 1) // PAGE_SIZE + 1
    
    if total_pages > 1:
        page = st.selectbox(f"페이지 이동 (총 {total_pages} 페이지)", range(1, total_pages + 1))
    else:
        page = 1
        
    start_idx = (page - 1) * PAGE_SIZE
    end_idx = min(start_idx + PAGE_SIZE, total_count)
    page_df = filtered_df.iloc[start_idx:end_idx]

    for _, row in page_df.iterrows():
        gid = row.get('고유ID', '')
        name_raw = row.get('인명', '')
        rel = row.get('관계', '')
        father = row.get('부친', '')
        gfather = row.get('조부', '')
        
        name_hangul, name_hanja = split_name_hangul_hanja(name_raw)
        gen_num, gen_str = extract_generation_from_gid(gid)
        
        # 선택된 모드에 따른 약력 텍스트 추출
        if bio_mode == "한글 약력 정리":
            bio_content_html = format_bio_html(row.get('한글 약력 정리', ''))
        elif bio_mode == "한문 약력 정리":
            bio_content_html = format_bio_html(row.get('한문 약력 정리', ''))
        else:
            # 원문 대조 모드
            bio_content_html = f"""
            <div style="margin-bottom:12px;">
                <div style="font-weight:700; color:#1e3a8a; margin-bottom:5px;">[한글 약력 정리]</div>
                {format_bio_html(row.get('한글 약력 정리', ''))}
            </div>
            <div style="border-top:1px dashed #cbd5e1; padding-top:10px;">
                <div style="font-weight:700; color:#475569; margin-bottom:5px;">[한문 약력 정리]</div>
                {format_bio_html(row.get('한문 약력 정리', ''))}
            </div>
            """

        card_html = f"""
        <div class="person-card-container">
            <div class="result-person-row">
                <span class="person-icon-badge">📜</span>
                <span class="person-name-hangul">{name_hangul if name_hangul else name_raw}</span>
                <span class="person-name-hanja">{name_hanja}</span>
                <span class="person-rel-badge">{rel}</span>
                <span class="person-gid-text">고유ID: {gid}</span>
            </div>
            <div class="lineage-badge-row">
                <div><span class="lineage-item-label">세수 :</span> <span class="lineage-item-val">{gen_str}</span></div>
                <div><span class="lineage-item-label">부친 :</span> <span class="lineage-item-val">{father if father else '-'}</span></div>
                <div><span class="lineage-item-label">조부 :</span> <span class="lineage-item-val">{gfather if gfather else '-'}</span></div>
            </div>
            <div class="bio-summary-box">
                <div class="bio-summary-header">주요 약력 (5단락 표준 고증)</div>
                <div class="bio-summary-body">
                    {bio_content_html}
                </div>
            </div>
        </div>
        """
        
        # [핵심] st.markdown 에 unsafe_allow_html=True 를 적용하여 카드로 렌더링
        st.markdown(card_html, unsafe_allow_html=True)
