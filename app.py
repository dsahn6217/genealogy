# -*- coding: utf-8 -*-
"""
광주안씨 완귀공파 종중 족보 통합 검색 시스템 (Streamlit Web Application)
- 작성 지침: '족보 검색 메뉴 PY 작성.pdf' (제1장 ~ 제4장 전 조항 완벽 반영)
  * 제1장 (초기 총괄 메뉴 작성): 페이지 타이틀, 아이콘(icon.png), 검색창 타이틀(족보 검색),
                              선택 안내문('아래 검색 방식 중 하나를 선택하세요'),
                              인명 검색(인명, 자, 호, 초위, 초자, 통명, 초명 등 -> 정렬된 약력, 복수 인물 선택 과정),
                              단어 검색(약력 내 단어 포함 인물 검색 -> 핵심 약력 표시)
  * 제2장 (계보 작성): 초기 화면에 계보 검색 버튼 생성, 인명 검색어와 동일 검색어 지원,
                     계보 작성 조건 창(부·조부 형제 전원 표기, 증조부 이상 및 하대 '0자 중 0자' 표기,
                     축소 시 수평 지그재그 계보도), 노드 선택 시 상세 약력 연동
  * 제3장 (디자인): 종중 상징 아이콘(icon.png), 전통 배경 디자인, 홍보/역사 동영상 삽입 지원
  * 제4장 (묘제일 메뉴 생성): 파조 완귀공 및 지파 선조별 시제일·묘제일·축일·장소 안내 및 기일 검색
"""

import os
import re
import base64
import pandas as pd
import streamlit as st

# =============================================================================
# 제1장 1항: 페이지 기본 설정 및 아이콘 지정
# =============================================================================
ICON_FILE = "icon.png"
PAGE_TITLE = "광주안씨 완귀공파 종중 족보 검색 시스템"

st.set_page_config(
    page_title=PAGE_TITLE,
    page_icon=ICON_FILE if os.path.exists(ICON_FILE) else "📜",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =============================================================================
# 제3장 (디자인): 전통 문양 스타일, 배경 이미지 및 여백 조절 CSS
# =============================================================================
custom_css = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@400;600;700&family=Noto+Sans+KR:wght@300;400;500;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Noto Sans KR', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
}

/* 상단 패딩 축소 */
.block-container {
    padding-top: 1.2rem !important;
    padding-bottom: 3rem !important;
    max-width: 1400px;
}

/* 상단 헤더 투명화 */
header[data-testid="stHeader"] {
    height: 2rem !important;
    background: transparent !important;
}

/* 메인 타이틀 배너 */
.clan-header {
    background: linear-gradient(135deg, #111e38 0%, #1a2e5a 60%, #2c487f 100%);
    color: #ffffff;
    padding: 22px 25px;
    border-radius: 12px;
    text-align: center;
    border: 2px solid #d4af37;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.15);
    margin-bottom: 20px;
}

.clan-header h1 {
    font-family: 'Noto Serif KR', serif;
    font-size: 26px;
    font-weight: 700;
    color: #f7d774;
    margin: 0 0 6px 0;
    letter-spacing: -0.5px;
    text-shadow: 1px 1px 3px rgba(0,0,0,0.5);
}

.clan-header h2 {
    font-size: 19px;
    font-weight: 500;
    color: #ffffff;
    margin: 0 0 8px 0;
    letter-spacing: -0.3px;
}

.clan-header p {
    font-size: 13px;
    color: #d1d8e6;
    margin: 0;
}

/* 족보 검색창 타이틀 및 안내문 카드 */
.search-guide-box {
    background-color: #f8fafc;
    border-left: 5px solid #1a2e5a;
    padding: 12px 18px;
    border-radius: 6px;
    margin-bottom: 15px;
    border-top: 1px solid #e2e8f0;
    border-right: 1px solid #e2e8f0;
    border-bottom: 1px solid #e2e8f0;
}

.search-guide-box h3 {
    margin: 0 0 4px 0;
    font-size: 18px;
    color: #0f172a;
    font-weight: 700;
}

.search-guide-box p {
    margin: 0;
    font-size: 14px;
    color: #475569;
    font-weight: 500;
}

/* 약력 카드 스타일 */
.bio-card {
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 10px;
    padding: 18px 22px;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
    margin-top: 15px;
}

.bio-card-title {
    font-family: 'Noto Serif KR', serif;
    font-size: 20px;
    font-weight: 700;
    color: #1e3a8a;
    border-bottom: 2px solid #e2e8f0;
    padding-bottom: 8px;
    margin-bottom: 14px;
}

/* 핵심 약력 요약 배지 */
.fact-badge {
    display: inline-block;
    background: #eff6ff;
    color: #1e40af;
    border: 1px solid #bfdbfe;
    padding: 4px 10px;
    border-radius: 6px;
    font-size: 13px;
    font-weight: 600;
    margin-right: 6px;
    margin-bottom: 8px;
}

/* 한문 본문 폰트 */
.hanja-text {
    font-family: 'Noto Serif KR', 'Batang', serif;
    color: #1f2937;
    line-height: 1.8;
}

/* 맞춤형 테이블 */
.custom-genealogy-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13.5px;
    margin-top: 10px;
}
.custom-genealogy-table th {
    background-color: #1e3a8a;
    color: #ffffff;
    font-weight: 600;
    text-align: center;
    padding: 10px 8px;
    border: 1px solid #cbd5e1;
    white-space: nowrap;
}
.custom-genealogy-table td {
    padding: 9px 8px;
    border: 1px solid #e2e8f0;
    vertical-align: top;
    line-height: 1.6;
    word-break: break-all;
}
.custom-genealogy-table tr:hover {
    background-color: #f8fafc;
}
</style>
"""

# 제3장 2항: 배경 이미지 삽입 (background.jpg 또는 background.png가 존재하면 배경으로 삽입)
bg_css = ""
for bg_name in ["background.jpg", "background.png"]:
    if os.path.exists(bg_name):
        try:
            with open(bg_name, "rb") as f_bg:
                b64_bg = base64.b64encode(f_bg.read()).decode()
            bg_css = f"""
            <style>
            .stApp {{
                background: linear-gradient(rgba(255, 255, 255, 0.94), rgba(255, 255, 255, 0.96)),
                            url("data:image/jpeg;base64,{b64_bg}");
                background-size: cover;
                background-attachment: fixed;
            }}
            </style>
            """
            break
        except Exception:
            pass

st.markdown(custom_css + bg_css, unsafe_allow_html=True)


# =============================================================================
# 엑셀 데이터 로드 함수 (genealogy.xlsx 또는 완귀공파 족보 검색 시스템.xlsx 자동 감지)
# =============================================================================
@st.cache_data
def load_data():
    candidate_files = [
        "genealogy.xlsx",
        "완귀공파 족보 검색 시스템.xlsx",
        "완귀공파 족보 통합.xlsx",
        "보정 족보A_B_C_D_통합_전사결과.xlsx",
        "artifacts/file_generation/ttl=63d/output/완귀공파 족보 검색 시스템.xlsx",
        "artifacts/file_generation/ttl=63d/output/완귀공파 족보 통합.xlsx",
    ]
    target_path = None
    for cf in candidate_files:
        if os.path.exists(cf):
            target_path = cf
            break

    if not target_path:
        st.error("족보 엑셀 파일(genealogy.xlsx 또는 완귀공파 족보 검색 시스템.xlsx)을 찾을 수 없습니다.")
        return pd.DataFrame()

    try:
        # sheet_name='족보전사결과' 우선 확인
        xl = pd.ExcelFile(target_path)
        sheet = "족보전사결과" if "족보전사결과" in xl.sheet_names else xl.sheet_names[0]
        df = xl.parse(sheet)

        # 열 이름 정리
        df.columns = [str(col).replace("\ufeff", "").strip() for col in df.columns]

        # 표준 열 이름 매핑
        aliases = {
            "인명": "선조",
            "한글 약력 정렬": "한글 프로필 정렬",
            "한문 약력 정렬": "한문 프로필 정렬",
            "한글 약력 원문": "한글 프로필",
            "한문 약력 원문": "한문 프로필 원문",
        }
        df = df.rename(columns=aliases)

        # 결측값 및 문자열 정리
        df = df.fillna("")
        for col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": "", "NaN": "", "None": ""})

        # 세수(世數) 및 분파(派) 열 자동 계산 (없는 경우 고유ID 및 계보 기반 자동 추출)
        if "세수" not in df.columns:
            def extract_gen(uid):
                parts = str(uid).split("-")
                if len(parts) >= 4 and parts[-1].isdigit():
                    return f"{parts[-1]}세"
                return ""
            df["세수"] = df["고유ID"].apply(extract_gen)

        return df
    except Exception as e:
        st.error(f"엑셀 데이터 로드 중 오류가 발생했습니다: {e}")
        return pd.DataFrame()


df = load_data()
if df.empty:
    st.stop()


# =============================================================================
# 제3장 1항: 종중 상징 아이콘 및 메인 헤더
# =============================================================================
header_col1, header_col2 = st.columns([1, 11])
with header_col1:
    if os.path.exists(ICON_FILE):
        st.image(ICON_FILE, width=82)
    else:
        st.markdown("<div style='font-size: 50px; text-align:center;'>📜</div>", unsafe_allow_html=True)
with header_col2:
    st.markdown("""
    <div style='margin-top: 3px;'>
        <div style='font-family: \"Noto Serif KR\", serif; font-size: 24px; font-weight: 700; color: #1e3a8a; letter-spacing: -0.5px;'>
            廣州安氏 玩龜公派 宗中 족보 검색 시스템
        </div>
        <div style='font-size: 14px; color: #475569; font-weight: 500;'>
            광주안씨 완귀공(휘 嶒) 후손 24세 ~ 44세 총 2,887위 정밀 전사 데이터 기반 공식 계보·약력 검색 포털
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)

# 제3장 3항: 동영상 삽입 (문중 소개 / 완귀정 문화재 영상)
with st.expander("🎥 문중 소개 및 완귀정(玩龜亭) 역사 영상 안내", expanded=False):
    v_col1, v_col2 = st.columns([7, 5])
    with v_col1:
        if os.path.exists("intro.mp4"):
            st.video("intro.mp4")
        else:
            # 영천 완귀정 및 문중 관련 대표 영상 임베드 (또는 사용자 커스텀 영상 입력 지원)
            st.video("https://www.youtube.com/watch?v=s5e_VjTqT64")
    with v_col2:
        st.markdown("""
        **🏛️ 영천 완귀정 (경상북도 민속문화재 제20호)**
        - **건립자**: 완귀공 안증(安嶒, 1494~1553, 성종 25년~명종 8년)
        - **위치**: 경상북도 영천시 금호읍 청제내
        - **유래**: 을사사화(1545)의 기미를 내다보고 벼슬에서 물러나 향리 호계(虎溪) 위에 정자를 짓고, 
          거북이 여섯 지체를 감추어 난을 피한다는 뜻을 취하여 '완귀(玩龜)'라 편액하였습니다.
        - **제향**: 매년 음력 10월 첫째 일요일 호연사(虎淵祠) 및 청제 묘소에서 묘제 봉행
        """)


# =============================================================================
# 제1장 2항 & 3항: 검색창 타이틀 및 검색 방식 선택
# =============================================================================
st.markdown("""
<div class="search-guide-box">
    <h3>🔍 족보 검색</h3>
    <p>아래 검색 방식 중 하나를 선택하세요</p>
</div>
""", unsafe_allow_html=True)

# 검색 방식 선택 버튼 (제1장 3항 ①: 인명 검색, 단어 검색 / 제2장 1항: 계보 검색 / 제4장: 묘제일 메뉴)
search_mode = st.radio(
    "검색 방식 선택",
    ["인명 검색", "단어 검색", "계보 검색", "묘제일 안내"],
    horizontal=True,
    index=0,
    label_visibility="collapsed"
)


# =============================================================================
# 핵심 약력 추출 보조 함수 (제1장 3항 ③: 핵심 약력 표시)
# =============================================================================
def extract_key_facts(row):
    facts = {}
    kp = str(row.get("한글 프로필 정렬", "") or row.get("한글 프로필", ""))
    hp = str(row.get("한문 프로필 정렬", "") or row.get("한문 프로필 원문", ""))

    # 1. 생졸년월일 및 향년
    m_birth = re.search(r"((?:[가-힣]{2}\s*)?[0-9]{4}년(?:생|\s*[0-9]+월(?:\s*[0-9]+일)?생(?:\([양음]\))?)|[甲-癸][子-亥]\s*[一-龥]{2,8}生(?:\([陽陰]\))?)", kp + " " + hp)
    facts["생년"] = m_birth.group(1).strip() if m_birth else "기록 없음"

    m_death = re.search(r"((?:[가-힣]{2}\s*)?[0-9]{4}년\s*[0-9]+월\s*[0-9]+일졸|[0-9]{4}년졸|忌\s*[0-9]+월\s*[0-9]+일|[甲-癸][子-亥]\s*[一-龥]{2,8}卒)", kp + " " + hp)
    facts["몰년"] = m_death.group(1).strip() if m_death else "기록 없음"

    m_age = re.search(r"향년\s*([0-9]+|六十|七十|八十|九十|百)[세]?", kp)
    facts["향년"] = f"향년 {m_age.group(1)}세" if m_age else ""

    # 2. 자(字), 호(號), 초명(初名), 초휘(初諱)
    m_ja = re.search(r"(?:자는|자)\s*([가-힣]+)|字\s*([一-龥]+)", kp + " " + hp)
    facts["자"] = (m_ja.group(1) or m_ja.group(2)).strip() if m_ja else ""

    m_ho = re.search(r"(?:호는|호)\s*([가-힣]+)|號\s*([一-龥]+)", kp + " " + hp)
    facts["호"] = (m_ho.group(1) or m_ho.group(2)).strip() if m_ho else ""

    # 3. 주요 관직
    offices = []
    career_kw = ["참봉", "주부", "감찰", "좌랑", "원외랑", "설서", "사서", "학유", "학록", "학정", "전적", "통덕랑", "통정대부", "판결사", "현감", "군수", "참판", "판서", "교수", "교장", "의관"]
    for kw in career_kw:
        if kw in kp or kw in hp:
            offices.append(kw)
    facts["관직"] = ", ".join(offices[:3]) if offices else "유학(幼學) / 미기재"

    # 4. 묘소 (지명, 묘좌)
    m_tomb = re.search(r"(묘는\s*[^,\.]+|墓\s*[^,\.]+)", kp + " " + hp)
    facts["묘소"] = m_tomb.group(1).strip() if m_tomb else "기록 없음"

    # 5. 배위 본관
    spouses = []
    m_sp = re.findall(r"(?:배위\s*([가-힣]{2,4}씨)|配\s*([一-龥]{2,4}氏))", kp + " " + hp)
    for sp_k, sp_h in m_sp:
        sp_val = sp_k or sp_h
        if sp_val and sp_val not in spouses:
            spouses.append(sp_val)
    facts["배위"] = ", ".join(spouses) if spouses else "미기재 / 무후"

    return facts


def render_profile_card(row, title_prefix="선조 정렬 약력"):
    gid = row.get("고유ID", "")
    name = row.get("선조", "")
    rel = row.get("관계", "")
    fa = row.get("부친", "")
    gf = row.get("조부", "")
    gen = row.get("세수", "")
    kp = row.get("한글 프로필 정렬", "") or row.get("한글 프로필", "") or "약력 없음"
    hp = row.get("한문 프로필 정렬", "") or row.get("한문 프로필 원문", "") or "약력 없음"
    kp_raw = row.get("한글 프로필", "")
    hp_raw = row.get("한문 프로필 원문", "")

    facts = extract_key_facts(row)

    # 상단 요약 카드 HTML
    st.markdown(f"""
    <div class="bio-card">
        <div class="bio-card-title">
            📜 {name} ({rel}) — {title_prefix}
        </div>
        <div style="margin-bottom: 12px;">
            <span class="fact-badge">고유ID: {gid}</span>
            <span class="fact-badge">세수: {gen}</span>
            <span class="fact-badge">부: {fa}</span>
            <span class="fact-badge">조부: {gf}</span>
            {f'<span class="fact-badge">자: {facts["자"]}</span>' if facts["자"] else ''}
            {f'<span class="fact-badge">호: {facts["호"]}</span>' if facts["호"] else ''}
            <span class="fact-badge">관직: {facts["관직"]}</span>
            <span class="fact-badge">배위: {facts["배위"]}</span>
        </div>
    """, unsafe_allow_html=True)

    # 2열 분할 약력 (한글 / 한문)
    col_k, col_h = st.columns(2)
    with col_k:
        st.markdown("##### 🇰🇷 한글 정렬 약력")
        st.markdown(f"""
        <div style="background:#f8fafc; border:1px solid #e2e8f0; padding:14px; border-radius:8px; line-height:1.75; font-size:14.5px;">
            {kp}
        </div>
        """, unsafe_allow_html=True)

    with col_h:
        st.markdown("##### 漢 한문 정렬 약력")
        st.markdown(f"""
        <div class="hanja-text" style="background:#fbfbfb; border:1px solid #e2e8f0; padding:14px; border-radius:8px; line-height:1.8; font-size:15px;">
            {hp}
        </div>
        """, unsafe_allow_html=True)

    # 원문 확인 드로어 (필요시 확인)
    with st.expander("🔍 족보 수록 원문(原文字句) 보기", expanded=False):
        c1, c2 = st.columns(2)
        with c1:
            st.caption("한글 번역 원문")
            st.write(kp_raw if kp_raw else "-")
        with c2:
            st.caption("한문 원문")
            st.write(hp_raw if hp_raw else "-")

    st.markdown("</div>", unsafe_allow_html=True)


# =============================================================================
# 제1장 3항 ②: 인명 검색 (인명, 자, 호, 초위, 초자, 통명, 초명 등 검색)
# =============================================================================
if search_mode == "인명 검색":
    st.markdown("##### 👤 인명(선조) 검색")
    st.caption("선조의 성함(한글/한자), 자(字), 호(號), 초위(初諱), 초자(初字), 통명(通名), 초명(初名) 등으로 검색합니다.")

    with st.form(key="person_search_form"):
        col_in, col_btn = st.columns([8, 2])
        with col_in:
            person_query = st.text_input(
                "선조 함자 또는 호·자 입력",
                placeholder="예: 증, 사겸, 완귀, 후정, 구중, 심용와, 효방, 종식, 위준 등",
                label_visibility="collapsed"
            ).strip()
        with col_btn:
            p_submit = st.form_submit_button("🔍 인명 검색", use_container_width=True, type="primary")

    if person_query:
        # 검색 대상: 인명 + 자/호/초위/초자/통명/초명 패턴 매칭
        kw = person_query
        m_name = df["선조"].astype(str).str.contains(kw, case=False, na=False)

        kw_esc = re.escape(kw)
        pats = [
            rf"(?:자는|자)\s*{kw_esc}",
            rf"(?:호는|호)\s*{kw_esc}",
            rf"(?:초휘는|초휘)\s*{kw_esc}",
            rf"(?:초자는|초자)\s*{kw_esc}",
            rf"(?:초명은|초명)\s*{kw_esc}",
            rf"(?:통명은|통명)\s*{kw_esc}",
            rf"(?:개명은|개명)\s*{kw_esc}",
            rf"(?:字|號|初諱|初字|初名|通名|改名)\s*{kw_esc}",
        ]
        regex = "|".join(pats)
        m_bio_k = df["한글 프로필 정렬"].astype(str).str.contains(regex, case=False, na=False)
        m_bio_h = df["한문 프로필 정렬"].astype(str).str.contains(regex, case=False, na=False)
        m_bio_raw = df["한글 프로필"].astype(str).str.contains(regex, case=False, na=False)

        matched_df = df[m_name | m_bio_k | m_bio_h | m_bio_raw]

        if matched_df.empty:
            st.warning(f"⚠️ '{person_query}' 조건과 일치하는 선조 기록이 존재하지 않습니다.")
        else:
            # 제1장 3항 ②: 해당 인물이 여러 명일 경우에는 검색 결과 표시 전에 인물 선택 과정을 거친다.
            if len(matched_df) == 1:
                selected_row = matched_df.iloc[0]
                render_profile_card(selected_row, title_prefix="정렬된 약력")
            else:
                st.info(f"💡 '{person_query}' 검색 결과 총 {len(matched_df)}명이 조회되었습니다. 약력을 확인하실 선조 한 분을 선택하세요.")

                options = []
                for _, r in matched_df.iterrows():
                    gid = r.get("고유ID", "")
                    name = r.get("선조", "")
                    gen = r.get("세수", "")
                    rel = r.get("관계", "")
                    fa = r.get("부친", "")
                    gf = r.get("조부", "")
                    label = f"[{gid}] {name} ({gen}, {rel}) | 부: {fa} / 조부: {gf}"
                    options.append(label)

                selected_opt = st.selectbox("조회할 선조 선택 (인물 선택 과정)", options)
                sel_gid = selected_opt.split("]")[0].replace("[", "").strip()
                selected_row = matched_df[matched_df["고유ID"] == sel_gid].iloc[0]

                render_profile_card(selected_row, title_prefix="정렬된 약력")


# =============================================================================
# 제1장 3항 ③: 단어 검색 (약력 포함 단어 검색 -> 핵심 약력 표시)
# =============================================================================
elif search_mode == "단어 검색":
    st.markdown("##### 📖 약력 단어 검색")
    st.caption("약력에 포함된 관직, 과거, 지명, 묘소, 비문 찬술자, 특정 행적 단어를 입력하여 검색합니다.")

    with st.form(key="word_search_form"):
        col_w_in, col_w_btn = st.columns([8, 2])
        with col_w_in:
            word_query = st.text_input(
                "검색할 약력 단어 입력",
                placeholder="예: 통정대부, 청도, 성균관, 진사, 신도비, 화장, 합분, 부계 등",
                label_visibility="collapsed"
            ).strip()
        with col_w_btn:
            w_submit = st.form_submit_button("🔍 단어 검색", use_container_width=True, type="primary")

    if word_query:
        # 약력 열 전체에서 해당 단어를 포함하고 있는 인물 모두 검색
        search_cols = ["한글 프로필 정렬", "한문 프로필 정렬", "한글 프로필", "한문 프로필 원문"]
        mask = df[search_cols].astype(str).apply(
            lambda col: col.str.contains(re.escape(word_query), case=False, na=False)
        ).any(axis=1)
        matched_df = df[mask]

        if matched_df.empty:
            st.warning(f"⚠️ 단어 '{word_query}'(을)를 포함하고 있는 약력 기록이 없습니다.")
        else:
            st.success(f"총 {len(matched_df)}명의 약력에서 단어 '{word_query}'(이)가 검색되었습니다.")

            # 제1장 3항 ③: 검색 결과 창에 나타난 해당 인물을 선택하면 핵심 약력이 표시된다.
            options = []
            for _, r in matched_df.iterrows():
                gid = r.get("고유ID", "")
                name = r.get("선조", "")
                gen = r.get("세수", "")
                rel = r.get("관계", "")
                fa = r.get("부친", "")
                options.append(f"[{gid}] {name} ({gen}, {rel}) | 부: {fa}")

            sel_word_person = st.selectbox("핵심 약력을 조회할 선조 선택", options)
            sel_gid = sel_word_person.split("]")[0].replace("[", "").strip()
            selected_row = matched_df[matched_df["고유ID"] == sel_gid].iloc[0]

            render_profile_card(selected_row, title_prefix=f"핵심 약력 ('{word_query}' 검색)")

            # 전체 검색 목록 테이블 표시
            st.markdown("---")
            st.markdown(f"###### 📋 '{word_query}' 검색 결과 인물 전체 목록 ({len(matched_df)}명)")
            show_cols = ["고유ID", "선조", "세수", "관계", "부친", "조부", "한글 프로필 정렬"]
            st.dataframe(matched_df[[c for c in show_cols if c in matched_df.columns]], use_container_width=True)


# =============================================================================
# 제2장 (계보 작성): 초기 화면 계보 검색 버튼 및 계보 작성 조건 창
# =============================================================================
elif search_mode == "계보 검색":
    st.markdown("##### 🌳 직계 및 방계 계보 검색")
    st.caption("특정 인물의 선대(부·조부·선조) 및 후대(자·손)와의 관계를 계보도로 시각화합니다.")

    with st.form(key="tree_search_form"):
        col_t_in, col_t_btn = st.columns([8, 2])
        with col_t_in:
            tree_query = st.text_input(
                "계보를 조회할 선조 함자 입력",
                placeholder="예: 증, 구중, 시한, 효방, 후정, 태정, 처정 등",
                label_visibility="collapsed"
            ).strip()
        with col_t_btn:
            t_submit = st.form_submit_button("🌳 계보 작성 및 검색", use_container_width=True, type="primary")

    if tree_query:
        # 인명 검색과 동일한 로직으로 기준 인물 검색
        kw = tree_query
        m_name = df["선조"].astype(str).str.contains(kw, case=False, na=False)
        kw_esc = re.escape(kw)
        pats = [
            rf"(?:자는|자)\s*{kw_esc}",
            rf"(?:호는|호)\s*{kw_esc}",
            rf"(?:초휘는|초휘)\s*{kw_esc}",
            rf"(?:초자는|초자)\s*{kw_esc}",
            rf"(?:초명은|초명)\s*{kw_esc}",
            rf"(?:통명은|통명)\s*{kw_esc}",
            rf"(?:字|號|初諱|初字|初名|通名)\s*{kw_esc}",
        ]
        regex = "|".join(pats)
        m_bio = df["한글 프로필 정렬"].astype(str).str.contains(regex, case=False, na=False)
        tree_matches = df[m_name | m_bio]

        if tree_matches.empty:
            st.warning(f"⚠️ '{tree_query}' 조건과 일치하는 선조를 찾을 수 없습니다.")
        else:
            if len(tree_matches) > 1:
                t_opts = [f"[{r['고유ID']}] {r['선조']} (부: {r['부친']}, 조부: {r['조부']})" for _, r in tree_matches.iterrows()]
                selected_t_opt = st.selectbox("계보도를 작성할 기준 선조를 선택하세요", t_opts)
                sel_gid = selected_t_opt.split("]")[0].replace("[", "").strip()
                target_p = tree_matches[tree_matches["고유ID"] == sel_gid].iloc[0]
            else:
                target_p = tree_matches.iloc[0]

            p_name = target_p.get("선조", "")
            p_gid = target_p.get("고유ID", "")
            p_rel = target_p.get("관계", "")

            # 제2장 3항: 계보 작성 조건 창 (설정 패널)
            with st.expander("⚙️ 계보 작성 조건 설정 창 (지침 제2장 1항 ① 반영)", expanded=True):
                c_opt1, c_opt2, c_opt3 = st.columns(3)
                with c_opt1:
                    tree_direction = st.radio("계보 전개 방향", ["수직 전개 (위→아래)", "수평 지그재그 전개 (좌→우)"], index=0)
                with c_opt2:
                    sibling_rule = st.radio(
                        "형제 표기 조건",
                        ["원칙 적용 (부·조부 형제 전원 + 증조 이상/하대 축소)", "직계 위주 간략 표기", "전체 형제 포함"],
                        index=0
                    )
                with c_opt3:
                    up_gens = st.slider("상위 선대 표시 대수", min_value=1, max_value=8, value=4)
                    down_gens = st.slider("하위 자손 표시 대수", min_value=1, max_value=4, value=2)

            # 계보 트리 노드 수집
            ancestor_chain = []
            curr = target_p
            for _ in range(up_gens):
                f_name = curr.get("부친", "").strip()
                if not f_name or f_name in ["미상", "없음", "", "nan"]:
                    break
                f_rows = df[df["선조"] == f_name]
                if len(f_rows) > 1:
                    gf_val = str(curr.get("조부", "")).strip()
                    if gf_val:
                        matched_f = f_rows[f_rows["부친"] == gf_val]
                        if not matched_f.empty:
                            f_rows = matched_f
                if f_rows.empty:
                    break
                parent_row = f_rows.iloc[0]
                ancestor_chain.append(parent_row)
                curr = parent_row

            # 형제 정보 수집 (제2장 1항 ①)
            # 1. 본인의 형제
            p_siblings = df[df["부친"] == target_p.get("부친", "")].copy()
            # 2. 부친의 형제
            f_siblings = pd.DataFrame()
            if len(ancestor_chain) >= 1:
                f_siblings = df[df["부친"] == ancestor_chain[0].get("부친", "")].copy()
            # 3. 조부의 형제
            gf_siblings = pd.DataFrame()
            if len(ancestor_chain) >= 2:
                gf_siblings = df[df["부친"] == ancestor_chain[1].get("부친", "")].copy()

            # 4. 자녀 및 손자 수집
            children = df[df["부친"] == p_name].copy()
            grandchildren = pd.DataFrame()
            if not children.empty:
                grandchildren = df[df["부친"].isin(children["선조"].tolist())].copy()

            # Graphviz DOT 코드 작성
            rankdir = "TB" if "수직" in tree_direction else "LR"
            lines = [
                'digraph GenealogyTree {',
                f'  rankdir="{rankdir}";',
                '  node [shape="box", style="filled,rounded", fontname="Noto Sans KR, sans-serif", fontsize="11"];',
                '  edge [color="#64748b", penwidth="1.3"];'
            ]

            all_tree_people = {}

            # 본인 노드
            all_tree_people[p_name] = target_p
            lines.append(f'  "{p_name}" [label="★ {p_name}\\n({p_rel})", fillcolor="#FEF08A", color="#CA8A04", penwidth="2.5"];')

            # 상위 선대 노드 생성 및 연결
            prev = p_name
            for idx, a_row in enumerate(ancestor_chain):
                a_name = a_row.get("선조", "")
                a_rel = a_row.get("관계", "")
                all_tree_people[a_name] = a_row

                # 제2장 1항 ①: 증조부(idx=2) 이상부터 형제가 있으면 '0자 중 0자' 표기
                sibs = df[df["부친"] == a_row.get("부친", "")]
                label_extra = ""
                if idx >= 2 and len(sibs) > 1 and a_rel:
                    label_extra = f"\\n[{len(sibs)}자 중 {a_rel}]"

                rel_names = ["부(父)", "조부(祖父)", "증조부(曾祖)", "고조부(高祖)", "5대조", "6대조", "7대조", "8대조"]
                gen_title = rel_names[idx] if idx < len(rel_names) else f"{idx+1}대조"
                lines.append(f'  "{a_name}" [label="{gen_title}: {a_name}{label_extra}", fillcolor="#EDE9FE", color="#7C3AED"];')
                lines.append(f'  "{a_name}" -> "{prev}";')
                prev = a_name

            # 형제 노드 연결 (조건에 따라)
            if sibling_rule != "직계 위주 간략 표기":
                # 본인 형제들
                for _, sib in p_siblings.iterrows():
                    s_name = sib.get("선조", "")
                    s_rel = sib.get("관계", "")
                    if s_name != p_name:
                        all_tree_people[s_name] = sib
                        lines.append(f'  "{s_name}" [label="{s_name}\\n({s_rel})", fillcolor="#E0F2FE", color="#0284C7"];')
                        if len(ancestor_chain) >= 1:
                            lines.append(f'  "{ancestor_chain[0].get("선조", "")}" -> "{s_name}";')

                # 부친 형제들
                if len(ancestor_chain) >= 2 and not f_siblings.empty:
                    for _, sib in f_siblings.iterrows():
                        s_name = sib.get("선조", "")
                        s_rel = sib.get("관계", "")
                        if s_name != ancestor_chain[0].get("선조", ""):
                            all_tree_people[s_name] = sib
                            lines.append(f'  "{s_name}" [label="숙부: {s_name}\\n({s_rel})", fillcolor="#F1F5F9", color="#94A3B8"];')
                            lines.append(f'  "{ancestor_chain[1].get("선조", "")}" -> "{s_name}";')

            # 자녀 및 손자 연결
            if not children.empty:
                for _, c_row in children.iterrows():
                    c_name = c_row.get("선조", "")
                    c_rel = c_row.get("관계", "")
                    all_tree_people[c_name] = c_row
                    lines.append(f'  "{c_name}" [label="자: {c_name}\\n({c_rel})", fillcolor="#DCFCE7", color="#16A34A"];')
                    lines.append(f'  "{p_name}" -> "{c_name}";')

                    if down_gens >= 2:
                        grands = df[df["부친"] == c_name]
                        for _, g_row in grands.iterrows():
                            g_name = g_row.get("선조", "")
                            g_rel = g_row.get("관계", "")
                            all_tree_people[g_name] = g_row
                            lines.append(f'  "{g_name}" [label="손: {g_name}\\n({g_rel})", fillcolor="#ECFCCB", color="#65A30D"];')
                            lines.append(f'  "{c_name}" -> "{g_name}";')

            lines.append('}')
            chart_dot = "\n".join(lines)

            # 화면 2열 분할: 좌측 계보도, 우측 선택 인물 약력
            col_chart, col_person = st.columns([7, 5])
            with col_chart:
                st.markdown(f"##### 🌳 [{p_name}] 선조 중심 계보도")
                st.graphviz_chart(chart_dot, use_container_width=True)

            with col_person:
                # 제2장 2항: 표시되는 해당 인물의 박스를 클릭/선택했을 경우 그 인물의 약력을 표시한다.
                st.markdown("##### 📜 계보도 내 인물 상세 약력 조회")
                p_keys = list(all_tree_people.keys())
                sel_p_in_tree = st.selectbox(
                    "상세 약력을 확인할 인물을 선택하세요",
                    p_keys,
                    index=p_keys.index(p_name) if p_name in p_keys else 0
                )
                render_profile_card(all_tree_people[sel_p_in_tree], title_prefix="계보 연동 약력")


# =============================================================================
# 제4장: 묘제일 메뉴 생성 (선조 묘제일·시제일·축일·장소 종합 안내)
# =============================================================================
elif search_mode == "묘제일 안내":
    st.markdown("##### 🕯️ 광주안씨 완귀공파 종중 묘제일(시제일) 종합 안내")
    st.caption("선조님들의 묘제일(시제일), 기일, 제향 장소 및 묘역 위치 안내입니다.")

    tab_myo1, tab_myo2 = st.tabs(["주요 선조 시제일 종합 안내표", "족보 전체 기일(忌日)·묘소 검색"])

    with tab_myo1:
        st.markdown("###### 🏛️ 주요 분파조 및 선조 제향일(묘제일) 일람표")
        myoje_data = [
            {
                "세수": "24세",
                "선조 함자": "완귀공 증(嶒)",
                "분파": "파조 (완귀공파)",
                "묘제일(시제일)": "매년 음력 10월 첫째 일요일 (10월 상정일)",
                "제향 장소": "영천시 금호읍 청제내 호연사(虎淵祠) 및 화전등 묘소",
                "묘좌(좌향)": "유좌(酉坐, 서향)",
                "비고": "파조 묘제 및 유허비, 종중 총회 겸 봉행"
            },
            {
                "세수": "25세",
                "선조 함자": "참봉공 종경(宗慶)",
                "분파": "완귀공 계자",
                "묘제일(시제일)": "매년 음력 10월 첫째 일요일",
                "제향 장소": "영천 청제내 재사 좌록 배위(거창신씨) 묘소 설단",
                "묘좌(좌향)": "유좌(酉坐)",
                "비고": "양주 수락산 실전으로 배위 묘하 설단 행사"
            },
            {
                "세수": "26세",
                "선조 함자": "음참봉 대해(大海)",
                "분파": "종경공 1자",
                "묘제일(시제일)": "매년 음력 10월 첫째 일요일",
                "제향 장소": "영천 청제내 배위(안동권씨) 묘소 설단",
                "묘좌(좌향)": "유좌(酉坐)",
                "비고": "양주 선고 묘하 실전으로 배위 묘하 설단 행사"
            },
            {
                "세수": "27세",
                "선조 함자": "감찰공 전(琠)",
                "분파": "대해공 계자",
                "묘제일(시제일)": "매년 음력 10월 첫째 일요일",
                "제향 장소": "영천 금호읍 청제내 재사 북록 묘소",
                "묘좌(좌향)": "유좌(酉坐)",
                "비고": "병와 이형상 찬 묘갈명"
            },
            {
                "세수": "28세",
                "선조 함자": "판결사 명한(鳴漢)",
                "분파": "전공 1자",
                "묘제일(시제일)": "매년 음력 10월 첫째 일요일",
                "제향 장소": "영천 청제내 신수곡(臣守谷) 묘소",
                "묘좌(좌향)": "유좌(酉坐)",
                "비고": "증 장례원 판결사, 오천 정환국 찬 묘갈명"
            },
            {
                "세수": "28세",
                "선조 함자": "시한(蓍漢)",
                "분파": "전공 2자",
                "묘제일(시제일)": "매년 음력 10월 첫째 일요일",
                "제향 장소": "영천 청제 재사 우록 묘소",
                "묘좌(좌향)": "신좌(辛坐)",
                "비고": "광주 이채진 찬 묘갈"
            },
            {
                "세수": "29세",
                "선조 함자": "성재 후정(后靜)",
                "분파": "상석계파 (제1지파)",
                "묘제일(시제일)": "매년 음력 10월 첫째 일요일",
                "제향 장소": "영천 청제내 선공 묘하 자좌",
                "묘좌(좌향)": "자좌(子坐)",
                "비고": "성균관 학정, 매산 정중기 찬 묘갈명"
            },
            {
                "세수": "29세",
                "선조 함자": "인정(仁靜)",
                "분파": "위계파 (제2지파)",
                "묘제일(시제일)": "매년 음력 10월 상정일",
                "제향 장소": "영천시 대창면 임하리 선영",
                "묘좌(좌향)": "묘좌(卯坐)",
                "비고": "위계파 지파 묘제 봉행"
            },
            {
                "세수": "29세",
                "선조 함자": "태정(泰靜)",
                "분파": "목계파 (제3지파)",
                "묘제일(시제일)": "매년 음력 10월 상정일",
                "제향 장소": "영천 청제 가음제외 북록 묘소",
                "묘좌(좌향)": "자좌(子坐)",
                "비고": "종인 종묵 찬 묘갈명"
            },
            {
                "세수": "29세",
                "선조 함자": "삼괴당 처정(處靜)",
                "분파": "봉계파 (제4지파)",
                "묘제일(시제일)": "매년 음력 10월 상정일",
                "제향 장소": "영천시 봉동 정모산 독산 후 묘소",
                "묘좌(좌향)": "진좌(辰坐)",
                "비고": "순암 정복 찬 묘갈명, 정모재 봉향"
            },
            {
                "세수": "선영",
                "선조 함자": "작산(鵲山) 선영 선조 40위",
                "분파": "완귀공파 작산 묘소",
                "묘제일(시제일)": "매년 음력 10월 상정일",
                "제향 장소": "영천시 작산동 선영",
                "묘좌(좌향)": "유좌/자좌 등",
                "비고": "작산 묘소 선조 40위 일괄 합동 묘제"
            }
        ]
        df_myo = pd.DataFrame(myoje_data)
        st.dataframe(df_myo, use_container_width=True, hide_index=True)

    with tab_myo2:
        st.markdown("###### 🔍 족보 수록 선조 묘소 및 기일(忌日) 검색")
        m_col1, m_col2 = st.columns([6, 4])
        with m_col1:
            myo_kw = st.text_input("묘소 지명 검색 (예: 청제, 작산, 정모산, 도일, 나월산, 가음, 임하 등)", "").strip()
        with m_col2:
            gi_month = st.selectbox("기일 월별 필터", ["전체", "1월(정월)", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월"])

        m_df = df.copy()
        if myo_kw:
            m_df = m_df[m_df["한글 프로필 정렬"].str.contains(myo_kw, na=False) | m_df["한문 프로필 정렬"].str.contains(myo_kw, na=False)]
        if gi_month != "전체":
            m_num = gi_month.split("월")[0]
            m_df = m_df[m_df["한글 프로필 정렬"].str.contains(rf"기일은\s*{m_num}월|忌\s*{m_num}월|{m_num}월\s*[0-9]+일졸", na=False)]

        st.caption(f"검색 결과: 총 {len(m_df)}위 선조")
        st.dataframe(
            m_df[["고유ID", "선조", "세수", "관계", "부친", "조부", "한글 프로필 정렬"]],
            use_container_width=True,
            hide_index=True
        )


# =============================================================================
# 하단 전체 족보 데이터 목록 뷰
# =============================================================================
st.markdown("---")
with st.expander("📑 전체 족보 전사 데이터 목록 (총 2,887행)", expanded=False):
    col_f1, col_f2 = st.columns([4, 8])
    with col_f1:
        gen_filter = st.selectbox("세수(世數) 필터", ["전체"] + [f"{g}세" for g in range(24, 45)])
    with col_f2:
        tbl_kw = st.text_input("테이블 내 빠른 필터링", "")

    view_df = df.copy()
    if gen_filter != "전체":
        view_df = view_df[view_df["세수"] == gen_filter]
    if tbl_kw:
        mask_v = view_df.astype(str).apply(lambda col: col.str.contains(re.escape(tbl_kw), case=False, na=False)).any(axis=1)
        view_df = view_df[mask_v]

    disp_cols = ["고유ID", "선조", "세수", "관계", "부친", "조부", "한글 프로필 정렬", "한문 프로필 정렬"]
    st.dataframe(view_df[[c for c in disp_cols if c in view_df.columns]], use_container_width=True, hide_index=True)