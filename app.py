# -*- coding: utf-8 -*-
"""
광주안씨 완귀공파 종중 족보 통합 검색 시스템 (Streamlit Web Application)
- 작성 지침: '족보 검색 메뉴 PY 작성.pdf' (제1장 ~ 제5장 전 조항 완벽 반영)
  * 제1장 (초기 총괄 메뉴 작성): 페이지 타이틀, 아이콘(icon.png), 검색창 타이틀(족보 검색),
                              선택 안내문('아래 검색 방식 중 하나를 선택하세요'),
                              인명 검색(인명, 자, 호, 초위, 초자, 통명, 초명 등 -> 정렬된 약력, 복수 인물 선택 과정),
                              단어 검색(약력 내 단어 포함 인물 검색 -> 핵심 약력 표시)
  * 제2장 (계보 작성): 초기 화면에 계보 검색 버튼 생성, 인명 검색어와 동일 검색어 지원,
                     계보 작성 조건 창(부·조부 형제 전원 표기, 증조부 이상 및 하대 '0자 중 0자' 표기,
                     축소 시 수평 지그재그 계보도), 노드 선택 시 상세 약력 연동
  * 제3장 (디자인): 종중의 상징 아이콘(icon.png), 배경 이미지(background.png), 동영상 삽입
  * 제4장 (파이션 입력 사항):
      1. 페이지 타이틀: 함초롱돋움체, '광주안씨 완귀공파 종중'(폰트 22, 짙은 청색) / '족보 검색 시스템'(폰트 20, 검정 글자) 줄바꿈
      2. 타이틀 아래 참고 글: '광주안씨 완귀공(휘 嶒) 후손 24세 ~ 44세 총 2,887위 (고위 기준)'
      3. 완귀정 소개 글 (제목, 건립자, 위치, 유래, 제향)
      4. 족보 검색 창: 타이틀 '족보 검색(아래 검색 방식을 선택하세요)', 방식(인명, 단어, 계보),
                     입력창 너비 최소화 & 오른쪽에 안내 문구 표시
      5. 검색 결과창:
         ① 세수(완귀공의 0대 손), 부, 조부, 자, 호, 초휘(初諱), 일휘(一諱), 초자(初字), 통명(通名), 초명(初名), 과거·관직, 배위 등을 한글한자 병기
         ② 별도로 아래에 상세한 한문·한글 약력 (정리된 한문 한글 약력 표기)
         ③ 마지막에 족보 수록 원문 보기 생성
  * 제5장 (묘제일 메뉴 생성): 주요 선조별 묘제일(시제일) 종합 일람표 및 족보 수록 선조 묘소·기일 검색
"""

import os
import re
import base64
import pandas as pd
import streamlit as st

# =============================================================================
# 제1장 1항: 페이지 기본 설정 및 아이콘(icon.png) 지정
# =============================================================================
ICON_FILE = "icon.png"
BG_FILE = "background.png"
PAGE_TITLE = "광주안씨 완귀공파 종중 족보 검색 시스템"

st.set_page_config(
    page_title=PAGE_TITLE,
    page_icon=ICON_FILE if os.path.exists(ICON_FILE) else "📜",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =============================================================================
# 제3장 2항 & 제4장 1항: 스타일, 폰트(함초롱돋움), 배경 이미지(background.png)
# =============================================================================
bg_style = ""
if os.path.exists(BG_FILE):
    try:
        with open(BG_FILE, "rb") as f_bg:
            b64_bg = base64.b64encode(f_bg.read()).decode()
        bg_style = f"""
        .stApp {{
            background: linear-gradient(rgba(255, 255, 255, 0.92), rgba(255, 255, 255, 0.95)),
                        url("data:image/png;base64,{b64_bg}") no-repeat center center fixed !important;
            background-size: cover !important;
        }}
        """
    except Exception:
        pass

custom_css = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@400;600;700&family=Noto+Sans+KR:wght@300;400;500;700&display=swap');

/* 전체 여백 및 컴포넌트 상하 간격 극소화 (최대 압축 컴팩트 레이아웃) */
html, body, [class*="css"] {
    font-family: 'Noto Sans KR', 'HCR Dotum', sans-serif;
}

.block-container {
    padding-top: 0.2rem !important;
    padding-bottom: 1.2rem !important;
    padding-left: 0.8rem !important;
    padding-right: 0.8rem !important;
    max-width: 1400px;
}

header[data-testid="stHeader"] {
    height: 0px !important;
    min-height: 0px !important;
    background: transparent !important;
    display: none !important;
}

/* Streamlit Cloud 우측 하단 Manage app 박스 및 배지/푸터 숨기기 */
#MainMenu { visibility: hidden !important; }
footer { visibility: hidden !important; }
[data-testid="manage-app-button"] { display: none !important; }
.viewerBadge_container__r5tak, .viewerBadge_link__qRIco { display: none !important; }
div[class*="viewerBadge"] { display: none !important; }
div[class*="manage-app"] { display: none !important; }
div[data-testid="stStatusWidget"] { display: none !important; }
div[data-testid="stDecoration"] { display: none !important; }

[data-testid="stVerticalBlock"] {
    gap: 0.25rem !important;
}

[data-testid="stHorizontalBlock"] {
    gap: 0.4rem !important;
    align-items: center;
}

div[data-testid="stElementContainer"] {
    margin-top: 0px !important;
    margin-bottom: 0px !important;
    padding-top: 0px !important;
    padding-bottom: 0px !important;
}

/* 상단 헤더: 아이콘 오른쪽에 타이틀 가로 나란히 배치 (모바일 화면에서도 줄바꿈 없이 가로 정렬 고정) */
.clan-header-wrapper {
    display: flex;
    flex-direction: row;
    align-items: center;
    gap: 14px;
    margin-top: 2px;
    margin-bottom: 12px;
    padding-bottom: 2px;
}

.clan-icon-img {
    width: 76px;
    height: 76px;
    object-fit: contain;
    flex-shrink: 0;
    background: transparent !important;
    background-color: transparent !important;
    border: none !important;
    box-shadow: none !important;
}

.clan-title-box {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
    line-height: 1.22;
    margin: 0;
    padding: 0;
    flex-grow: 1;
}

.clan-title-main {
    font-size: 22px !important;
    font-weight: 700 !important;
    color: #0f2b5c !important; /* 짙은 청색 */
    letter-spacing: -0.5px;
    display: block;
    line-height: 1.22;
}

.clan-title-sub {
    font-size: 20px !important;
    font-weight: 700 !important;
    color: #111111 !important; /* 검정 글자 */
    letter-spacing: -0.5px;
    display: block;
    line-height: 1.22;
}

.clan-title-ref {
    font-size: 13px !important;
    color: #4b5563;
    font-weight: 500;
    margin-top: 2px;
    line-height: 1.2;
}

/* 완귀정 소개 expander */
div[data-testid="stExpander"] {
    border: 1px solid #dbe2ea !important;
    border-radius: 6px !important;
    margin-top: 6px !important;
    margin-bottom: 6px !important;
    background-color: rgba(255, 255, 255, 0.92) !important;
}

div[data-testid="stExpander"] summary {
    padding: 6px 12px !important;
    font-size: 16px !important;
    font-weight: 700 !important;
    color: #0f2b5c !important;
}

div[data-testid="stExpander"] div[data-testid="stExpanderDetails"] {
    padding: 8px 14px !important;
}

.wanguijeong-card {
    background-color: #fbfbfe;
    border: 1px solid #dbe2ea;
    border-left: 5px solid #1a365d;
    padding: 12px 16px !important;
    border-radius: 6px;
    margin-bottom: 4px !important;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.03);
}

.wanguijeong-card h4 {
    margin: 0 0 8px 0 !important;
    font-size: 17.5px !important;
    color: #1a365d;
    font-weight: 700;
}

.wanguijeong-card ul {
    margin: 0 !important;
    padding-left: 20px !important;
    font-size: 15px !important;
    color: #1e293b;
    line-height: 1.75 !important;
}

.wanguijeong-card li {
    margin-bottom: 6px !important;
}

.wanguijeong-card li b {
    color: #0f2b5c;
    font-weight: 700;
}

/* 검색창 헤더 박스 */
.search-header-box {
    background: #f1f5f9;
    border-left: 4px solid #0f2b5c;
    padding: 4px 10px !important;
    border-radius: 5px;
    margin-top: 2px !important;
    margin-bottom: 3px !important;
}

.search-header-box h3 {
    margin: 0 !important;
    font-size: 15px !important;
    color: #0f172a;
    font-weight: 700;
    line-height: 1.25 !important;
}

/* 인명, 단어, 계보 라디오 버튼 선택 문구 중앙 정렬 */
div[data-testid="stRadio"] {
    display: flex !important;
    justify-content: center !important;
    align-items: center !important;
    width: 100% !important;
    margin-top: 6px !important;
    margin-bottom: 8px !important;
    text-align: center !important;
}

div[data-testid="stRadio"] > div[role="radiogroup"],
div[data-testid="stRadio"] > div {
    display: flex !important;
    justify-content: center !important;
    align-items: center !important;
    gap: 24px !important; /* 선택 문구 간 여유로운 간격 */
    width: 100% !important;
    margin: 0 auto !important;
}

div[data-testid="stRadio"] label {
    font-size: 15px !important;
    font-weight: 600 !important;
    color: #0f2b5c !important;
    cursor: pointer !important;
    padding: 2px 8px !important;
}

/* 폼, 입력창, 버튼 뚜렷한 강조 스타일 */
form[data-testid="stForm"] {
    border: none !important; /* 중앙 테두리 삭제: 폼 자체의 내부 테두리 완전 제거 */
    padding: 6px 4px !important;
    border-radius: 0px !important;
    margin-top: 2px !important;
    margin-bottom: 2px !important;
    background-color: transparent !important;
    box-shadow: none !important;
}

div[data-testid="stTextInput"] {
    margin-bottom: 0px !important;
}

/* 검색어 입력창: 테두리 3배 확대 (두께 3.5px), 색상 검정색 */
div[data-testid="stTextInput"] input {
    height: 40px !important;
    padding: 4px 12px !important;
    font-size: 14.5px !important;
    font-weight: 600 !important;
    color: #000000 !important;
    background-color: #ffffff !important;
    border: 3.5px solid #000000 !important; /* 3배 확대된 굵은 검정 테두리 */
    border-radius: 6px !important;
    box-shadow: inset 0 1px 2px rgba(0, 0, 0, 0.06) !important;
    transition: all 0.2s ease-in-out !important;
}

div[data-testid="stTextInput"] input:focus {
    border-color: #000000 !important;
    box-shadow: 0 0 0 2px rgba(0, 0, 0, 0.3) !important;
    outline: none !important;
}

div[data-testid="stTextInput"] input::placeholder {
    color: #d1d5db !important; /* 매우 희미한 색 */
    opacity: 0.55 !important;
    font-size: 13.5px !important;
}

div[data-testid="stTextInput"] input::-webkit-input-placeholder {
    color: #d1d5db !important; /* 매우 희미한 색 */
    opacity: 0.55 !important;
}

div[data-testid="stTextInput"] input::-moz-placeholder {
    color: #d1d5db !important; /* 매우 희미한 색 */
    opacity: 0.55 !important;
}

/* 검색 실행 버튼도 38px로 높이를 맞추고 그림자 추가 */
div[data-testid="stButton"] button, div[data-testid="stFormSubmitButton"] button {
    height: 38px !important;
    padding: 4px 16px !important;
    font-size: 14px !important;
    font-weight: 600 !important;
    min-height: 38px !important;
    line-height: 1 !important;
    margin-top: 0px !important;
    border-radius: 6px !important;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1) !important;
}

/* 검색창 우측 안내문구 */
.search-hint-text {
    padding-top: 3px !important;
    font-size: 12.5px !important;
    color: #475569;
    font-weight: 500;
    line-height: 1.35 !important;
}

/* 결과 카드 */
.result-card {
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    padding: 8px 12px !important;
    box-shadow: 0 2px 5px rgba(0, 0, 0, 0.04);
    margin-top: 3px !important;
    margin-bottom: 4px !important;
}

.result-card-title {
    font-family: 'Noto Serif KR', serif;
    font-size: 17px !important;
    font-weight: 700;
    color: #0f2b5c;
    border-bottom: 1.5px solid #e2e8f0;
    padding-bottom: 3px !important;
    margin-bottom: 5px !important;
}

.result-summary-para {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-left: 4.5px solid #0f2b5c;
    border-radius: 6px;
    padding: 10px 14px !important;
    margin-bottom: 8px !important;
}

.info-para-line {
    margin: 0 0 5px 0 !important;
    font-size: 13.5px !important;
    line-height: 1.6 !important;
    color: #1f2937;
}

.info-para-line:last-child {
    margin-bottom: 0px !important;
}

.info-label {
    font-weight: 700 !important;
    color: #0f2b5c !important;
}

.info-val {
    color: #111827 !important;
    font-weight: 500;
}

.info-sep {
    color: #94a3b8 !important;
    margin: 0 8px !important;
}

.result-gid-badge {
    font-size: 13px !important;
    font-weight: normal !important;
    color: #64748b !important;
    margin-left: 8px;
}

.hanja-box {
    font-family: 'Noto Serif KR', 'Batang', serif;
    background: #fafaf9;
    border: 1px solid #e2e8f0;
    padding: 6px 10px !important;
    border-radius: 6px;
    line-height: 1.55 !important;
    font-size: 13px !important;
    color: #18181b;
    margin-top: 2px !important;
    margin-bottom: 3px !important;
}

.hangul-box {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    padding: 6px 10px !important;
    border-radius: 6px;
    line-height: 1.5 !important;
    font-size: 12.5px !important;
    color: #1e293b;
    margin-top: 2px !important;
    margin-bottom: 3px !important;
}

.condition-box {
    background: #eff6ff;
    border: 1px solid #bfdbfe;
    border-left: 4px solid #2563eb;
    padding: 5px 8px !important;
    border-radius: 6px;
    margin-bottom: 4px !important;
    font-size: 12px !important;
    color: #1e40af;
    line-height: 1.35 !important;
}

/* 알림 및 캡션 상하 간격 */
div[data-testid="stAlert"] {
    padding: 4px 10px !important;
    margin: 2px 0 !important;
}

""" + bg_style + """
/* 족보 검색창: 안쪽/바깥쪽 중복 테두리 삭제, 단일 테두리로 두께 1/3 축소 (1.3px 청색) */
div[data-testid="stVerticalBlockBorderWrapper"] {
    margin-top: 25px !important; /* 아래로 살짝 하향 유지 */
    margin-bottom: 20px !important;
    background-color: #ffffff !important;
    border-radius: 8px !important;
    padding: 14px 18px !important;
    
    /* 테두리 두께 1/3로 축소 (4px -> 1.3px), 색상 청색 */
    border: 1.3px solid #1e40af !important;
    outline: none !important; /* 바깥 테두리 삭제 */
    box-shadow: none !important; /* 바깥 중복선으로 보일 수 있는 그림자 삭제 */
}

/* 안쪽 테두리 완전 삭제: 내부 stVerticalBlock 및 form에 중복 생성되던 안쪽 테두리 제거 */
div[data-testid="stVerticalBlockBorderWrapper"] > div[data-testid="stVerticalBlock"],
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stVerticalBlock"],
div[data-testid="stVerticalBlockBorderWrapper"] form[data-testid="stForm"],
form[data-testid="stForm"] {
    border: none !important; /* 안쪽 테두리 완전 제거 */
    outline: none !important;
    box-shadow: none !important;
    background: transparent !important;
    padding: 0px !important;
    margin-top: 2px !important;
    margin-bottom: 2px !important;
}

/* stVerticalBlockBorderWrapper가 없는 환경 대비 단일 Fallback */
div[data-testid="stVerticalBlock"]:not([data-testid="stVerticalBlockBorderWrapper"] *):has(> div:first-child .unified-search-title) {
    margin-top: 25px !important;
    margin-bottom: 20px !important;
    background-color: #ffffff !important;
    border-radius: 8px !important;
    padding: 14px 18px !important;
    border: 1.3px solid #1e40af !important;
    outline: none !important;
    box-shadow: none !important;
}

/* 족보 검색 타이틀: 테두리 1/5로 축소 (1px 실선), 색상 검정색 (중앙 가로 분할선 삭제) */
.unified-search-title {
    font-size: 16px;
    font-weight: 700;
    color: #000000 !important;
    border: 1px solid #000000 !important; /* 1/5로 축소된 1px 검정 테두리 */
    border-radius: 4px !important;
    padding: 3px 14px !important;
    background-color: #f8fafc !important;
    display: inline-block !important;
    margin: 0 auto !important;
}

/* 결과 카드 외곽: 상단 창 테두리 중앙에 타이틀 배치 */
.result-outer-card {
    position: relative;
    border: 2px solid #0f2b5c;
    border-radius: 8px;
    padding: 18px 16px 14px 16px;
    margin-top: 18px;
    margin-bottom: 10px;
    background-color: #ffffff;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
}

.result-card-top-title {
    position: absolute;
    top: -13px;
    left: 50%;
    transform: translateX(-50%);
    background: #ffffff;
    padding: 2px 18px;
    font-size: 14.5px;
    font-weight: 700;
    color: #0f2b5c;
    border: 2px solid #0f2b5c;
    border-radius: 14px;
    white-space: nowrap;
    z-index: 2;
    letter-spacing: -0.3px;
}

.result-person-row {
    display: flex;
    flex-direction: row;
    align-items: center;
    flex-wrap: wrap;
    gap: 8px;
    padding-bottom: 8px;
    border-bottom: 1px dashed #cbd5e1;
    margin-bottom: 8px;
}

.person-icon-badge {
    font-size: 20px;
    line-height: 1;
}

.person-name-hangul {
    font-family: 'HCR Dotum', 'Noto Sans KR', sans-serif;
    font-size: 20px;
    font-weight: 700;
    color: #0f2b5c;
    letter-spacing: -0.5px;
}

.person-name-hanja {
    font-family: 'Noto Serif KR', 'Batang', serif;
    font-size: 18px;
    font-weight: 600;
    color: #334155;
    margin-left: 2px;
}

.person-rel-badge {
    font-size: 14.5px;
    font-weight: 600;
    color: #1e3a8a;
    background-color: #eff6ff;
    padding: 2px 8px;
    border-radius: 4px;
    border: 1px solid #bfdbfe;
}

.person-gid-text {
    font-size: 12.5px;
    color: #64748b;
    margin-left: auto;
}

/* 주요 약력: 테두리 표시 뚜렷하게 */
.bio-summary-box {
    border: 2px solid #0f2b5c; /* 테두리 뚜렷하게 */
    border-radius: 6px;
    padding: 10px 14px;
    margin-top: 6px;
    margin-bottom: 8px;
    background-color: #fbfcfe;
}

.bio-summary-header {
    font-size: 14.5px;
    font-weight: 700;
    color: #0f2b5c;
    border-bottom: 1.5px solid #cbd5e1;
    padding-bottom: 4px;
    margin-bottom: 8px;
}

.bio-item-line {
    margin: 0 0 6px 0;
    font-size: 14px;
    line-height: 1.6;
    color: #1f2937;
}

.bio-item-line:last-child {
    margin-bottom: 0;
}

.bio-label {
    font-weight: 700;
    color: #0f2b5c;
    margin-right: 4px;
}

.bio-val {
    color: #111827;
    font-weight: 500;
}

.spouse-line {
    margin-top: 4px;
    padding-left: 12px;
    color: #1e293b;
    line-height: 1.55;
}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)


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
        xl = pd.ExcelFile(target_path)
        sheet = "족보전사결과" if "족보전사결과" in xl.sheet_names else xl.sheet_names[0]
        df = xl.parse(sheet)

        df.columns = [str(col).replace("\ufeff", "").strip() for col in df.columns]

        aliases = {
            "인명": "선조",
            "한글 약력 정렬": "한글 프로필 정렬",
            "한문 약력 정렬": "한문 프로필 정렬",
            "한글 약력 원문": "한글 프로필",
            "한문 약력 원문": "한문 프로필 원문",
        }
        df = df.rename(columns=aliases)

        df = df.fillna("")
        for col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": "", "NaN": "", "None": ""})

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
# 제4장 1항 & 2항: 페이지 타이틀(함초롱돋움 22/20pt) 및 참고 글 (아이콘 옆 가로 배치)
# =============================================================================
b64_icon_html = ""
if os.path.exists(ICON_FILE):
    try:
        from PIL import Image
        import io
        img_icon = Image.open(ICON_FILE).convert("RGBA")
        img_icon.thumbnail((160, 160), Image.Resampling.LANCZOS)
        buf_icon = io.BytesIO()
        img_icon.save(buf_icon, format="PNG", optimize=True)
        b64_icon = base64.b64encode(buf_icon.getvalue()).decode()
        b64_icon_html = f'<img src="data:image/png;base64,{b64_icon}" class="clan-icon-img" alt="완귀정 현판">'
    except Exception:
        with open(ICON_FILE, "rb") as f_ic:
            b64_icon = base64.b64encode(f_ic.read()).decode()
        b64_icon_html = f'<img src="data:image/png;base64,{b64_icon}" class="clan-icon-img" alt="완귀정 현판">'
else:
    b64_icon_html = '<div style="font-size: 42px; line-height: 1; flex-shrink: 0;">📜</div>'

st.markdown(f"""
<div class="clan-header-wrapper">
    {b64_icon_html}
    <div class="clan-title-box">
        <span class="clan-title-main">광주안씨 완귀공파 종중</span>
        <span class="clan-title-sub">족보 검색 시스템</span>
        <div class="clan-title-ref">광주안씨 완귀공(휘 嶒) 후손 24세 ~ 44세 총 2,887위 (고위 기준)</div>
    </div>
</div>
""", unsafe_allow_html=True)


# =============================================================================






# =============================================================================
# =============================================================================
# 제4장 5항 ①: 세부 항목 파싱 및 렌더링 함수
# =============================================================================
def split_name(name_raw):
    m_h = re.findall(r"[一-龥]+", str(name_raw))
    m_k = re.findall(r"[가-힣]+", str(name_raw))
    hanja = "".join(m_h)
    hangul = "".join(m_k)
    if not hangul and not hanja:
        hangul = str(name_raw)
    return hangul, hanja

def extract_tomb_clean(kp, hp):
    p_kp = kp.split("배위")[0]
    p_hp = hp.split("配")[0]
    m = re.search(r"(?:^|,\s*|\.\s*)(?:묘는|묘소는|묘소가|묘소|◯묘는|◯墓는|◯墓|墓는|墓)\s*(.+?)(?=(?:,\s*(?:헌종|고종|순조|정조|숙종|영조|철종|경종|인조|효종|현종|명종|선조|중종|성종)|설서|유허비|$))", p_kp, re.DOTALL)
    if m:
        return m.group(1).strip().rstrip(",")
    m_h = re.search(r"(?:^|,\s*)(?:墓\s*(?:는)?|◯墓)\s*([^\r\n]+)", p_hp)
    if m_h:
        return m_h.group(1).strip().rstrip(",")
    return ""

def extract_spouses(kp, hp):
    spouses = []
    if "배위" in kp:
        k_parts = kp.split("배위")[1:]
        for idx, k_text in enumerate(k_parts):
            val = k_text.strip().rstrip(",")
            if val and val not in spouses:
                spouses.append(val)
    elif "配" in hp:
        h_parts = hp.split("配")[1:]
        for idx, h_text in enumerate(h_parts):
            val = h_text.strip().rstrip(",")
            if val and val not in spouses:
                spouses.append(val)
    return spouses

def parse_detailed_facts(row):
    kp = str(row.get("한글 프로필 정렬", "") or row.get("한글 프로필", ""))
    hp = str(row.get("한문 프로필 정렬", "") or row.get("한문 프로필 원문", ""))
    uid = str(row.get("고유ID", ""))
    
    parts = uid.split("-")
    gen_num = int(parts[-1]) if len(parts) >= 4 and parts[-1].isdigit() else 0
    if gen_num == 24:
        se_su = "24세손 (파조 완귀공)"
    elif gen_num > 24:
        se_su = f"{gen_num}세손 (완귀공의 {gen_num - 24}대손)"
    elif gen_num > 0:
        se_su = f"{gen_num}세손"
    else:
        se_su = ""
    
    main_hp = hp.split("配")[0]
    main_kp = kp.split("배위")[0]
    
    items = {}
    if se_su:
        items["세수"] = se_su
    
    f_val = str(row.get("부친", "")).strip()
    gf_val = str(row.get("조부", "")).strip()
    if f_val in ["미상", "없음", "nan", "None"]: f_val = ""
    if gf_val in ["미상", "없음", "nan", "None"]: gf_val = ""
    
    if f_val and gf_val:
        items["부·조부"] = f"부 {f_val} / 조부 {gf_val}"
    elif f_val:
        items["부"] = f_val
    elif gf_val:
        items["조부"] = gf_val
        
    for tag_h, tag_k, key in [
        ("字", r"(?:자는|자)", "자"),
        ("初字", r"(?:초자는|초자)", "초자"),
        ("號", r"(?:호는|호)", "호"),
        ("初諱", r"(?:초휘는|초휘)", "초휘"),
        ("一諱", r"(?:일휘는|일휘)", "일휘"),
        ("通名", r"(?:통명은|통명)", "통명"),
        ("初名", r"(?:초명은|초명)", "초명"),
    ]:
        mh = re.search(rf"{tag_h}\s*([一-龥]+)", main_hp)
        mk = re.search(rf"{tag_k}\s*([가-힣]+)", main_kp)
        if mh and mk:
            items[key] = f"{mh.group(1)}({mk.group(1)})"
        elif mh:
            items[key] = mh.group(1)
        elif mk:
            items[key] = mk.group(1)
            
    offices = []
    career_pairs = [
        ("文科", "문과"), ("司馬", "사마"), ("生員", "생원"), ("進士", "진사"),
        ("登第", "등제"), ("參奉", "참봉"), ("主簿", "주부"), ("司憲府監察", "사헌부감찰"),
        ("刑曹員外郎", "형조원외랑"), ("佐郎", "좌랑"), ("司書", "사서"), ("說書", "설서"),
        ("成均館學正", "성균관학정"), ("學諭", "학유"), ("學錄", "학록"), ("典籍", "전적"),
        ("通德郎", "통덕랑"), ("通政大夫", "통정대부"), ("掌隷院判決事", "장례원판결사"),
        ("縣監", "현감"), ("郡守", "군수"), ("都事", "도사"), ("洗馬", "세마"),
        ("校長", "교장"), ("敎授", "교수"), ("專門醫", "전문의")
    ]
    for hj, hg in career_pairs:
        if hj in hp or hg in kp:
            offices.append(f"{hj}({hg})")
    if offices:
        items["과거 및 관직"] = ", ".join(offices[:4])
        
    tomb = extract_tomb_clean(kp, hp)
    if tomb:
        items["묘"] = tomb
        
    sps = extract_spouses(kp, hp)
    if sps:
        items["배위"] = sps
        
    return items

def render_search_result_view(row, title_prefix="인명 검색결과"):
    gid = row.get("고유ID", "")
    name_raw = row.get("선조", "")
    rel = row.get("관계", "")
    kp = row.get("한글 프로필 정렬", "") or row.get("한글 프로필", "") or "약력 없음"
    hp = row.get("한문 프로필 정렬", "") or row.get("한문 프로필 원문", "") or "약력 없음"
    kp_raw = row.get("한글 프로필", "")
    hp_raw = row.get("한문 프로필 원문", "")
    
    hangul, hanja = split_name(name_raw)
    items = parse_detailed_facts(row)
    
    lines_html = []
    for k in ["세수", "부·조부", "부", "조부", "자", "초자", "호", "초휘", "일휘", "통명", "초명", "과거 및 관직", "묘"]:
        if k in items and items[k]:
            lines_html.append(f'<p class="bio-item-line"><span class="bio-label">{k} :</span> <span class="bio-val">{items[k]}</span></p>')
            
    if "배위" in items and items["배위"]:
        sps = items["배위"]
        if len(sps) == 1:
            lines_html.append(f'<p class="bio-item-line"><span class="bio-label">배위 :</span> <span class="bio-val">{sps[0]}</span></p>')
        else:
            circ_nums = ["①", "②", "③", "④", "⑤", "⑥"]
            sp_lines = "".join([f'<div class="spouse-line">{circ_nums[i] if i < len(circ_nums) else f"({i+1})"} {s}</div>' for i, s in enumerate(sps)])
            lines_html.append(f'<div class="bio-item-line"><span class="bio-label">배위 :</span>{sp_lines}</div>')
            
    bio_content = "".join(lines_html)
    
    st.markdown(f"""
    <div class="result-outer-card">
        <div class="result-card-top-title">{title_prefix}</div>
        
        <div class="result-person-row">
            <span class="person-icon-badge">📜</span>
            <span class="person-name-hangul">{hangul}</span>
            <span class="person-name-hanja">{hanja}</span>
            <span class="person-rel-badge">({rel})</span>
            <span class="person-gid-text">고유ID: {gid}</span>
        </div>
        
        <div class="bio-summary-box">
            <div class="bio-summary-header">주요 약력</div>
            <div class="bio-summary-body">
                {bio_content}
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div style="font-weight: 700; font-size: 13.5px; color: #1e3a8a; margin: 4px 0 2px 0;">🇰🇷 정리된 한글 약력</div>
    <div class='hangul-box'>{kp}</div>
    <div style="font-weight: 700; font-size: 13.5px; color: #1e3a8a; margin: 5px 0 2px 0;">漢 정리된 한문 약력</div>
    <div class='hanja-box'>{hp}</div>
    """, unsafe_allow_html=True)
    
    with st.expander("📖 족보 수록 원문(原文字句) 보기", expanded=False):
        c1, c2 = st.columns(2)
        with c1:
            st.caption("한글 번역 원문")
            st.write(kp_raw if kp_raw else "-")
        with c2:
            st.caption("한문 원문")
            st.write(hp_raw if hp_raw else "-")

# =============================================================================
# 족보 검색 통합 창 (족보 검색, 검색 방식, 입력 창, 검색실행 버튼, 입력 요령 - 단일 창 통합)
# =============================================================================
person_query, word_query, tree_query = "", "", ""
with st.container(border=True):
    st.markdown('<div style="text-align: center; margin-bottom: 8px;"><div class="unified-search-title">🔍 족보 검색 <span style="font-size: 13px; font-weight: normal; color: #374151;">(아래 검색 방식을 선택하세요)</span></div></div>', unsafe_allow_html=True)
    
    search_mode = st.radio(
        "검색 방식 선택",
        ["인명", "단어", "계보"],
        horizontal=True,
        index=0,
        label_visibility="collapsed"
    )
    
    if search_mode == "인명":
        with st.form(key="search_person_form"):
            col_input, col_hint = st.columns([3.5, 6.5])
            with col_input:
                person_query = st.text_input(
                    "인명 검색어 입력",
                    placeholder=" 여기에 입력하세요",
                    label_visibility="collapsed"
                ).strip()
                p_submit = st.form_submit_button("🔍 검색 실행", use_container_width=True, type="primary")
            with col_hint:
                st.markdown("""
                <div class="search-hint-text">
                    👉 <b>인물의 함자, 자(字), 호(號), 초휘(初諱), 일휘(一諱), 초자(初字), 통명(通名), 초명(初名) 등을 입력</b>
                </div>
                """, unsafe_allow_html=True)
    elif search_mode == "단어":
        with st.form(key="search_word_form"):
            col_input, col_hint = st.columns([3.5, 6.5])
            with col_input:
                word_query = st.text_input(
                    "단어 검색어 입력",
                    placeholder=" 여기에 입력하세요",
                    label_visibility="collapsed"
                ).strip()
                w_submit = st.form_submit_button("🔍 검색 실행", use_container_width=True, type="primary")
            with col_hint:
                st.markdown("""
                <div class="search-hint-text">
                    👉 <b>약력에 포함된 단어(관직, 지명, 묘소, 비문 찬술자, 사적 등)를 입력</b>
                </div>
                """, unsafe_allow_html=True)
    elif search_mode == "계보":
        with st.form(key="search_tree_form"):
            col_input, col_hint = st.columns([3.5, 6.5])
            with col_input:
                tree_query = st.text_input(
                    "계보 검색어 입력",
                    placeholder=" 여기에 입력하세요",
                    label_visibility="collapsed"
                ).strip()
                t_submit = st.form_submit_button("🌳 계보 작성 및 검색", use_container_width=True, type="primary")
            with col_hint:
                st.markdown("""
                <div class="search-hint-text">
                    👉 <b>인물의 함자, 자(字), 호(號) 입력 후 계보 작성 조건에 따라 계보도 생성</b>
                </div>
                """, unsafe_allow_html=True)

# =============================================================================
# 검색 결과 처리
# =============================================================================
if search_mode == "인명":

    if person_query:
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
            st.warning(f"⚠️ '{person_query}' 조건과 일치하는 인물 기록이 없습니다.")
        else:
            # 제1장 3항 ②: 해당 인물이 여러 명일 경우에는 검색 결과 표시 전에 인물 선택 과정을 거친다.
            if len(matched_df) == 1:
                selected_row = matched_df.iloc[0]
                render_search_result_view(selected_row, title_prefix="인명 검색 결과")
            else:
                st.info(f"💡 '{person_query}' 검색 결과 총 {len(matched_df)}명이 조회되었습니다. 조회할 선조를 선택하세요.")
                opts = [
                    f"[{r['고유ID']}] {r['선조']} ({r.get('세수', '')}, {r.get('관계', '')}) | 부: {r.get('부친', '')} / 조부: {r.get('조부', '')}"
                    for _, r in matched_df.iterrows()
                ]
                sel_opt = st.selectbox("선조 선택 (인물 선택 과정)", opts)
                sel_gid = sel_opt.split("]")[0].replace("[", "").strip()
                selected_row = matched_df[matched_df["고유ID"] == sel_gid].iloc[0]
                render_search_result_view(selected_row, title_prefix="인명 검색 결과")


# =============================================================================
# 검색 방식 2: 단어 검색 (제4장 4항 ③ 너비 최소화 & 우측 문구 표시)
# =============================================================================
elif search_mode == "단어":

    if word_query:
        search_cols = ["한글 프로필 정렬", "한문 프로필 정렬", "한글 프로필", "한문 프로필 원문"]
        mask = df[search_cols].astype(str).apply(
            lambda col: col.str.contains(re.escape(word_query), case=False, na=False)
        ).any(axis=1)
        matched_df = df[mask]

        if matched_df.empty:
            st.warning(f"⚠️ 단어 '{word_query}'(을)를 포함하고 있는 약력 기록이 없습니다.")
        else:
            st.success(f"총 {len(matched_df)}명의 약력에서 단어 '{word_query}'(이)가 검색되었습니다.")
            opts = [
                f"[{r['고유ID']}] {r['선조']} ({r.get('세수', '')}, {r.get('관계', '')}) | 부: {r.get('부친', '')}"
                for _, r in matched_df.iterrows()
            ]
            sel_word_person = st.selectbox("핵심 약력을 조회할 선조 선택", opts)
            sel_gid = sel_word_person.split("]")[0].replace("[", "").strip()
            selected_row = matched_df[matched_df["고유ID"] == sel_gid].iloc[0]

            render_search_result_view(selected_row, title_prefix=f"핵심 약력 ('{word_query}' 검색)")

            st.markdown("---")
            st.markdown(f"###### 📋 '{word_query}' 검색 결과 인물 목록 ({len(matched_df)}명)")
            show_cols = ["고유ID", "선조", "세수", "관계", "부친", "조부", "한글 프로필 정렬"]
            st.dataframe(matched_df[[c for c in show_cols if c in matched_df.columns]], use_container_width=True)


# =============================================================================
# 검색 방식 3: 계보 검색 (제2장 계보 작성 요령 전 조항 완벽 반영)
# =============================================================================
elif search_mode == "계보":

    if tree_query:
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
                selected_t_opt = st.selectbox("계보도를 작성할 기준 선조 선택", t_opts)
                sel_gid = selected_t_opt.split("]")[0].replace("[", "").strip()
                target_p = tree_matches[tree_matches["고유ID"] == sel_gid].iloc[0]
            else:
                target_p = tree_matches.iloc[0]

            p_name = target_p.get("선조", "")
            p_gid = target_p.get("고유ID", "")
            p_rel = target_p.get("관계", "")

            # 제2장 3항 ①: 계보 작성 조건 창
            with st.expander("⚙️ 계보 작성 조건 창 (제2장 1항 ① 규정)", expanded=True):
                c_opt1, c_opt2, c_opt3 = st.columns(3)
                with c_opt1:
                    tree_direction = st.radio("계보 전개 방향", ["수직 전개 (위→아래)", "수평 지그재그 전개 (좌→우)"], index=0)
                with c_opt2:
                    sibling_rule = st.radio(
                        "형제 표기 요령",
                        ["원칙 적용 (부·조부 형제 전원 + 증조 이상/하대 축소)", "직계 위주 표기", "전체 형제 포함"],
                        index=0
                    )
                with c_opt3:
                    up_gens = st.slider("상위 선대 표시 대수", min_value=1, max_value=8, value=4)
                    down_gens = st.slider("하위 자손 표시 대수", min_value=1, max_value=4, value=2)

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

            p_siblings = df[df["부친"] == target_p.get("부친", "")].copy()
            f_siblings = pd.DataFrame()
            if len(ancestor_chain) >= 1:
                f_siblings = df[df["부친"] == ancestor_chain[0].get("부친", "")].copy()
            gf_siblings = pd.DataFrame()
            if len(ancestor_chain) >= 2:
                gf_siblings = df[df["부친"] == ancestor_chain[1].get("부친", "")].copy()

            children = df[df["부친"] == p_name].copy()

            rankdir = "TB" if "수직" in tree_direction else "LR"
            lines = [
                'digraph GenealogyTree {',
                f'  rankdir="{rankdir}";',
                '  node [shape="box", style="filled,rounded", fontname="Noto Sans KR, sans-serif", fontsize="11"];',
                '  edge [color="#64748b", penwidth="1.3"];'
            ]

            all_tree_people = {}
            all_tree_people[p_name] = target_p
            lines.append(f'  "{p_name}" [label="★ {p_name}\\n({p_rel})", fillcolor="#FEF08A", color="#CA8A04", penwidth="2.5"];')

            prev = p_name
            for idx, a_row in enumerate(ancestor_chain):
                a_name = a_row.get("선조", "")
                a_rel = a_row.get("관계", "")
                all_tree_people[a_name] = a_row

                # 제2장 1항 ①: 증조부부터는 형제가 있는 경우 함자 외에 '0자 중 0자' 추가
                sibs = df[df["부친"] == a_row.get("부친", "")]
                label_extra = ""
                if idx >= 2 and len(sibs) > 1 and a_rel:
                    label_extra = f"\\n[{len(sibs)}자 중 {a_rel}]"

                rel_names = ["부(父)", "조부(祖父)", "증조부(曾祖)", "고조부(高祖)", "5대조", "6대조", "7대조", "8대조"]
                gen_title = rel_names[idx] if idx < len(rel_names) else f"{idx+1}대조"
                lines.append(f'  "{a_name}" [label="{gen_title}: {a_name}{label_extra}", fillcolor="#EDE9FE", color="#7C3AED"];')
                lines.append(f'  "{a_name}" -> "{prev}";')
                prev = a_name

            if sibling_rule != "직계 위주 표기":
                for _, sib in p_siblings.iterrows():
                    s_name = sib.get("선조", "")
                    s_rel = sib.get("관계", "")
                    if s_name != p_name:
                        all_tree_people[s_name] = sib
                        lines.append(f'  "{s_name}" [label="{s_name}\\n({s_rel})", fillcolor="#E0F2FE", color="#0284C7"];')
                        if len(ancestor_chain) >= 1:
                            lines.append(f'  "{ancestor_chain[0].get("선조", "")}" -> "{s_name}";')

                if len(ancestor_chain) >= 2 and not f_siblings.empty:
                    for _, sib in f_siblings.iterrows():
                        s_name = sib.get("선조", "")
                        s_rel = sib.get("관계", "")
                        if s_name != ancestor_chain[0].get("선조", ""):
                            all_tree_people[s_name] = sib
                            lines.append(f'  "{s_name}" [label="숙부: {s_name}\\n({s_rel})", fillcolor="#F1F5F9", color="#94A3B8"];')
                            lines.append(f'  "{ancestor_chain[1].get("선조", "")}" -> "{s_name}";')

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

            col_chart, col_person = st.columns([7, 5])
            with col_chart:
                st.markdown(f"##### 🌳 [{p_name}] 선조 계보도")
                st.graphviz_chart(chart_dot, use_container_width=True)

            with col_person:
                # 제2장 2항: 표시되는 해당 인물의 박스를 클릭했을 경우 그 인물의 약력을 표시한다.
                st.markdown("##### 📜 계보 인물 상세 약력")
                p_keys = list(all_tree_people.keys())
                sel_p_in_tree = st.selectbox(
                    "확인할 계보 인물을 선택하세요",
                    p_keys,
                    index=p_keys.index(p_name) if p_name in p_keys else 0
                )
                render_search_result_view(all_tree_people[sel_p_in_tree], title_prefix="계보 연동 약력")


# =============================================================================
# 제4장 3항: 완귀정 소개 글 (화면 하단 배치)
# =============================================================================
with st.expander("🏛️ 완귀정(玩龜亭) 유래", expanded=False):
    st.markdown("""
    <div class="wanguijeong-card">
        <h4>영천 완귀정 (경상북도 민속문화재 제20호)</h4>
        <ul>
            <li><b>건립자</b> : 완귀 안증(1494~1553)</li>
            <li><b>위치</b> : 경상북도 영천시 도남동</li>
            <li><b>유래</b> : 을사사화(1545)의 기미를 내다보고 벼슬에서 물러나 향리 호계(虎溪) 위에 정자를 짓고, 거북이 여섯 지체를 감추어 난을 피한다는 뜻을 취하여 '완귀(玩龜)'라 편액하였습니다.</li>
            <li><b>제향</b> : 매년 양력 4월 2째주 일요일 구암리 상로재(霜露齋)에서 봉행</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)


# =============================================================================
# 하단 전체 족보 데이터 목록 뷰
# =============================================================================
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
