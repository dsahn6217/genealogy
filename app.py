import streamlit as st
import pandas as pd

# 1. 페이지 설정
st.set_page_config(
    page_title="문중 족보 통합 검색 시스템",
    page_icon="📜",
    layout="wide"
)

st.title("📜 문중 족보 통합 검색 시스템")
st.markdown("전산화된 족보 데이터를 언제 어디서나 편리하게 다각도로 조회하세요.")

# 2. 엑셀 데이터 로드 함수
@st.cache_data
def load_data():
    try:
        df = pd.read_excel("genealogy.xlsx")
        df = df.astype(str).apply(lambda x: x.str.strip())
        df = df.replace("nan", "")
        return df
    except Exception as e:
        st.error(f"엑셀 파일(genealogy.xlsx)을 읽는 중 오류가 발생했습니다: {e}")
        return pd.DataFrame()

df = load_data()

if df.empty:
    st.stop()

# 3. 사이드바: 다각적 검색 조건
st.sidebar.header("🔍 검색 조건")

search_mode = st.sidebar.radio(
    "검색 방식 선택",
    ["인명(선조) 검색", "직계(부친/조부) 검색", "통합 키워드 검색"],
    index=0
)

selected_pa = "전체"
if "파" in df.columns:
    pa_list = ["전체"] + sorted([p for p in df["파"].unique() if p and p != "nan"])
    selected_pa = st.sidebar.selectbox("분파(派) 선택", pa_list)

filtered_df = df.copy()

if selected_pa != "전체" and "파" in filtered_df.columns:
    filtered_df = filtered_df[filtered_df["파"] == selected_pa]

target_query = ""

if search_mode == "인명(선조) 검색":
    target_query = st.sidebar.text_input("선조 성함 입력", placeholder="예: 효방, 효빈, 황원").strip()
    if target_query and "선조" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["선조"].str.contains(target_query, na=False, case=False)]

elif search_mode == "직계(부친/조부) 검색":
    f_search = st.sidebar.text_input("부친 성함", placeholder="예: 부친 이름 입력").strip()
    gf_search = st.sidebar.text_input("조부 성함", placeholder="예: 조부 이름 입력").strip()
    
    if f_search and "부친" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["부친"].str.contains(f_search, na=False, case=False)]
    if gf_search and "조부" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["조부"].str.contains(gf_search, na=False, case=False)]
    target_query = f"{f_search} {gf_search}".strip()

elif search_mode == "통합 키워드 검색":
    target_query = st.sidebar.text_input("통합 검색어 (지명, 묘소, 관직, 약력 등)", placeholder="예: 청도, 참판, 통정대부 등").strip()
    if target_query:
        mask = filtered_df.astype(str).apply(lambda col: col.str.contains(target_query, case=False, na=False)).any(axis=1)
        filtered_df = filtered_df[mask]

# 4. 상세 계보 및 약력표 영역
if target_query:
    if filtered_df.empty:
        st
