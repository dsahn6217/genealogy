import streamlit as st
import pandas as pd

# 1. 페이지 설정
st.set_page_config(
    page_title="문중 족보 통합 검색 시스템",
    page_icon="📜",
    layout="wide"
)

st.title("📜 문중 족보 통합 검색 시스템")
st.markdown("전산화된 족보 데이터를 언제 어디서나 편리하게 조회하세요.")

# 2. 엑셀 데이터 로드 함수
@st.cache_data
def load_data():
    try:
        df = pd.read_excel("genealogy.xlsx")
        df = df.astype(str).apply(lambda x: x.str.strip())
        df = df.replace("nan", "")
        return df
    except Exception as e:
        st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
        return pd.DataFrame()

df = load_data()

if not df.empty:
    # 3. 사이드바 검색 조건
    st.sidebar.header("🔍 검색 조건")
    search_query = st.sidebar.text_input("성함 검색", placeholder="예: 세영, 지선").strip()

    if "파" in df.columns:
        pa_list = ["전체"] + sorted([p for p in df["파"].unique() if p])
        selected_pa = st.sidebar.selectbox("파 선택", pa_list)
    else:
        selected_pa = "전체"

    filtered_df = df.copy()

    # 필터 적용
    if search_query and "선조" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["선조"].str.contains(search_query, na=False)]

    if selected_pa != "전체" and "파" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["파"] == selected_pa]

    # 4. 인명 검색 및 1명 선택 영역
    if search_query and not filtered_df.empty:
        st.markdown("---")
        
        # 2명 이상일 때 선택 박스
        if len(filtered_df) > 1:
            st.info
