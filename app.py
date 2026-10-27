import streamlit as st
import pandas as pd

# 1. 페이지 기본 환경 설정
st.set_page_config(
    page_title="문중 족보 통합 검색 시스템",
    page_icon="📜",
    layout="wide"
)

st.title("📜 문중 족보 통합 검색 시스템")
st.markdown("전산화된 족보 데이터를 언제 어디서나 편리하게 조회하세요.")

# 2. 엑셀 데이터 로드 함수 (캐시 적용)
@st.cache_data
def load_data():
    excel_file = "genealogy.xlsx"
    df = pd.read_excel(excel_file)
    # 모든 열의 데이터를 문자열로 변환하고 앞뒤 공백 제거
    df = df.astype(str).apply(lambda x: x.str.strip())
    # 결측치('nan') 문자열 정리
    df = df.replace("nan", "")
    return df

# 3. 메인 실행 및 필터링 로직
try:
    df = load_data()

    # --- [사이드바 검색 조건] ---
    st.sidebar.header("🔍 검색 조건")
    
    # 성함 검색어 입력창
    search_query = st.sidebar.text_input("성함 검색", placeholder="예: 세영, 지선").strip()

    # 파 필터 (엑셀에 '파' 열이 있을 경우 선택 상자 제공)
    if "파" in df.columns:
        pa_list = ["전체"] + sorted([p for p in df["파"].unique() if p])
        selected_pa = st.sidebar.selectbox("파 선택", pa_list)
    else:
        selected_pa = "전체"

    # --- [데이터 필터링] ---
    filtered_df = df.copy()

    # 성함 검색 필터
    if search_query:
        if "선조" in filtered_df.columns:
            filtered_df = filtered_df[filtered_df["선조"].str.contains(search_query, na=False)]

    # 파 필터
    if selected_pa != "전체" and "파" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["파"] == selected_pa]

    # --- [검색 결과 출력 화면] ---
    st.subheader(f"검색 결과 (총 {len(filtered_df)}명)")

    # 주군께서 명하신 열 순서 재배치
    cols_order = [
        "고유ID", "선조", "관계", "부친", "조부",
        "한글 프로필 정렬", "한문 프로필 정렬", "한글 프로필", "한문 프로필 원문"
    ]
    # 실제 엑셀 파일에 존재하는 열만 안전하게 선택
    display_cols = [c for c in cols_order if c in filtered_df.columns]

    # 프로필 내용이 시원하게 보이도록 TextColumn 및 너비 확장 적용
    st.dataframe(
        filtered_df[display_cols],
        use_container_width=True,
        hide_index=True,
        column_config={
            "한글 프로필 정렬": st.column_config.TextColumn("한글 프로필 정렬", width="large"),
            "한문 프로필 정렬": st.column_config.TextColumn("한문 프로필 정렬", width="large"),
            "한글 프로필": st.column_config.TextColumn("한글 프로필", width="large"),
            "한문 프로필 원문": st.column_config.TextColumn("한문 프로필 원문", width="large"),
        }
    )

except Exception as e:
    st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
