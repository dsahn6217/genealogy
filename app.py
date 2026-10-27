import pandas as pd
import streamlit as st

# 페이지 설정
st.set_page_config(
    page_title="완귀공파 족보 통합 검색 시스템", page_icon="📜", layout="wide"
)

st.title("📜 문중 족보 통합 검색 시스템")
st.markdown("전산화된 족보 데이터를 언제 어디서나 편리하게 조회하세요.")


# 데이터 로드 함수 (캐싱 적용)
@st.cache_data
def load_data():
  # 첫 번째 시트 '족보전사결과' 로드
  df = pd.read_excel("genealogy.xlsx", sheet_name="족보전사결과")
  # 모든 열을 문자열로 변환하고 공백 정리
  df = df.astype(str).apply(lambda x: x.str.strip())
  return df


# 데이터 불러오기 시도
try:
  df = load_data()

  # --- [사이드바 검색 조건 패널] ---
  st.sidebar.header("🔍 검색 조건")

  # 1. 성함 검색 (실제 엑셀 열 이름인 '선조' 기준)
  search_query = st.sidebar.text_input("성함 검색", "")

  # 2. 파(派) 선택 필터 (엑셀에 '파' 또는 관련 열이 있을 경우 활용)
  # 만약 파 분류 열이 따로 없다면 '선조' 기준으로 대체 가능하옵니다.
  if "파" in df.columns:
    pa_list = ["전체"] + sorted(df["파"].unique().tolist())
    selected_pa = st.sidebar.selectbox("파(派) 선택", pa_list)
  else:
    selected_pa = "전체"

  # --- [데이터 필터링 로직] ---
  filtered_df = df.copy()

  # 성함 검색어 적용
  if search_query:
    filtered_df = filtered_df[
        filtered_df["선조"].str.contains(search_query, na=False)
    ]

  # 파 선택 적용
  if selected_pa != "전체" and "파" in df.columns:
    filtered_df = filtered_df[filtered_df["파"] == selected_pa]

  # --- [검색 결과 출력 화면] ---
  st.subheader(f"검색 결과 (총 {len(filtered_df)}명)")

  # 표 형태로 데이터 출력
  st.dataframe(filtered_df, use_container_width=True)

except Exception as e:
  st.error(f"엑셀 파일('genealogy.xlsx')을 읽어오는 중 오류가 발생했습니다: {e}")
