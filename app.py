import pandas as pd
import streamlit as st

# 페이지 설정
st.set_page_config(
    page_title="문중 족보 통합 조회 시스템", page_icon="📜", layout="wide"
)

# [보안] 간단한 비밀번호 설정 (문중원 공용 비밀번호)
PASSWORD = (  # 추후 변경 가능
    "wangui1494"
)


def check_password():
  """비밀번호 인증 함수"""
  if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

  if st.session_state["authenticated"]:
    return True

  st.title("🔒 문중 족보 보안 로그인")
  st.markdown("본 족보는 문중원 전용 공간입니다. 공용 비밀번호를 입력해주세요.")
  input_pw = st.text_input("비밀번호", type="password")

  if st.button("로그인"):
    if input_pw == PASSWORD:
      st.session_state["authenticated"] = True
      st.rerun()
    else:
      st.error("비밀번호가 올바르지 않습니다.")
  return False


if check_password():
  # 로그인 성공 시 나타나는 메인 화면
  st.title("📜 문중 족보 통합 검색 시스템")
  st.markdown("전산화된 족보 데이터를 언제 어디서나 편리하게 조회하세요.")


  # 엑셀 데이터 로드
  @st.cache_data
  def load_data():
    df = pd.read_excel("genealogy.xlsx")
    return df.astype(str)


  try:
    df = load_data()

    # 사이드바 검색 필터
    st.sidebar.header("🔍 검색 조건")
    search_query = st.sidebar.text_input("성함 검색", "")
    selected_派 = st.sidebar.selectbox(
        "파(派) 선택", ["전체"] + list(df["파수"].unique())
    )

    # 데이터 필터링
    filtered_df = df.copy()
    if search_query:
      filtered_df = filtered_df[
          filtered_df["이름"].str.contains(search_query, na=False)
      ]
    if selected_派 != "전체":
      filtered_df = filtered_df[filtered_df["파수"] == selected_派]

    # 결과 출력
    st.subheader(f"검색 결과 (총 {len(filtered_df)}명)")
    st.dataframe(filtered_df, use_container_width=True)

    # 상세 조회
    st.markdown("---")
    st.subheader("👤 인물 상세 정보")
    if not filtered_df.empty:
      selected_name = st.selectbox(
          "상세 조회할 인물을 선택하세요", filtered_df["이름"].tolist()
      )
      if selected_name:
        person_info = df[df["이름"] == selected_name].iloc[0]
        col1, col2 = st.columns(2)
        with col1:
          st.metric(label="성함", value=person_info["이름"])
          st.write(f"**파수:** {person_info.get('파수', '')}")
          st.write(f"**항렬:** {person_info.get('항렬', '')}")
        with col2:
          st.write(f"**배우자:** {person_info.get('배우자', '')}")
          st.write(f"**기타/거주지:** {person_info.get('거주지/비고', '')}")
    else:
      st.warning("검색 결과가 없습니다.")

  except Exception as e:
    st.error(
        f"엑셀 파일('genealogy.xlsx')을 읽어오는 중 오류가 발생했습니다: {e}"
    )
