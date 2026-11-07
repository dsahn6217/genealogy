import base64
import os
import re
import sqlite3
import openpyxl
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# ==============================================================================
# 0. 기본 설정 및 아이콘 / 데이터 / 배경 / PDF 파일 경로
# ==============================================================================
ICON_FILE = "icon.png"
DB_FILE = "genealogy.db"
DATA_FILE = "genealogy.xlsx"
PDF_HOYEONSA = "호연사(서원) 소개문.pdf"
PDF_CHEONGJE = "영천 청제 및 청제비 소개문.pdf"

# 깃허브 저장소 이미지 URL
HOYEONSA_IMG_URL = "https://github.com/dsahn6217/genealogy/blob/main/%ED%98%B8%EC%97%B0%EC%82%AC%20%EC%98%9B%20%EC%82%AC%EC%A7%84.png?raw=true"

# 저장소 업로드 배경 이미지 파일 탐색
BG_PATHS = [
    "background.png",
    "background.jpg",
    "background.jpeg",
    "bg.png",
    "bg.jpg",
]

# ⭐ st.set_page_config는 반드시 모든 st.* 명령보다 가장 먼저 호출되어야 함
st.set_page_config(
    page_title="광주안씨 완귀공파 종중 족보 검색 시스템",
    page_icon=ICON_FILE if os.path.exists(ICON_FILE) else "📜",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ⭐ 이미지 Base64 인코딩을 @st.cache_resource로 메모리에 캐싱하여 부팅 시 디스크 재읽기 제거
@st.cache_resource(show_spinner=False)
def get_base64_image(image_path):
  if image_path and os.path.exists(image_path):
    try:
      with open(image_path, "rb") as f:
        return (
            f"data:image/png;base64,{base64.b64encode(f.read()).decode('utf-8')}"
        )
    except Exception:
      return None
  return None


# ⭐ PDF 파일을 Base64로 인코딩하여 iframe 뷰어 생성하는 헬퍼 함수
def render_pdf_viewer(pdf_path, download_label, file_name):
  if os.path.exists(pdf_path):
    with open(pdf_path, "rb") as f:
      pdf_bytes = f.read()

    col1, col2 = st.columns([1, 4])
    with col1:
      st.download_button(
          label=download_label,
          data=pdf_bytes,
          file_name=file_name,
          mime="application/pdf",
          use_container_width=True,
      )

    base64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
    pdf_display = f"""
            <iframe 
                src="data:application/pdf;base64,{base64_pdf}" 
                width="100%" 
                height="800" 
                type="application/pdf"
                style="border: 1px solid #ced4da; border-radius: 8px; margin-top: 10px;">
            </iframe>
        """
    st.markdown(pdf_display, unsafe_allow_html=True)
  else:
    st.info(
        f"💡 안내: 현재 저장소에 `{pdf_path}` 파일이 준비되어 있지 않습니다."
    )


# ⭐ 데이터베이스 로드 및 캐싱
@st.cache_data(show_spinner=False)
def load_genealogy_data():
  if os.path.exists(DB_FILE):
    try:
      conn = sqlite3.connect(DB_FILE)
      df = pd.read_sql_query("SELECT * FROM genealogy", conn)
      conn.close()
      return df
    except Exception:
      pass

  if os.path.exists(DATA_FILE):
    try:
      df = pd.read_excel(DATA_FILE)
      return df
    except Exception:
      pass

  return pd.DataFrame()


# ==============================================================================
# 메인 화면 구성
# ==============================================================================
def main():
  # 메인 타이틀 배너
  st.title("📜 광주안씨 완귀공파 종중 대동보(전자족보)")
  st.caption(
      "Gwangju Ahn Clan (Wan-gwi-gong Branch) Genealogy & Cultural Heritage"
      " System"
  )
  st.write("---")

  # 상단 탭 메뉴 구성: 유적지 소개문을 첫 번째 탭으로 배치하여 첫 화면 노출
  tab1, tab2, tab3 = st.tabs(
      ["🏛️ 문중 문화 유적지", "🔍 선조·종친 인물 검색", "🌳 계보도(분파도)"]
  )

  # --------------------------------------------------------------------------
  # 탭 1: 문중 문화 유적지 소개 (첫 화면)
  # --------------------------------------------------------------------------
  with tab1:
    st.subheader("호연사(虎淵祠) 및 호연서원 유허비각")

    # 깃허브 실물 사진 표시
    st.image(
        HOYEONSA_IMG_URL,
        caption="영천 호연서원 유허비각(遺墟碑閣) 전경",
        use_container_width=True,
    )

    st.markdown("""
        > **광주 안씨 완귀공파 종중 문화 유적지**  
        > 영천 호연사(서원)는 완귀 안증(安潧) 선생의 도학과 덕업을 기리기 위해 숙종 12년(1686) 사림에서 창건하여 위패를 모셨던 유서 깊은 유적지입니다.
        """)

    # 호연사 PDF 뷰어 및 다운로드 단추
    render_pdf_viewer(
        pdf_path=PDF_HOYEONSA,
        download_label="📥 호연사 소개문(PDF) 인쇄·다운로드",
        file_name="호연사(서원)_소개문.pdf",
    )

    st.write("---")
    st.subheader("영천 청제(菁堤) 및 청제비(靑堤碑)")
    render_pdf_viewer(
        pdf_path=PDF_CHEONGJE,
        download_label="📥 청제 및 청제비 소개문(PDF) 인쇄·다운로드",
        file_name="영천_청제_소개문.pdf",
    )

  # --------------------------------------------------------------------------
  # 탭 2: 선조·종친 인물 검색
  # --------------------------------------------------------------------------
  with tab2:
    st.subheader("종친 및 조상님 인명 검색")
    df = load_genealogy_data()

    if not df.empty:
      search_col1, search_col2 = st.columns([2, 1])
      with search_col1:
        search_query = st.text_input(
            "성함(이름) 또는 자(字), 호(號)를 입력하십시오", ""
        )
      with search_col2:
        search_btn = st.button("검색 실행", use_container_width=True)

      if search_query:
        # 데이터프레임 내 문자열 검색
        mask = df.astype(str).apply(
            lambda row: row.str.contains(search_query, regex=False).any(),
            axis=1,
        )
        result_df = df[mask]

        if not result_df.empty:
          st.success(f"총 {len(result_df)}건의 기록이 검색되었습니다.")
          st.dataframe(result_df, use_container_width=True)
        else:
          st.warning(f"'{search_query}'에 대한 검색 결과가 없습니다.")
      else:
        st.dataframe(df.head(50), use_container_width=True)
        st.caption("※ 상위 50건의 수록 인명이 기본 표시됩니다.")
    else:
      st.warning(
          f"데이터 파일(`{DATA_FILE}` 또는 `{DB_FILE}`)을 불러올 수 없습니다."
      )

  # --------------------------------------------------------------------------
  # 탭 3: 계보도(분파도) 보기
  # --------------------------------------------------------------------------
  with tab3:
    st.subheader("완귀공파 분파도 및 계보도")
    hangul_diagram = "branch_diagram_hangul.PNG"
    hanmoon_diagram = "branch_diagram_hanmoon.PNG"

    view_mode = st.radio(
        "계보도 표기 선택", ["한글 계보도", "한문 계보도"], horizontal=True
    )

    if view_mode == "한글 계보도" and os.path.exists(hangul_diagram):
      st.image(
          hangul_diagram,
          caption="광주안씨 완귀공파 한글 분파도",
          use_container_width=True,
      )
    elif view_mode == "한문 계보도" and os.path.exists(hanmoon_diagram):
      st.image(
          hanmoon_diagram,
          caption="광주안씨 완귀공파 한문 분파도",
          use_container_width=True,
      )
    else:
      st.info(
          "계보도 이미지 파일(`branch_diagram_hangul.PNG` /"
          " `branch_diagram_hanmoon.PNG`)이 저장소에 등록되어 있습니다."
      )


if __name__ == "__main__":
  main()
