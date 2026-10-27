import streamlit as st
import pandas as pd

# 1. 페이지 기본 환경 설정 (와이드 레이아웃 적용)
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
    df = df.astype(str).apply(lambda x: x.str.strip())
    df = df.replace("nan", "")
    return df

# 3. 메인 실행 및 필터링 로직
try:
    df = load_data()

    # --- [사이드바 검색 조건] ---
    st.sidebar.header("🔍 검색 조건")
    search_query = st.sidebar.text_input("성함 검색", placeholder="예: 세영, 지선").strip()

    if "파" in df.columns:
        pa_list = ["전체"] + sorted([p for p in df["파"].unique() if p])
        selected_pa = st.sidebar.selectbox("파 선택", pa_list)
    else:
        selected_pa = "전체"

    filtered_df = df.copy()

    # 성함 검색 필터
    if search_query:
        if "선조" in filtered_df.columns:
            filtered_df = filtered_df[filtered_df["선조"].str.contains(search_query, na=False)]

    # 파 필터
    if selected_pa != "전체" and "파" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["파"] == selected_pa]

    # --- [인명 검색 및 1명 선택 영역] ---
    if search_query and not filtered_df.empty:
        st.markdown("---")
        
        # 검색 결과가 2명 이상일 경우 선택 상자 제공
        if len(filtered_df) > 1:
            st.info(f"💡 '{search_query}' 검색 결과 총 {len(filtered_df)}명이 조회되었습니다. 열람할 선조 한 분을 선택하십시오.")
            
            def make_label(row):
                gid = row.get("고유ID", "")
                name = row.get("선조", "")
                father = row.get("부친", "")
                rel = row.get("관계", "")
                label = f"[{gid}] {name}"
                if father:
                    label += f" (부: {father}"
                    if rel:
                        label += f", {rel}"
                    label += ")"
                return label

            options = [make_label(row) for _, row in filtered_df.iterrows()]
            selected_label = st.selectbox("조회할 선조 선택", options)
            selected_gid = selected_label.split("]")[0].replace("[", "").strip()
            selected_person_row = filtered_df[filtered_df["고유ID"] == selected_gid].iloc[0]
        else:
            selected_person_row = filtered_df.iloc[0]

        person_name = selected_person_row.get("선조", "").strip()
        person_gid = selected_person_row.get("고유ID", "").strip()
        relation = selected_person_row.get("관계", "").strip()

        st.subheader(f"🌳 [{person_name}] 선조 중심 7대 직계 계보 및 상세 약력")

        # --- [직계 7대 탐색 알고리즘] ---
        # 1. 상위 직계 추적
        ancestors = []
        curr_row = selected_person_row
        while len(ancestors) < 5:
            f_name = curr_row.get("부친", "").strip()
            if not f_name or f_name in ["미상", "없음", ""]:
                break
            ancestors.append(f_name)
            f_rows = df[df["선조"] == f_name]
            if f_rows.empty:
                break
            curr_row = f
