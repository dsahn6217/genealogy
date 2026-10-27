import streamlit as st
import pandas as pd
import graphviz

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

    # --- [계보도 시각화 영역 (인물 검색 시 자동 표시)] ---
    if search_query and not filtered_df.empty:
        st.markdown("---")
        st.subheader(f"🌳 '{search_query}' 관련 직계 계보도")
        
        # 검색된 인물들 중 최대 3명에 대해 각각 계보도 트리 생성
        target_people = filtered_df["선조"].unique()[:3]
        
        for person in target_people:
            dot = graphviz.Digraph(comment=f"{person} 계보도")
            dot.attr(rankdir="TB", size="8,5")
            dot.attr("node", shape="box", style="filled,rounded", fontname="Malgun Gothic", fontsize="11")

            # 해당 인물 레코드 추출
            person_rows = df[df["선조"] == person]
            if person_rows.empty:
                continue
            row = person_rows.iloc[0]

            father = row.get("부친", "").strip()
            grandfather = row.get("조부", "").strip()
            relation = row.get("관계", "").strip()

            # 1. 본인 노드 (강조 색상)
            dot.node(person, f"★ {person}\n({relation})", fillcolor="#FFE0B2", color="#F57C00", penwidth="2")

            # 2. 직계 상위 (조부 ➔ 부친 ➔ 본인)
            if father:
                dot.node(father, f"부: {father}", fillcolor="#E1F5FE", color="#0288D1")
                dot.edge(father, person)

                if grandfather:
                    dot.node(grandfather, f"조부: {grandfather}", fillcolor="#EDE7F6", color="#512DA8")
                    dot.edge(grandfather, father)

            # 3. 직계 하위 (자녀 탐색: 부친 이름이 본인인 인물)
            children = df[df["부친"] == person]
            if not children.empty:
                for _, child_row in children.iterrows():
                    child_name = child_row.get("선조", "").strip()
                    child_rel = child_row.get("관계", "").strip()
                    child_label = f"{child_name}\n({child_rel})" if child_rel else child_name
                    dot.node(child_name, child_label, fillcolor="#E8F5E9", color="#388E3C")
                    dot.edge(person, child_name)

            st.caption(f"**[{person}] 선조 중심 직계 계보**")
            st.graphviz_chart(dot)

    st.markdown("---")

    # --- [검색 결과 표 출력 화면] ---
    st.subheader(f"검색 결과 (총 {len(filtered_df)}명)")

    cols_order = [
        "고유ID", "선조", "관계", "부친", "조부",
        "한글 프로필 정렬", "한문 프로필 정렬", "한글 프로필", "한문 프로필 원문"
    ]
    display_cols = [c for c in cols_order if c in filtered_df.columns]
    display_df = filtered_df[display_cols]

    # 내용 전체가 아래로 줄바꿈되어 보이도록 스타일 적용
    table_html = display_df.to_html(classes="custom-genealogy-table", index=False, escape=False)

    st.markdown(
        """
        <style>
        .custom-genealogy-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
            margin-top: 10px;
        }
        .custom-genealogy-table th {
            background-color: #f0f2f6;
            color: #31333F;
            font-weight: 600;
            text-align: center;
            padding: 10px 8px;
            border: 1px solid #dcdcdc;
            white-space: nowrap;
        }
        .custom-genealogy-table td {
            padding: 10px 8px;
            border: 1px solid #e6e6e6;
            vertical-align: top;
            line-height: 1.6;
            word-break: break-all;
            white-space: normal !important;
        }
        .custom-genealogy-table td:nth-child(1),
        .custom-genealogy-table td:nth-child(2),
        .custom-genealogy-table td:nth-child(3),
        .custom-genealogy-table td:nth-child(4),
        .custom-genealogy-table td:nth-child(5) {
            text-align: center;
            white-space: nowrap !important;
        }
        .custom-genealogy-table td:nth-child(n+6) {
            text-align: left;
            min-width: 260px;
        }
        .custom-genealogy-table tr:hover {
            background-color: #f9fbfd;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown(table_html, unsafe_allow_html=True)

except Exception as e:
    st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
