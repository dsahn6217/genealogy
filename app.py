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

        # --- [직계 7대 탐색 알고리즘 (상위 선대 거슬러 올라가기 & 하위 자손)] ---
        # 1. 상위 직계 추적 (부친 연결고리를 따라 최대 선대 추적)
        ancestors = [] # [(이름, 관계호칭)]
        curr_row = selected_person_row
        while len(ancestors) < 5: # 상위 최대 5대까지 추적
            f_name = curr_row.get("부친", "").strip()
            if not f_name or f_name in ["미상", "없음", ""]:
                break
            ancestors.append(f_name)
            # 부친의 행 데이터 찾기
            f_rows = df[df["선조"] == f_name]
            if f_rows.empty:
                break
            curr_row = f_rows.iloc[0]

        # 2. 하위 직계 추적 (자녀 및 손자)
        children = df[df["부친"] == person_name]
        
        # --- [Graphviz DOT 언어 조립] ---
        dot_lines = [
            'digraph G {',
            '  rankdir="TB";',
            '  node [shape="box", style="filled,rounded", fontname="sans-serif", fontsize="11"];'
        ]

        # 본인 노드 (황금색 강조)
        p_label = f"★ {person_name}\\n({relation})" if relation else f"★ {person_name}"
        dot_lines.append(f'  "{person_name}" [label="{p_label}", fillcolor="#FFE0B2", color="#F57C00", penwidth="2.5"];')

        # 상위 선대 노드 및 연결선 생성 (선조 ➔ 선조)
        prev_node = person_name
        relation_labels = ["부(父)", "조부(祖父)", "증조부(曾祖)", "고조부(高祖)", "5대조", "6대조"]
        
        for idx, anc in enumerate(ancestors):
            lbl = relation_labels[idx] if idx < len(relation_labels) else f"{idx+1}대조"
            dot_lines.append(f'  "{anc}" [label="{lbl}: {anc}", fillcolor="#EDE7F6", color="#512DA8"];')
            dot_lines.append(f'  "{anc}" -> "{prev_node}";')
            prev_node = anc

        # 하위 자녀 노드 및 연결선 생성 (선택된 선조의 자녀 및 손자)
        if not children.empty:
            for _, c_row in children.iterrows():
                c_name = c_row.get("선조", "").strip()
                c_rel = c_row.get("관계", "").strip()
                c_lbl = f"자: {c_name}\\n({c_rel})" if c_rel else f"자: {c_name}"
                dot_lines.append(f'  "{c_name}" [label="{c_lbl}", fillcolor="#E8F5E9", color="#388E3C"];')
                dot_lines.append(f'  "{person_name}" -> "{c_name}";')

                # 손(孫) 탐색 (최대 1대 더 하위 탐색으로 7대 구성)
                grands = df[df["부친"] == c_name]
                if not grands.empty:
                    for _, g_row in grands.iterrows():
                        g_name = g_row.get("선조", "").strip()
                        g_rel = g_row.get("관계", "").strip()
                        g_lbl = f"손: {g_name}\\n({g_rel})" if g_rel else f"손: {g_name}"
                        dot_lines.append(f'  "{g_name}" [label="{g_lbl}", fillcolor="#F1F8E9", color="#689F38"];')
                        dot_lines.append(f'  "{c_name}" -> "{g_name}";')

        dot_lines.append('}')
        dot_code = "\n".join(dot_lines)

        # --- [화면 분할: 좌측 계보도(5) : 우측 상세 약력 표(7)] ---
        col_tree, col_profile = st.columns([5, 7])

        with col_tree:
            st.markdown("##### 🌳 직계 7대 계보도")
            st.graphviz_chart(dot_code, use_container_width=True)

        with col_profile:
            st.markdown("##### 📜 인물 상세 약력표")
            
            # 이해하기 쉬운 2단 상세 약력 테이블 데이터 생성
            k_prof = selected_person_row.get("한글 프로필 정렬", "") or selected_person_row.get("한글 프로필", "") or "기록 없음"
            c_prof = selected_person_row.get("한문 프로필 정렬", "") or selected_person_row.get("한문 프로필 원문", "") or "기록 없음"
            
            profile_table_html = f"""
            <table style="width:100%; border-collapse:collapse; font-size:14px; margin-top:5px; border:1px solid #dcdcdc;">
                <tr style="background-color:#f8f9fa;">
                    <th style="width:25%; padding:10px; border:1px solid #dcdcdc; text-align:center;">항목</th>
                    <th style="width:75%; padding:10px; border:1px solid #dcdcdc; text-align:left;">상세 내용</th>
                </tr>
                <tr>
                    <td style="padding:9px; border:1px solid #e6e6e6; text-align:center; font-weight:600; background-color:#fafafa;">고유 ID / 선조</td>
                    <td style="padding:9px; border:1px solid #e6e6e6;"><b>{person_gid}</b> | <span style="font-size:16px; color:#1a237e;"><b>{person_name}</b></span> ({relation})</td>
                </tr>
                <tr>
                    <td style="padding:9px; border:1px solid #e6e6e6; text-align:center; font-weight:600; background-color:#fafafa;">부친 / 조부</td>
                    <td style="padding:9px; border:1px solid #e6e6e6;">부: <b>{selected_person_row.get('부친', '')}</b> | 조부: <b>{selected_person_row.get('조부', '')}</b></td>
                </tr>
                <tr>
                    <td style="padding:10px; border:1px solid #e6e6e6; text-align:center; font-weight:600; background-color:#f0f7ff; color:#0d47a1;">한글 약력</td>
                    <td style="padding:10px; border:1px solid #e6e6e6; line-height:1.7; word-break:break-all; white-space:normal; background-color:#ffffff;">{k_prof}</td>
                </tr>
                <tr>
                    <td style="padding:10px; border:1px solid #e6e6e6; text-align:center; font-weight:600; background-color:#fbf8f3; color:#b78103;">한문 원문 약력</td>
                    <td style="padding:10px; border:1px solid #e6e6e6; line-height:1.7; word-break:break-all; white-space:normal; background-color:#ffffff; font-family:'Batang', serif;">{c_prof}</td>
                </tr>
            </table>
            """
            st.markdown(profile_table_html, unsafe_allow_html=True)

    st.markdown("---")

    # --- [검색 결과 전체 표 출력 화면] ---
    st.subheader(f"📋 검색 결과 목록 (총 {len(filtered_df)}명)")

    cols_order = [
        "고유ID", "선조", "관계", "부친", "조부",
        "한글 프로필 정렬", "한문 프로필 정렬", "한글 프로필", "한문 프로필 원문"
    ]
    display_cols = [c for c in cols_order if c in filtered_df.columns]
    display_df = filtered_df[display_cols]

    table_html = display_df.to_html(classes="custom-genealogy-table", index=False, escape=False)

    st.markdown(
        """
        <style>
        .custom-genealogy-table {
            width
