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

if df.empty:
    st.stop()

# 3. 사이드바 검색 조건
st.sidebar.header("🔍 검색 조건")
search_query = st.sidebar.text_input("성함 검색", placeholder="예: 효방, 효빈, 황원").strip()

selected_pa = "전체"
if "파" in df.columns:
    pa_list = ["전체"] + sorted([p for p in df["파"].unique() if p])
    selected_pa = st.sidebar.selectbox("파 선택", pa_list)

# 데이터 필터링
filtered_df = df.copy()

if search_query and "선조" in filtered_df.columns:
    filtered_df = filtered_df[filtered_df["선조"].str.contains(search_query, na=False, case=False)]

if selected_pa != "전체" and "파" in filtered_df.columns:
    filtered_df = filtered_df[filtered_df["파"] == selected_pa]

# 4. 인명 검색 결과 영역
if search_query:
    if filtered_df.empty:
        st.warning(f"⚠️ '{search_query}'(으)로 등록된 선조가 족보 데이터에 존재하지 않습니다.")
    else:
        st.markdown("---")
        
        # 조회 대상 인물 확정
        if len(filtered_df) == 1:
            selected_person_row = filtered_df.iloc[0]
        else:
            st.info(f"💡 '{search_query}' 검색 결과 총 {len(filtered_df)}명이 조회되었습니다. 열람할 선조 한 분을 선택하십시오.")
            options = []
            for _, r in filtered_df.iterrows():
                opt_gid = r.get("고유ID", "")
                opt_name = r.get("선조", "")
                opt_father = r.get("부친", "")
                opt_rel = r.get("관계", "")
                lbl = f"[{opt_gid}] {opt_name}"
                if opt_father:
                    lbl += f" (부: {opt_father}"
                    if opt_rel:
                        lbl += f", {opt_rel}"
                    lbl += ")"
                options.append(lbl)
            
            selected_label = st.selectbox("조회할 선조 선택", options)
            selected_gid = selected_label.split("]")[0].replace("[", "").strip()
            selected_person_row = filtered_df[filtered_df["고유ID"] == selected_gid].iloc[0]

        person_name = selected_person_row.get("선조", "").strip()
        person_gid = selected_person_row.get("고유ID", "").strip()
        relation = selected_person_row.get("관계", "").strip()

        st.subheader(f"🌳 [{person_name}] 선조 중심 7대 직계 계보 및 상세 약력")

        # 직계 선대 추적 (상위 5대)
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
            curr_row = f_rows.iloc[0]

        # 직계 자녀 탐색
        children = df[df["부친"] == person_name]

        # 계보도 DOT 문법 조립
        chart_lines = []
        chart_lines.append('digraph G {')
        chart_lines.append('  rankdir="TB";')
        chart_lines.append('  node [shape="box", style="filled,rounded", fontname="sans-serif", fontsize="11"];')

        # 본인 노드
        p_label = f"★ {person_name}\\n({relation})" if relation else f"★ {person_name}"
        chart_lines.append(f'  "{person_name}" [label="{p_label}", fillcolor="#FFE0B2", color="#F57C00", penwidth="2.5"];')

        # 선대 노드 연결
        prev_node = person_name
        rel_labels = ["부(父)", "조부(祖父)", "증조부(曾祖)", "고조부(高祖)", "5대조", "6대조"]
        for idx, anc in enumerate(ancestors):
            lbl = rel_labels[idx] if idx < len(rel_labels) else f"{idx+1}대조"
            chart_lines.append(f'  "{anc}" [label="{lbl}: {anc}", fillcolor="#EDE7F6", color="#512DA8"];')
            chart_lines.append(f'  "{anc}" -> "{prev_node}";')
            prev_node = anc

        # 자손 노드 연결
        if not children.empty:
            for _, c_row in children.iterrows():
                c_name = c_row.get("선조", "").strip()
                c_rel = c_row.get("관계", "").strip()
                c_lbl = f"자: {c_name}\\n({c_rel})" if c_rel else f"자: {c_name}"
                chart_lines.append(f'  "{c_name}" [label="{c_lbl}", fillcolor="#E8F5E9", color="#388E3C"];')
                chart_lines.append(f'  "{person_name}" -> "{c_name}";')

                grands = df[df["부친"] == c_name]
                if not grands.empty:
                    for _, g_row in grands.iterrows():
                        g_name = g_row.get("선조", "").strip()
                        g_rel = g_row.get("관계", "").strip()
                        g_lbl = f"손: {g_name}\\n({g_rel})" if g_rel else f"손: {g_name}"
                        chart_lines.append(f'  "{g_name}" [label="{g_lbl}", fillcolor="#F1F8E9", color="#689F38"];')
                        chart_lines.append(f'  "{c_name}" -> "{g_name}";')

        chart_lines.append('}')
        final_chart_code = "\n".join(chart_lines)

        # 좌측 계보도 : 우측 약력표 나란히 배치
        col_tree, col_profile = st.columns([5, 7])

        with col_tree:
            st.markdown("##### 🌳 직계 7대 계보도")
            st.graphviz_chart(final_chart_code, use_container_width=True)

        with col_profile:
            st.markdown("##### 📜 인물 상세 약력표")
            k_prof = selected_person_row.get("한글 프로필 정렬", "") or selected_person_row.get("한글 프로필", "") or "기록 없음"
            c_prof = selected_person_row.get("한문 프로필 정렬", "") or selected_person_row.get("한문 프로필 원문", "") or "기록 없음"
            f_name_val = selected_person_row.get("부친", "")
            gf_name_val = selected_person_row.get("조부", "")

            profile_table_html = (
                '<table style="width:100%; border-collapse:collapse; font-size:14px; margin-top:5px; border:1px solid #dcdcdc;">'
                '<tr style="background-color:#f8f9fa;">'
                '<th style="width:25%; padding:10px; border:1px solid #dcdcdc; text-align:center;">항목</th>'
                '<th style="width:75%; padding:10px; border:1px solid #dcdcdc; text-align:left;">상세 내용</th>'
                '</tr>'
                '<tr>'
                '<td style="padding:9px; border:1px solid #e6e6e6; text-align:center; font-weight:600; background-color:#fafafa;">고유 ID / 선조</td>'
                f'<td style="padding:9px; border:1px solid #e6e6e6;"><b>{person_gid}</b> | <span style="font-size:16px; color:#1a237e;"><b>{person_name}</b></span> ({relation})</td>'
                '</tr>'
                '<tr>'
                '<td style="padding:9px; border:1px solid #e6e6e6; text-align:center; font-weight:600; background-color:#fafafa;">부친 / 조부</td>'
                f'<td style="padding:9px; border:1px solid #e6e6e6;">부: <b>{f_name_val}</b> | 조부: <b>{gf_name_val}</b></td>'
                '</tr>'
                '<tr>'
                '<td style="padding:10px; border:1px solid #e6e6e6; text-align:center; font-weight:600; background-color:#f0f7ff; color:#0d47a1;">한글 약력</td>'
                f'<td style="padding:10px; border:1px solid #e6e6e6; line-height:1.7; word-break:break-all; white-space:normal; background-color:#ffffff;">{k_prof}</td>'
                '</tr>'
                '<tr>'
                '<td style="padding:10px; border:1px solid #e6e6e6; text-align:center; font-weight:600; background-color:#fbf8f3; color:#b78103;">한문 원문 약력</td>'
                f'<td style="padding:10px; border:1px solid #e6e6e6; line-height:1.7; word-break:break-all; white-space:normal; background-color:#ffffff; font-family:\'Batang\', serif;">{c_prof}</td>'
                '</tr>'
                '</table>'
            )
            st.markdown(profile_table_html, unsafe_allow_html=True)

st.markdown("---")

# 5. 검색 결과 전체 표 출력
st.subheader(f"📋 검색 결과 목록 (총 {len(filtered_df)}명)")

cols_order = [
    "고유ID", "선조", "관계", "부친", "조부",
    "한글 프로필 정렬", "한문 프로필 정렬", "한글 프로필", "한문 프로필 원문"
]
display_cols = [c for c in cols_order if c in filtered_df.columns]
display_df = filtered_df[display_cols]

table_html = display_df.to_html(classes="custom-genealogy-table", index=False, escape=False)

style_block = """
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
"""
st.markdown(style_block, unsafe_allow_html=True)
st.markdown(table_html, unsafe_allow_html=True)
