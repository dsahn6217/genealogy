import streamlit as st
import pandas as pd
import base64
import os

# 0. 아이콘 이미지 읽기 (파일 존재 여부 확인 후 Base64 변환)
ICON_PATH = "icon.png"
icon_html = "📜 "

if os.path.exists(ICON_PATH):
    try:
        with open(ICON_PATH, "rb") as f:
            encoded = base64.b64encode(f.read()).decode()
            icon_html = f'<img src="data:image/png;base64,{encoded}" style="width: 52px; height: 52px; vertical-align: middle; margin-right: 10px; border-radius: 50%; box-shadow: 0 2px 6px rgba(0,0,0,0.2);">'
    except Exception:
        icon_html = "📜 "

# 1. 페이지 기본 설정
st.set_page_config(
    page_title="광주안씨 완귀공파 종중 족보 통합 검색 시스템",
    page_icon=ICON_PATH if os.path.exists(ICON_PATH) else "📜",
    layout="wide"
)

# 상단 헤더 및 본문 여백 조절 CSS
st.markdown("""
<style>
/* 본문 최상단 패딩을 0으로 만들어 Share 라인에 바짝 밀착 */
.block-container {
    padding-top: 0rem !important;
    padding-bottom: 2rem !important;
}

/* 상단 Share 툴바(헤더) 높이 및 투명도 유지 */
header[data-testid="stHeader"] {
    height: 2rem !important;
    background: transparent !important;
}
</style>
""", unsafe_allow_html=True)

# 2. 메인 타이틀 (아이콘을 두 문장의 세로 정중앙에 배치)
title_html = f"""
<div style="display: flex; justify-content: center; align-items: center; margin-top: 15px; margin-bottom: 20px;">
    <!-- 왼쪽: 원형 현판 아이콘 -->
    <div style="display: flex; align-items: center; justify-content: center; margin-right: 14px;">
        {icon_html}
    </div>
    <!-- 오른쪽: 두 줄 문장 묶음 (세로 정렬) -->
    <div style="display: flex; flex-direction: column; justify-content: center; text-align: left;">
        <div style="font-size: 26px; font-weight: 700; color: #1a237e; letter-spacing: -0.5px; line-height: 1.15; margin: 0; padding: 0;">
            광주안씨 완귀공파 종중
        </div>
        <div style="font-size: 20px; font-weight: 600; color: #212121; letter-spacing: -0.5px; line-height: 1.15; margin: 0; padding: 0; margin-top: 2px;">
            족보 통합 검색 시스템
        </div>
    </div>
</div>
"""
st.markdown(title_html, unsafe_allow_html=True)

# 3. 엑셀 데이터 로드 함수
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

# 4. 화면 중앙 상단 검색창 배치
with st.container():
    st.markdown("#### 🔍 선조 및 인물 검색")
    
    with st.form(key="main_search_form"):
        col_mode, col_pa = st.columns([2, 1])
        with col_mode:
            search_mode = st.radio(
                "검색 방식 선택",
                ["인명(선조) 검색", "직계(부친/조부) 검색", "통합 키워드 검색"],
                horizontal=True,
                index=0
            )
        with col_pa:
            selected_pa = "전체"
            if "파" in df.columns:
                pa_list = ["전체"] + sorted([p for p in df["파"].unique() if p and p != "nan"])
                selected_pa = st.selectbox("분파(派) 선택", pa_list)

        target_input = ""
        f_search = ""
        gf_search = ""

        if search_mode == "인명(선조) 검색":
            target_input = st.text_input("선조 성함 입력", placeholder="예: 효방, 효빈, 황원").strip()
        elif search_mode == "직계(부친/조부) 검색":
            c1, c2 = st.columns(2)
            with c1:
                f_search = st.text_input("부친 성함", placeholder="예: 부친 이름").strip()
            with c2:
                gf_search = st.text_input("조부 성함", placeholder="예: 조부 이름").strip()
            target_input = f"{f_search} {gf_search}".strip()
        elif search_mode == "통합 키워드 검색":
            target_input = st.text_input("통합 검색어 (지명, 묘소, 관직, 약력 등)", placeholder="예: 통정대부, 청도, 참판 등").strip()

        submit_btn = st.form_submit_button("🔍 검색 실행", use_container_width=True, type="primary")

# 5. 검색 필터링 로직
filtered_df = df.copy()

if selected_pa != "전체" and "파" in filtered_df.columns:
    filtered_df = filtered_df[filtered_df["파"] == selected_pa]

target_query = ""
if search_mode == "직계(부친/조부) 검색":
    target_query = f"{f_search} {gf_search}".strip()
else:
    target_query = target_input

# 6. 상세 계보 및 약력표 영역
if target_query:
    if search_mode == "인명(선조) 검색" and "선조" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["선조"].str.contains(target_query, na=False, case=False)]
    elif search_mode == "직계(부친/조부) 검색":
        if f_search and "부친" in filtered_df.columns:
            filtered_df = filtered_df[filtered_df["부친"].str.contains(f_search, na=False, case=False)]
        if gf_search and "조부" in filtered_df.columns:
            filtered_df = filtered_df[filtered_df["조부"].str.contains(gf_search, na=False, case=False)]
    elif search_mode == "통합 키워드 검색":
        mask = filtered_df.astype(str).apply(lambda col: col.str.contains(target_query, case=False, na=False)).any(axis=1)
        filtered_df = filtered_df[mask]

    if filtered_df.empty:
        st.warning(f"⚠️ '{target_query}' 조건으로 등록된 기록이 족보 데이터에 존재하지 않습니다.")
    else:
        st.markdown("---")
        
        # 단일 인물 또는 동명이인 판별
        if len(filtered_df) == 1:
            selected_person_row = filtered_df.iloc[0]
        else:
            msg = f"💡 검색 결과 총 {len(filtered_df)}명이 조회되었습니다. 계보도를 확인할 선조 한 분을 선택하십시오."
            st.info(msg)
            
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

        # 상위 5대 선대 추적
        ancestors = []
        curr_row = selected_person_row
        while len(ancestors) < 5:
            f_name = curr_row.get("부친", "").strip()
            if not f_name or f_name in ["미상", "없음", "", "nan"]:
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

        # 계보도와 약력표 나란히 배치
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

# 7. 검색 결과 전체 목록 표 출력
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
