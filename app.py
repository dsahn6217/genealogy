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
