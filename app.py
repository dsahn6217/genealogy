import os
import re
import difflib
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

f1_path = "/working_dir/c_4447873e205b8969/user_uploaded_files/ttl=63d/user_uploaded/보정 족보A_B_C_D_통합_전사결과.xlsx"
f2_path = "/working_dir/c_4447873e205b8969/user_uploaded_files/ttl=63d/user_uploaded/완귀공파 족보 통합.xlsx"
out_dir = "/working_dir/c_4447873e205b8969/artifacts/file_generation/ttl=63d/output"
os.makedirs(out_dir, exist_ok=True)
out_path = os.path.join(out_dir, "두파일_전사결과_시트_차이비교.xlsx")

wb1 = openpyxl.load_workbook(f1_path, data_only=True)
wb2 = openpyxl.load_workbook(f2_path, data_only=True)

ws1 = wb1["족보전사결과"]
ws2 = wb2["족보전사결과"]

rows1 = list(ws1.iter_rows(values_only=True))
rows2 = list(ws2.iter_rows(values_only=True))
headers = rows1[0]


def normalize_fmt(s):
    if not s:
        return ""
    s = s.replace("정월", "1월").replace("正월", "1월").replace("년", "年")
    s = s.replace("〇〇", "00")
    rep = [
        ("三十一", "31"), ("三十", "30"),
        ("二十九", "29"), ("二十八", "28"), ("二十七", "27"), ("二十六", "26"), ("二十五", "25"),
        ("二十四", "24"), ("二十三", "23"), ("二十二", "22"), ("二十一", "21"), ("二十", "20"),
        ("十九", "19"), ("十八", "18"), ("十七", "17"), ("十六", "16"), ("十五", "15"),
        ("十四", "14"), ("十三", "13"), ("十二", "12"), ("十一", "11"), ("十", "10"),
    ]
    for k, v in rep:
        s = s.replace(k, v)
    s = s.translate(str.maketrans("一二三四五六七八九〇", "1234567890"))
    s = re.sub(r"[\s,]+", "", s)
    return s


def summarize_str_diff(s1, s2):
    s1 = str(s1 or "")
    s2 = str(s2 or "")
    if s1 == s2:
        return "동일"
    if re.sub(r"\s+", "", s1) == re.sub(r"\s+", "", s2):
        matcher = difflib.SequenceMatcher(None, s1, s2)
        parts = []
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag != "equal":
                c1 = s1[max(0, i1 - 3):min(len(s1), i2 + 3)]
                c2 = s2[max(0, j1 - 3):min(len(s2), j2 + 3)]
                parts.append(f"띄어쓰기 차이 ('{c1}' ↔ '{c2}')")
        return "; ".join(parts[:3]) if parts else "띄어쓰기(공백) 차이"

    matcher = difflib.SequenceMatcher(None, s1, s2)
    changes = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag != "equal":
            seg1 = s1[i1:i2]
            seg2 = s2[j1:j2]
            ctx1 = s1[max(0, i1 - 4):min(len(s1), i2 + 4)]
            ctx2 = s2[max(0, j1 - 4):min(len(s2), j2 + 4)]
            if len(seg1) > 25 or len(seg2) > 25:
                changes.append(f"'{seg1[:25]}...' ↔ '{seg2[:25]}...'")
            else:
                changes.append(f"'{ctx1}'({seg1}) ↔ '{ctx2}'({seg2})")
    return " / ".join(changes[:4])


def classify_row_diff(r_num, uid, name, rel, diff_cols):
    if r_num in (117, 222):
        return "① 약력 정렬 내용 전면 상이", "심각(내용불일치)"
    if r_num in (695, 711, 764, 898, 1169):
        return "② 계자(系子) 생부 인명·독음 상이", "중요(인명불일치)"
    if r_num in (357, 781, 794, 960, 1005, 1127, 2068):
        return "③ 한문 약력 정렬 내 한글 토씨(조사) 유무 차이", "경미(조사유무)"
    if r_num in (44, 90, 147, 158, 172):
        return "④ 한글 약력 원문 내 한자('正월'·'年') 표기 차이", "경미(한자표기)"
    if "녀" not in str(rel or ""):
        return "⑤ 남자 항목 한문 약력 정렬 단순 띄어쓰기 차이", "서식(띄어쓰기)"
    return "⑥ 녀(女) 항목 약력 정렬 숫자(한자↔아라비아)·띄어쓰기 차이", "서식(숫자/공백)"


specific_notes = {
    117: "파일1은 타 인물(字 聖任, 英祖辛未生...)의 약력이 잘못 정렬됨 / 파일2는 원문과 일치하는 경익(字 而久, 英祖庚午一七四九年生...) 약력으로 정렬됨",
    222: "파일1은 정렬 열에 '出系 養父 景贇 系子'('양부 경빈의 계자')만 축약 기재됨 / 파일2는 원문 전체 약력(字 致華, 正祖戊申一七八八年生... 배위 월성손씨 열행 기록 포함)이 정상 정렬됨",
    695: "한문·한글 약력 원문에서 생부 표기가 파일1은 '璆遠의子'('구원의 아들'), 파일2는 '璡遠의子'('진원의 아들')로 상이함 (정렬 열은 두 파일 모두 '璡遠/진원'으로 동일)",
    711: "한문·한글 약력 정렬에서 생부 표기가 파일1은 '生父 相遠 一子'('생부 상원의 1자'), 파일2는 '生父 必遠 一子'('생부 필원의 1자')로 상이함 (원문 열은 두 파일 모두 '相遠의子/상원의 아들')",
    764: "한문·한글 약력 정렬에서 생부 표기가 파일1은 '生父 鍾祿 一子'('생부 종록의 1자') 및 '墓 서울市', 파일2는 '生父 鍾百 一子'('생부 종백의 1자') 및 '墓 는서울市'로 상이함 (원문은 '鍾祿의子')",
    898: "4개 약력 열 모두 생부 표기 상이: 파일1은 원문·정렬 모두 '彥壽'('언수') / 파일2는 원문 '彥燾'('언도'), 정렬 '彥鐸'('언탁')으로 기재됨",
    1169: "한문 약력 정렬에서 생부 한자가 파일1은 '生父 泓洙 二子', 파일2는 '生父 弘洙 二子'로 상이함 (한문 원문은 두 파일 모두 '泓洙의二子')",
    357: "한문 약력 정렬에서 파일1은 한글 토씨 제거('齊曰枕流 於焉棲息'), 파일2는 한글 토씨 잔존('齊曰枕流라하여於焉棲息')",
    781: "한문 약력 정렬에서 파일1은 '墓 경기도안산시...', 파일2는 조사 '는' 잔존('墓 는경기도안산시...')",
    794: "한문 약력 정렬에서 파일1은 '墓 서울基督教...', 파일2는 조사 '는' 잔존('墓 는서울基督教...')",
    960: "한문 약력 정렬에서 파일1은 조사 제거('好學生計 晩年 士林...'), 파일2는 조사 잔존('好學生計에 晩年에야士林...')",
    1005: "한문 약력 정렬에서 파일1은 '墓 요시동골서남間', 파일2는 조사 '는' 잔존('墓 는요시동골서남間')",
    1127: "한문 약력 정렬에서 파일1은 '墓 경기고양시일산...', 파일2는 조사 '는' 잔존('墓 는경기고양시일산...')",
    2068: "한문 약력 정렬에서 파일1은 '新舊學兼備 以孝善行...', 파일2는 어미 '며' 잔존('新舊學兼備며以孝善行...')",
    44: "한글 약력 원문에서 파일1은 한글 '정월 십삼일 졸', 파일2는 한자 혼용 '正월 십삼일 졸'",
    90: "한글 약력 원문에서 파일1은 한글 '정월 十二일 졸', 파일2는 한자 혼용 '正월 十二일 졸'",
    147: "한글 약력 원문에서 파일1은 한글 '병신 一八二六년', 파일2는 한자 혼용 '병신 一八二六年'",
    158: "한글 약력 원문에서 파일1은 한글 '정월 二十일졸'·'정월 十八일졸', 파일2는 한자 혼용 '正월 二十일졸'·'正월 十八일졸'",
    172: "한글 약력 원문에서 파일1은 한글 '정월 十六일졸', 파일2는 한자 혼용 '正월 十六일졸'",
    173: "한문 약력 정렬 내 띄어쓰기 차이: 파일1 '墓 后'(2회) ↔ 파일2 '墓后'(2회)",
    187: "한문 약력 정렬 내 띄어쓰기 차이: 파일1 '墓 后' ↔ 파일2 '墓后'",
    221: "한문 약력 정렬 내 띄어쓰기 차이: 파일1 '兩配 位' ↔ 파일2 '兩配位'",
    372: "한문 약력 정렬 내 띄어쓰기 차이: 파일1 '墓 后' ↔ 파일2 '墓后'",
    683: "한문 약력 정렬 내 띄어쓰기 차이: 파일1 '公墓 同麓坐' ↔ 파일2 '公墓同麓坐'",
    893: "한문 약력 정렬 내 띄어쓰기 차이: 파일1 '墓 后' ↔ 파일2 '墓后'",
    1478: "한문 약력 정렬 내 띄어쓰기 차이: 파일1 '墓 同園' ↔ 파일2 '墓同園'",
}

all_diff_rows = []
cell_diff_records = []

for r_idx in range(1, len(rows1)):
    r1 = rows1[r_idx]
    r2 = rows2[r_idx]
    if r1 != r2:
        r_num = r_idx + 1
        uid, name, rel, father, grandpa = r1[0], r1[1], r1[2], r1[3], r1[4]
        diff_cols = []
        for c_idx, h in enumerate(headers):
            v1 = r1[c_idx]
            v2 = r2[c_idx]
            if v1 != v2:
                diff_cols.append((h, v1, v2))
        cat_label, severity = classify_row_diff(r_num, uid, name, rel, diff_cols)
        if r_num in specific_notes:
            summary_note = specific_notes[r_num]
        else:
            col_summaries = []
            for h, v1, v2 in diff_cols:
                if h == "한글 약력 정렬":
                    s1_str, s2_str = str(v1 or ""), str(v2 or "")
                    extra = []
                    if "〇〇" in s1_str and "00" in s2_str:
                        extra.append("'〇〇'↔'00'")
                    if normalize_fmt(s1_str) != normalize_fmt(s2_str):
                        extra.append(summarize_str_diff(s1_str, s2_str))
                    else:
                        extra.append("연월일 숫자 표기(파일1 한자숫자 '一九..' ↔ 파일2 아라비아숫자 '19..')")
                    col_summaries.append(f"[{h}] " + ", ".join(extra))
                elif h == "한문 약력 정렬":
                    s1_str, s2_str = str(v1 or ""), str(v2 or "")
                    if s1_str.replace(" ", "") == s2_str.replace(" ", ""):
                        col_summaries.append(f"[{h}] '女' 뒤 공백 등 띄어쓰기 차이 ({summarize_str_diff(s1_str, s2_str)})")
                    else:
                        col_summaries.append(f"[{h}] {summarize_str_diff(s1_str, s2_str)}")
                else:
                    col_summaries.append(f"[{h}] {summarize_str_diff(v1, v2)}")
            summary_note = " / ".join(col_summaries)

        all_diff_rows.append({
            "r_num": r_num,
            "uid": uid,
            "name": name,
            "rel": rel,
            "father": father,
            "grandpa": grandpa,
            "cat_label": cat_label,
            "severity": severity,
            "diff_col_names": ", ".join(h for h, _, _ in diff_cols),
            "diff_cols": diff_cols,
            "summary_note": summary_note,
            "r1": r1,
            "r2": r2,
        })

        for h, v1, v2 in diff_cols:
            cell_note = specific_notes.get(r_num, summarize_str_diff(v1, v2))
            cell_diff_records.append({
                "r_num": r_num,
                "uid": uid,
                "name": name,
                "rel": rel,
                "father": father,
                "col_name": h,
                "cat_label": cat_label,
                "severity": severity,
                "v1": v1,
                "v2": v2,
                "cell_note": cell_note,
            })

print(f"Total diff rows: {len(all_diff_rows)}, Total diff cells: {len(cell_diff_records)}")

wb_out = openpyxl.Workbook()

FONT_NAME = "Malgun Gothic"
font_title = Font(name=FONT_NAME, size=15, bold=True, color="1E293B")
font_subtitle = Font(name=FONT_NAME, size=10, italic=False, color="475569")
font_sec_hdr = Font(name=FONT_NAME, size=12, bold=True, color="1E3A8A")
font_tbl_hdr = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
font_data = Font(name=FONT_NAME, size=10, bold=False, color="0F172A")
font_data_bold = Font(name=FONT_NAME, size=10, bold=True, color="0F172A")
font_diff_red = Font(name=FONT_NAME, size=10, bold=True, color="991B1B")
font_diff_amber = Font(name=FONT_NAME, size=10, bold=True, color="92400E")
font_diff_blue = Font(name=FONT_NAME, size=10, bold=False, color="1E40AF")

fill_navy = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
fill_slate_hdr = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
fill_alert_red = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
fill_alert_amber = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
fill_alert_blue = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
fill_total = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

thin_side = Side(border_style="thin", color="CBD5E1")
med_navy = Side(border_style="medium", color="1E3A8A")
double_bottom = Side(border_style="double", color="1E293B")

border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
border_hdr = Border(left=thin_side, right=thin_side, top=med_navy, bottom=med_navy)
border_total = Border(top=thin_side, bottom=double_bottom, left=thin_side, right=thin_side)

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

# ==============================================================================
# SHEET 1: 비교요약_및_주요차이 (26개 핵심 행 + 요약 통계)
# ==============================================================================
ws_sum = wb_out.active
ws_sum.title = "1_비교요약_및_주요차이(26행)"
ws_sum.views.sheetView[0].showGridLines = True

ws_sum["A1"] = "완귀공파 족보 전사결과 시트(족보전사결과) 행별·열별 비교 분석 보고서"
ws_sum["A1"].font = font_title
ws_sum["A2"] = "비교 대상: [파일1] 보정 족보A_B_C_D_통합_전사결과.xlsx  vs  [파일2] 완귀공파 족보 통합.xlsx  |  기준 시트: 족보전사결과 (총 2,887개 데이터 행, 9개 열)"
ws_sum["A2"].font = font_subtitle
ws_sum.row_dimensions[1].height = 28
ws_sum.row_dimensions[2].height = 20

ws_sum["A4"] = "1. 전사결과 시트 비교 총괄 및 차이 유형별 집계"
ws_sum["A4"].font = font_sec_hdr
ws_sum.row_dimensions[4].height = 24

sum_headers = ["구분 코드", "차이 유형 분류", "중요도", "해당 행수(건)", "비율(%)", "주요 발생 열(컬럼)", "차이 내용 및 원인 요약"]
ws_sum.row_dimensions[5].height = 26
for c_idx, h_text in enumerate(sum_headers, 1):
    cell = ws_sum.cell(row=5, column=c_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_navy
    cell.alignment = align_center
    cell.border = border_hdr

cat_rows_data = [
    ("유형 ①", "① 약력 정렬 내용 전면 상이", "심각(내용불일치)", 2, "한문 약력 정렬, 한글 약력 정렬",
     "행 117(경익: 파일1에 타 인물 약력 오배치), 행 222(상중: 파일1에 출계 주기만 축약, 파일2에 전체 약력 수록)"),
    ("유형 ②", "② 계자(系子) 생부 인명·독음 상이", "중요(인명불일치)", 5, "한문/한글 원문 및 정렬",
     "행 695(종기: 구원↔진원), 행 711(종채: 상원↔필원), 행 764(진수: 종록↔종백), 행 898(명수: 언수↔언도/언탁), 행 1169(병재: 泓洙↔弘洙)"),
    ("유형 ③", "③ 한문 약력 정렬 내 한글 토씨(조사) 유무 차이", "경미(조사유무)", 7, "한문 약력 정렬",
     "파일1은 한문 정렬 시 한글 조사('라하여', '는', '에', '에야', '며')를 제거했으나, 파일2는 원문의 한글 조사가 남아 있음 (행 357, 781, 794, 960, 1005, 1127, 2068)"),
    ("유형 ④", "④ 한글 약력 원문 내 한자('正월'·'年') 표기 차이", "경미(한자표기)", 5, "한글 약력 원문",
     "파일1은 한글로 '정월', '년'으로 전사했으나, 파일2는 한자 '正월'(4건: 행 44, 90, 158, 172) 및 '年'(1건: 행 147)으로 표기함"),
    ("유형 ⑤", "⑤ 남자 항목 한문 약력 정렬 단순 띄어쓰기 차이", "서식(띄어쓰기)", 7, "한문 약력 정렬",
     "파일1은 '墓 后', '兩配 位', '公墓 同麓坐', '墓 同園'으로 띄어 쓰고, 파일2는 '墓后', '兩配位', '公墓同麓坐', '墓同園'으로 붙여 씀"),
    ("유형 ⑥", "⑥ 녀(女) 항목 약력 정렬 숫자(한자↔아라비아)·띄어쓰기 차이", "서식(숫자/공백)", 594, "한글 약력 정렬, 한문 약력 정렬",
     "C·D권 녀(女) 항목 한글 정렬에서 파일1은 한자 숫자('一九三一년 三월...')·'〇〇' 유지, 파일2는 아라비아 숫자('1931년 3월...')·'00' 변환 및 한문 정렬 '女 ' 뒤 공백 추가"),
]

for idx, (code, cname, sev, cnt, cols_affected, desc) in enumerate(cat_rows_data, start=6):
    ws_sum.row_dimensions[idx].height = 26
    ws_sum.cell(row=idx, column=1, value=code).alignment = align_center
    ws_sum.cell(row=idx, column=2, value=cname).alignment = align_left
    ws_sum.cell(row=idx, column=3, value=sev).alignment = align_center
    c_cnt = ws_sum.cell(row=idx, column=4, value=cnt)
    c_cnt.alignment = align_right
    c_cnt.number_format = "#,##0"
    c_pct = ws_sum.cell(row=idx, column=5, value=f"=D{idx}/$D$12")
    c_pct.alignment = align_right
    c_pct.number_format = "0.0%"
    ws_sum.cell(row=idx, column=6, value=cols_affected).alignment = align_center
    ws_sum.cell(row=idx, column=7, value=desc).alignment = align_left

    row_fill = fill_alert_red if idx == 6 else (fill_alert_amber if idx in (7, 8, 9) else fill_white)
    for col_i in range(1, 8):
        c = ws_sum.cell(row=idx, column=col_i)
        c.font = font_data_bold if idx in (6, 7) else font_data
        c.fill = row_fill
        c.border = border_data

ws_sum.row_dimensions[12].height = 25
ws_sum.cell(row=12, column=1, value="합계").alignment = align_center
ws_sum.cell(row=12, column=2, value="전체 상이 행 합계 (총 2,887행 중 2,267행 완전 일치)").alignment = align_left
ws_sum.cell(row=12, column=3, value="-").alignment = align_center
c_tot = ws_sum.cell(row=12, column=4, value="=SUM(D6:D11)")
c_tot.alignment = align_right
c_tot.number_format = "#,##0"
c_tot_pct = ws_sum.cell(row=12, column=5, value="=SUM(E6:E11)")
c_tot_pct.alignment = align_right
c_tot_pct.number_format = "0.0%"
ws_sum.cell(row=12, column=6, value="4개 약력 열 (ID·인명·관계·부친·조부 100% 일치)").alignment = align_center
ws_sum.cell(row=12, column=7, value="실질적 내용·표기·띄어쓰기 차이(유형 ①~⑤) 26개 행 + 녀(女) 일괄 서식 차이(유형 ⑥) 594개 행 = 총 620개 행").alignment = align_left
for col_i in range(1, 8):
    c = ws_sum.cell(row=12, column=col_i)
    c.font = font_data_bold
    c.fill = fill_total
    c.border = border_total

ws_sum["A14"] = "2. 주요 내용·인명·원문 표기 및 남자 항목 상이 내역 상세 (유형 ①~⑤, 총 26개 행)"
ws_sum["A14"].font = font_sec_hdr
ws_sum.row_dimensions[14].height = 24

det_headers = [
    "연번", "엑셀 행번호", "고유ID", "인명", "관계", "차이 유형",
    "상이 열(컬럼)", "[파일1] 보정 족보A_B_C_D_통합_전사결과.xlsx", "[파일2] 완귀공파 족보 통합.xlsx", "구체적 차이 내용 설명"
]
ws_sum.row_dimensions[15].height = 28
for c_idx, h_text in enumerate(det_headers, 1):
    cell = ws_sum.cell(row=15, column=c_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_navy
    cell.alignment = align_center
    cell.border = border_hdr

major_rows = [item for item in all_diff_rows if not item["cat
