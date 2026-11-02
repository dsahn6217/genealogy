import sqlite3
import pandas as pd
import openpyxl

# 엑셀 읽기
excel_file = "genealogy.xlsx"
wb = openpyxl.load_workbook(excel_file, read_only=True)
sheet_name = "족보전사결과" if "족보전사결과" in wb.sheetnames else wb.sheetnames[0]
ws = wb[sheet_name]

rows = list(ws.iter_rows(values_only=True))
header = [str(c).strip() if c is not None else f"Col{idx}" for idx, c in enumerate(rows[0])]
df = pd.DataFrame(rows[1:], columns=header).fillna("")

# SQLite DB 저장 및 인덱스 생성
conn = sqlite3.connect("genealogy.db")
df.to_sql("genealogy", conn, if_exists="replace", index=False)
conn.execute("CREATE INDEX IF NOT EXISTS idx_gid ON genealogy(고유ID);")
conn.execute("CREATE INDEX IF NOT EXISTS idx_name ON genealogy(인명);")
conn.commit()
conn.close()

print("✅ genealogy.db 변환 완료!")
