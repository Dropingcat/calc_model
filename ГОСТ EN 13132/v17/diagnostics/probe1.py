import openpyxl, json
from openpyxl.utils import get_column_letter

PATH = r"C:\Temp\v17_recalc\EN13132_v17_RECALC.xlsx"
wb = openpyxl.load_workbook(PATH, data_only=True)
print("SHEETS:", json.dumps(wb.sheetnames, ensure_ascii=True))

ws = wb["06_\u0418\u0442\u043e\u0433"]
print("\n=== 06_Итог B1:B30 ===")
for r in range(1, 31):
    a = ws.cell(r, 1).value
    b = ws.cell(r, 2).value
    if a is not None or b is not None:
        print(r, "| A=", repr(a), "| B=", repr(b))