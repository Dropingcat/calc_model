import openpyxl, io, json

SRC = r"C:\Users\Arhys\Desktop\каллибровка\ГОСТ EN 13132\v17\EN13132_v17_LIVE.xlsx"
OUT = io.open(r"C:\Temp\v17_recalc\report9.txt", "w", encoding="utf-8")
def log(*a):
    OUT.write(" ".join(str(x) for x in a) + "\n")

wbf = openpyxl.load_workbook(SRC, data_only=False)
wbv = openpyxl.load_workbook(r"C:\Temp\v17_recalc\EN13132_v17_RECALC.xlsx", data_only=True)

# 01_Ввод: столбцы J (SS_PE?) — посмотрим, где столбец J
s1v = wbv["01_\u0412\u0432\u043e\u0434"]
log("=== 01_Ввод headers row 4 ===")
for c in range(1, 14):
    v = s1v.cell(4, c).value
    if v is not None:
        log(chr(64+c), repr(v))
log("\n=== 01_Ввод строки 5..16, столбцы I..M (помощь/SS_PE) ===")
for r in range(5, 17):
    row = [s1v.cell(r, c).value for c in range(9, 14)]
    if any(x is not None for x in row):
        log(r, [repr(x) for x in row])
    else:
        log(r, "(пусто)")

# 02_Регрессия заголовки
s2v = wbv["02_\u0420\u0435\u0433\u0440\u0435\u0441\u0441\u0438\u044f"]
log("\n=== 02_Регрессия строки 1..3 и 37..38, 51 (заголовки) ===")
for r in [1,2,3,37,38,51]:
    row = [s2v.cell(r, c).value for c in range(1, 9)]
    log(r, [repr(x) for x in row])

OUT.close()
print("done")