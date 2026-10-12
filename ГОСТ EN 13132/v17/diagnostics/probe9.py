import openpyxl, io

REC = r"C:\Temp\v17_recalc\EN13132_v17_RECALC.xlsx"
OUT = io.open(r"C:\Temp\v17_recalc\report8.txt", "w", encoding="utf-8")
def log(*a):
    OUT.write(" ".join(str(x) for x in a) + "\n")

wbv = openpyxl.load_workbook(REC, data_only=True)
s2 = wbv["02_\u0420\u0435\u0433\u0440\u0435\u0441\u0441\u0438\u044f"]
log("=== 02_Регрессия E,F,G,D 39..50 (доверительная полоса/остатки) ===")
for r in range(39, 51):
    log(r, ["C", s2.cell(r,3).value], ["D", s2.cell(r,4).value], ["E", s2.cell(r,5).value], ["F", s2.cell(r,6).value], ["G", s2.cell(r,7).value])

log("\n=== 02_Регрессия: A39:A50 значения (сетка по уровням) ===")
for r in range(39, 51):
    log(r, repr(s2.cell(r,1).value), "B:", repr(s2.cell(r,2).value))

# B-столбец 02 полный
log("\n=== 02_Регрессия B1:B36 значения ===")
for r in range(1, 37):
    b = s2.cell(r,2).value
    if b is not None:
        log("B"+str(r), repr(b))

OUT.close()
print("done")