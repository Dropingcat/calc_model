import openpyxl, io

SRC = r"C:\Users\Arhys\Desktop\каллибровка\ГОСТ EN 13132\v17\EN13132_v17_LIVE.xlsx"
REC = r"C:\Temp\v17_recalc\EN13132_v17_RECALC.xlsx"

OUT = io.open(r"C:\Temp\v17_recalc\report4.txt", "w", encoding="utf-8")
def log(*a):
    OUT.write(" ".join(str(x) for x in a) + "\n")

wbf = openpyxl.load_workbook(SRC, data_only=False)
wbv = openpyxl.load_workbook(REC, data_only=True)
s2_f = wbf["02_\u0420\u0435\u0433\u0440\u0435\u0441\u0441\u0438\u044f"]
s2_v = wbv["02_\u0420\u0435\u0433\u0440\u0435\u0441\u0441\u0438\u044f"]

log("=== Формулы A52:A151 (сетка X) — паттерн ===")
pat = {}
for r in range(52, 152):
    f = s2_f.cell(r, 1).value
    a = s2_v.cell(r, 1).value
    if isinstance(f, str):
        key = f
        pat.setdefault(key, (r, []))
        pat[key][1].append((r, a))
# покажем уникальные варианты формул
for k, (firstrow, vals) in pat.items():
    log("первая строка:", firstrow, "форм:", repr(k))
# выведем выборку значений
log("\nЗначения A52:A151 (каждые 10 строк):")
for r in range(52, 152, 10):
    log(r, repr(s2_v.cell(r,1).value))
log("мин/макс в A52:A151:", min((s2_v.cell(r,1).value for r in range(52,152))), max((s2_v.cell(r,1).value for r in range(52,152))))

# Позиция формулы сетки: проверим, есть ли варианты с (ROW()-52) в других местах
log("\n=== 02_Регрессия: где ещё (ROW()-52)? ===")
for row in s2_f.iter_rows():
    for c in row:
        v = c.value
        if isinstance(v, str) and "ROW()-52" in v:
            log(c.coordinate, repr(v[:120]))

OUT.close()
print("done")