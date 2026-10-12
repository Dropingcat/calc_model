import openpyxl, json, io
from openpyxl.utils import get_column_letter

SRC = r"C:\Users\Arhys\Desktop\каллибровка\ГОСТ EN 13132\v17\EN13132_v17_LIVE.xlsx"
REC = r"C:\Temp\v17_recalc\EN13132_v17_RECALC.xlsx"

OUT = io.open(r"C:\Temp\v17_recalc\report2.txt", "w", encoding="utf-8")
def log(*a):
    OUT.write(" ".join(str(x) for x in a) + "\n")

wbf = openpyxl.load_workbook(SRC, data_only=False)
wbv = openpyxl.load_workbook(REC, data_only=True)

# 1. Формула 06_Итог A30
log("=== 06_Итог формулы A28:A35 ===")
wsf = wbf["06_\u0418\u0442\u043e\u0433"]
wsv = wbv["06_\u0418\u0442\u043e\u0433"]
for r in range(28, 36):
    log("row", r, "| A formula:", repr(wsf.cell(r,1).value), "| A value:", repr(wsv.cell(r,1).value))
    for c in range(2,4):
        log("   ", get_column_letter(c)+str(r), "formula:", repr(wsf.cell(r,c).value), "value:", repr(wsv.cell(r,c).value))

# 2. Все формулы с делением в 06_Итог, 03_ANOVA, 04_Профиль, 05_Бюджет
log("\n=== ФОРМУЛЫ С ДЕЛЕНИЕМ (03_ANOVA, 05_Бюджет, 06_Итог) ===")
for sname in ["03_ANOVA", "05_\u0411\u044e\u0434\u0436\u0435\u0442", "06_\u0418\u0442\u043e\u0433"]:
    ws = wbf[sname]
    cnt = 0
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("=") and "/" in v:
                log(sname, cell.coordinate, "->", v)
                cnt += 1
    log(sname, "всего формул с `/`:", cnt)

# 3. Обёртки IFERROR в 03/04/05/06
log("\n=== IFERROR/IF-обёртки, маскирующие ошибки ===")
for sname in ["03_ANOVA", "04_\u041f\u0440\u043e\u0444\u0438\u043b\u044c U(X)", "05_\u0411\u044e\u0434\u0436\u0435\u0442", "06_\u0418\u0442\u043e\u0433", "02_\u0420\u0435\u0433\u0440\u0435\u0441\u0441\u0438\u044f"]:
    ws = wbf[sname]
    addrs_iferr = []
    addrs_ifempty = []
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                if "IFERROR(" in v.upper():
                    addrs_iferr.append(cell.coordinate)
                if "/" in v and ("IF(" in v.upper() or 'IF(' in v):
                    addrs_ifempty.append(cell.coordinate)
    log(sname, "| IFERROR ячеек:", len(addrs_iferr), "| ячеек с `/` внутри IF:", len(addrs_ifempty))
    if addrs_iferr:
        log("   IFERROR:", addrs_iferr[:60])

# 4. Где None в 04_Профиль
log("\n=== 04_Профиль: распределение None по колонкам (пересчитано) ===")
sw = wbv["04_\u041f\u0440\u043e\u0444\u0438\u043b\u044c U(X)"]
colstats = {}
for r in range(5,205):
    for c in range(1,13):
        v = sw.cell(r,c).value
        colstats.setdefault(get_column_letter(c), [0,0,0])
        if v is None:
            colstats[get_column_letter(c)][0]+=1
        elif isinstance(v,str) and v.startswith("#"):
            colstats[get_column_letter(c)][1]+=1
        else:
            colstats[get_column_letter(c)][2]+=1
for c in sorted(colstats.keys(), key=lambda x: (len(x),x)):
    log(c, "None:", colstats[c][0], "errors:", colstats[c][1], "ok:", colstats[c][2])

# 5. Строка маркера
log("\n=== маркер-строка 67 все колонки (пересчитано) ===")
for r in [66,67,68]:
    log(r, [sw.cell(r,c).value for c in range(1,13)])

# 6. 05_Бюджет формулы B/I столбцы (вклады)
log("\n=== 05_Бюджет формулы (столбец B) ===")
ws5 = wbf["05_\u0411\u044e\u0434\u0436\u0435\u0442"]
for r in range(4,17):
    log("B"+str(r), repr(ws5.cell(r,2).value))

# 7. 06_Итог формулы ключевых ячеек (B19-B27, B3)
log("\n=== 06_Итог формулы ключевых ===")
for r in [3,19,20,21,22,23,24,25]:
    log(get_column_letter(2)+str(r), repr(wsf.cell(r,2).value))
    v = wsv.cell(r,2).value
    log("   value:", repr(v))

OUT.close()
print("done")