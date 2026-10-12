import openpyxl, json, io, sys

PATH = r"C:\Temp\v17_recalc\EN13132_v17_RECALC.xlsx"
OUT = io.open(r"C:\Temp\v17_recalc\report.txt", "w", encoding="utf-8")
def log(*a):
    OUT.write(" ".join(str(x) for x in a) + "\n")

wb = openpyxl.load_workbook(PATH, data_only=True)

ws = wb["06_\u0418\u0442\u043e\u0433"]
log("=== 06_\u0418\u0442\u043e\u0433 B1:B30 ===")
for r in range(1, 31):
    a = ws.cell(r, 1).value
    b = ws.cell(r, 2).value
    if a is not None or b is not None:
        log(r, "| A=", repr(a), "| B=", repr(b))

# Скан ошибок по всем листам
ERR = ["#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NAME?", "#NUM!", "#NULL!"]
log("\n=== ЯЧЕЙКИ С ОШИБКАМИ (data_only=True) ===")
for sname in wb.sheetnames:
    s = wb[sname]
    cnt = 0
    for row in s.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v.startswith("#") and v in ERR:
                log(sname, "!", c.coordinate, "=", repr(v))
                cnt += 1
    log(sname, ": кол-во ошибок =", cnt)

# 04_Профиль: считаем ок/ошибки/None
s = wb["04_\u041f\u0440\u043e\u0444\u0438\u043b\u044c U(X)"]
ok = err = none = 0
for r in range(5, 205):
    for col in range(1, 13):
        v = s.cell(r, col).value
        if isinstance(v, str) and v.startswith("#"):
            err += 1
        elif v is None:
            none += 1
        else:
            ok += 1
log("\n=== 04_Профиль строки 5..204 (200 строк x 12 колонок) ===")
log("ok:", ok, "errors:", err, "None:", none)

log("\n=== 04_Профиль первые 5 строк (B..H) ===")
for r in range(5, 10):
    vals = [s.cell(r, col).value for col in range(2, 9)]
    log(r, vals)

log("\n=== 04_Профиль поиск маркера x_pred ===")
for r in range(5, 205):
    h = s.cell(r, 8).value
    i = s.cell(r, 9).value
    if isinstance(i, str) and "x_pred" in str(i):
        log("маркер-строка", r, "B,H:", s.cell(r,2).value, s.cell(r,8).value, "| I:", repr(i))
    if h is not None and h != "" and str(h).strip() == "\u2190 x_pred":
        log("строки с указателем:", r, s.cell(r,2).value, s.cell(r,8).value)

# 02_Регрессия результат
s2 = wb["02_\u0420\u0435\u0433\u0440\u0435\u0441\u0441\u0438\u044f"]
log("\n=== 02_Регрессия B1:B40 ===")
for r in range(1, 41):
    b = s2.cell(r, 2).value
    if b is not None:
        log(r, "B=", repr(b))

# 03_ANOVA, 05_Бюджет ключевое
s3 = wb["03_ANOVA"]
log("\n=== 03_ANOVA (первые 40 строк) ===")
for r in range(1, 41):
    row = [s3.cell(r, c).value for c in range(1, 9)]
    if any(x is not None for x in row):
        log(r, row)

s5 = wb["05_\u0411\u044e\u0434\u0436\u0435\u0442"]
log("\n=== 05_Бюджет (B..K, строки 1..40) ===")
for r in range(1, 41):
    row = [s5.cell(r, c).value for c in range(1, 12)]
    if any(x is not None for x in row):
        log(r, [repr(x) for x in row])

OUT.close()
print("written")