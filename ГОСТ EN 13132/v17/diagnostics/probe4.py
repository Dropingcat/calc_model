import openpyxl, json, io

SRC = r"C:\Users\Arhys\Desktop\каллибровка\ГОСТ EN 13132\v17\EN13132_v17_LIVE.xlsx"
REC = r"C:\Temp\v17_recalc\EN13132_v17_RECALC.xlsx"

OUT = io.open(r"C:\Temp\v17_recalc\report3.txt", "w", encoding="utf-8")
def log(*a):
    OUT.write(" ".join(str(x) for x in a) + "\n")

wbf = openpyxl.load_workbook(SRC, data_only=False)
wbv = openpyxl.load_workbook(REC, data_only=True)

# 1. K4/L4 и K5-L5 в 04_Профиль (значения и формулы)
sw_f = wbf["04_\u041f\u0440\u043e\u0444\u0438\u043b\u044c U(X)"]
sw_v = wbv["04_\u041f\u0440\u043e\u0444\u0438\u043b\u044c U(X)"]
log("=== 04_Профиль K/L ячейки ===")
for addr in ["A4","B4","C4","H4","I4","J4","K4","L4","K5","L5","K67","L67"]:
    log(addr, "| formula:", repr(sw_f[addr].value), "| value:", repr(sw_v[addr].value))

# 2. 02_Регрессия A52:A151 и C52:C151 (chart0 ser1: x A52:A151, y C52:C151)
s2_f = wbf["02_\u0420\u0435\u0433\u0440\u0435\u0441\u0441\u0438\u044f"]
s2_v = wbv["02_\u0420\u0435\u0433\u0440\u0435\u0441\u0441\u0438\u044f"]
log("\n=== 02_Регрессия строки 52..151 (данные для серии x A52:A151) ===")
filled = 0
for r in range(52, 152):
    vals = [s2_v.cell(r, c).value for c in range(1, 8)]
    if any(v is not None for v in vals):
        filled += 1
log("строк с данными в 52..151:", filled)
for r in [52, 60, 80, 100, 120, 140, 151]:
    a = s2_f.cell(r, 1).value
    c = s2_f.cell(r, 3).value
    av = s2_v.cell(r, 1).value
    cv = s2_v.cell(r, 3).value
    log(r, "A_f:", repr(a), "C_f:", repr(c), "| A_v:", repr(av), "C_v:", repr(cv))

# 3. Chart cache после пересчёта
wbc = openpyxl.load_workbook(REC, data_only=True)
wsc = wbc["\u0413\u0440\u0430\u0444\u0438\u043a\u0438"]
log("\n=== Графики в ПЕРЕСЧИТАННОЙ книге (кэш серий) ===")
for i, ch in enumerate(wsc._charts):
    log("chart", i, type(ch).__name__)
    for j, series in enumerate(ch.series):
        x, y = series.xVal, series.yVal
        def ref(v):
            if v is None: return None
            inner = getattr(v, 'numRef', None) or getattr(v, 'strRef', None)
            if inner is None: return str(type(v))
            return json.dumps({'f': getattr(inner, 'f', None), 'pts': len(getattr(getattr(inner,'numCache',None) or getattr(inner,'strCache',None), 'pt', []) or [])}, ensure_ascii=True)
        log("  ser", j, "x:", ref(x), "y:", ref(y))

# 4. Hex A30 формула (ищем скрытые символы)
log("\n=== 06_Итог A30 — codepoints формулы ===")
f30 = wbf["06_\u0418\u0442\u043e\u0433"]["A30"].value
log("len:", len(f30))
code = [c for c in f30]
uniq = sorted(set(code))
log("уникальные символы:", [(c, hex(ord(c))) for c in uniq])

# 5. chart titles / anchors для сопоставления
log("\n=== anchors графиков (положение на листе) ===")
for i, ch in enumerate(wsc._charts):
    a = ch.anchor
    try:
        frm = (a._from.col, a._from.row)
        to = (a._to.col, a._to.row)
    except AttributeError:
        frm = (a._from.col, a._from.row)
        to = None
    ttl = None
    try:
        if ch.title and ch.title.tx and ch.title.tx.rich:
            ttl = " | ".join(p.text for p in ch.title.tx.rich.p)
    except Exception:
        ttl = "?"
    log("chart", i, "anchor:", frm, "->", to, "| title:", ttl)

# 6. Проверим литерал запятой vs точка в TEXT формате A30 (локаль!)
# В русской локали Excel 16 запятая - десятичный разделитель; "0.0" все равно ок обычно.
# но проверим ещё ячейки B в 06_Итог 26,27
log("\n=== 06_Итог B26,B27 формулы и значения ===")
log("B26 formula:", repr(wbf["06_\u0418\u0442\u043e\u0433"]["B26"].value), "value:", repr(wbv["06_\u0418\u0442\u043e\u0433"]["B26"].value))
log("B27 formula:", repr(wbf["06_\u0418\u0442\u043e\u0433"]["B27"].value), "value:", repr(wbv["06_\u0418\u0442\u043e\u0433"]["B27"].value))
log("A30 value:", repr(wbv["06_\u0418\u0442\u043e\u0433"]["A30"].value))

# 7. 02_Регрессия строки 39-50 - G столбец формула
log("\n=== 02_Регрессия G39..G50 формула ===")
for r in range(39, 51):
    log("G"+str(r), repr(s2_f.cell(r,7).value), "| value:", repr(s2_v.cell(r,7).value))

OUT.close()
print("done")