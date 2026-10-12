# -*- coding: utf-8 -*-
"""Независимый аудит v16 (reviewer): чтение реальных ячеек книги + ядро."""
import io, sys, json, math, os
sys.path.insert(0, r"C:\Users\Arhys\Desktop\каллибровка\ГОСТ EN 13132\v16")
import numpy as np
import openpyxl

import anova_gum as ag

OUT = {}
# имя листа профиля — найти по префиксу "04_"
wb_path = r"C:\Users\Arhys\Desktop\каллибровка\ГОСТ EN 13132\v16\EN13132_v16_CALC.xlsx"
wb = openpyxl.load_workbook(wb_path, data_only=True)
sheets = wb.sheetnames
prof_sheet = next(s for s in sheets if s.startswith("04_"))
itog_sheet = next(s for s in sheets if s.startswith("06_"))
OUT["sheets"] = sheets
OUT["prof_sheet"] = prof_sheet
OUT["itog_sheet"] = itog_sheet

# ---- 06_Итог: результат/U ----
ws6 = wb[itog_sheet]
OUT["ws6_rows"] = []
for r in range(1, min(ws6.max_row, 60) + 1):
    a = ws6.cell(r, 1).value
    b = ws6.cell(r, 2).value
    if a is not None:
        OUT["ws6_rows"].append([r, str(a)[:46], b])

# ---- 04_Профиль: шапка, D4-маркер, максимум шага u_repr ----
wsp = wb[prof_sheet]
hdr_row = None
for r in range(1, 12):
    v = wsp.cell(r, 1).value
    if isinstance(v, str) and ("№" in v or "X (г/г)" in v):
        hdr_row = r
        break
OUT["prof_hdr_row"] = hdr_row
if hdr_row:
    # u_repr столбец (берём реальный столбец D: 4)
    uvals = []
    xvals = []
    for r in range(hdr_row + 1, hdr_row + 202):
        u = wsp.cell(r, 4).value
        x = wsp.cell(r, 2).value
        if isinstance(u, (int, float)) and isinstance(x, (int, float)):
            uvals.append(u)
            xvals.append(x)
        else:
            break
    OUT["prof_points"] = len(uvals)
    if len(uvals) > 2:
        du = [abs(uvals[i + 1] - uvals[i]) for i in range(len(uvals) - 1)]
        OUT["max_step_u_repr"] = max(du)
        OUT["u_repr_first"] = uvals[0]
        OUT["u_repr_last"] = uvals[-1]
        # гладкость: максимальный шаг между соседями в % мас -> в г/г?
        # столбец D это u_repr в % мас (г/г в г/г?) — сравним с ядром в те же x
    OUT["x_first"] = xvals[0] if xvals else None
    OUT["x_last"] = xvals[-1] if xvals else None

# -- D4 обычно на листе "Графики" --
wbg = wb["Графики"]
OUT["gch_D4"] = wbg["D4"].value
OUT["gch_C4"] = wbg["C4"].value

# ---- Бюджет 05: u_repr(06_Итог) ----
ws5 = wb["05_Бюджет"]
OUT["ws5_rows"] = []
for r in range(1, min(ws5.max_row, 30) + 1):
    a = ws5.cell(r, 1).value
    b = ws5.cell(r, 2).value
    if a is not None:
        OUT["ws5_rows"].append([r, str(a)[:44], b])

# ---- Ядро: w0 из книги 01_Ввод ----
wbv = openpyxl.load_workbook(wb_path, data_only=False)
ws1 = wbv["01_Ввод"]
def find_v15_label(ws, needle):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and needle in v:
            return r
    return None

# соберём уровни как в ядре (H-колонка и Y1..Y3)
def read_levels(ws):
    # попробуем динамически: ищем заголовок с X (г/г)
    hdr = None
    for r in range(1, min(ws.max_row, 40) + 1):
        rowvals = [ws.cell(r, c).value for c in range(1, 10)]
        if any(isinstance(v, str) and "г/г" in v for v in rowvals):
            hdr = r
            break
    if not hdr:
        return None
    col = {}
    for c in range(1, 10):
        v = ws.cell(hdr, c).value
        if isinstance(v, str):
            col[v.strip()] = c
    XC = None
    for k in col:
        if k.strip().lower().startswith("x (") or k.strip().lower().startswith("x("):
            XC = col[k]
    ycols = [col[k] for k in col if k.strip() in ("Y1", "Y2", "Y3")]
    if not XC:
        XC = 8  # H
    levels = []
    for r in range(hdr + 1, min(ws.max_row, 40) + 1):
        x = ws.cell(r, XC).value
        ys = [ws.cell(r, c).value for c in ycols if isinstance(ws.cell(r, c).value, (int, float))]
        if x is not None and ys:
            levels.append((float(x), ys))
        elif x is None and not ys:
            break
    return levels

lv = read_levels(ws1)
OUT["n_levels"] = len(lv) if lv else None

if lv and len(lv) >= 3:
    snap = ag.ols_slopes_by_means(lv)
    OUT["a"] = snap["a"]
    OUT["b"] = snap["b"]
    OUT["r2"] = snap["r2"]
    OUT["s2_res"] = snap["s2_res"]
    # результат пробы — y0 из среднего уровня? возьмём Y0 пробы из книги:
    # Y0_PROBA (06_Итог строка) — ищем в 01_Ввод "Y0 пробы"
    y0r = find_v15_label(ws1, "Y0 пробы")
    OUT["y0_label_row"] = y0r
    y0 = None
    if y0r:
        y0 = ws1.cell(y0r, 2).value
    OUT["y0_proba_value"] = y0
    # X0
    if y0 is not None and abs(snap["a"]) > 1e-12:
        x0 = (float(y0) - snap["b"]) / snap["a"]
        OUT["X0"] = x0
        w0 = x0 * 3.966567307590398  # K_G2PCT?
        OUT["w0"] = w0
        # smooth u_repr
        try:
            rp = ag.u_repr_from_tables(w0, ag.TABLE1, "mean2", "smooth")
            OUT["u_repr_smooth"] = rp["u_repr"]
        except Exception as e:
            OUT["u_repr_smooth"] = "ERR:" + str(e)
        # step
        try:
            rp2 = ag.u_repr_from_tables(w0, ag.TABLE1, "mean2", "step")
            OUT["u_repr_step"] = rp2["u_repr"]
        except Exception as e:
            OUT["u_repr_step"] = "ERR:" + str(e)

with io.open(r"C:\Temp\uaudit_v16.json", "w", encoding="utf-8") as f:
    json.dump(OUT, f, ensure_ascii=False, indent=1, default=str)
print("OK")
