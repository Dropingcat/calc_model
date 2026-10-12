# -*- coding: utf-8 -*-
"""Аудит v16: чтение реальных чисел из книги + независимая сверка профиля."""
import io, json, math, os, sys, warnings
import numpy as np
import openpyxl

P = r"C:\Users\Arhys\Desktop\каллибровка\ГОСТ EN 13132\v16\EN13132_v16_CALC.xlsx"
OUT = {}

wb = openpyxl.load_workbook(P, data_only=True)
OUT["sheets"] = wb.sheetnames

# --- 06_Итог: ищем строки с ключевыми метками ---
ws6 = wb["06_Итог"]
OUT["ws6_rows"] = []
for r in range(1, min(ws6.max_row, 80) + 1):
    a = ws6.cell(r, 1).value
    b = ws6.cell(r, 2).value
    if a is not None:
        OUT["ws6_rows"].append([r, str(a)[:46], b])

# --- 05_Бюджет: строки с "прецизионность"/"эталоны"/"чистота" ---
ws5 = wb["05_Бюджет"]
OUT["ws5_rows"] = []
for r in range(1, min(ws5.max_row, 60) + 1):
    a = ws5.cell(r, 1).value
    b = ws5.cell(r, 2).value
    if a is not None:
        OUT["ws5_rows"].append([r, str(a)[:46], b])

# --- 04_Профиль U(X): заголовок + первые/последние строки ---
ws4 = wb["04_Профиль U(X)"]
OUT["ws4_meta"] = []
for r in range(1, min(ws4.max_row, 15) + 1):
    vals = [ws4.cell(r, c).value for c in range(1, 9)]
    OUT["ws4_meta"].append([vals[0], vals[1], vals[6], vals[7]])
# заголовок профиля: ищем строку с "X (г/г)"
hdr_row = None
for r in range(1, ws4.max_row + 1):
    if ws4.cell(r, 2).value == "X (г/г)":
        hdr_row = r
        break
OUT["hdr_row"] = hdr_row
profile = []
if hdr_row:
    for r in range(hdr_row + 1, ws4.max_row + 1):
        x = ws4.cell(r, 2).value
        u = ws4.cell(r, 7).value  # столбец G = u_c? проверим из шапки
        if isinstance(x, (int, float)) and isinstance(u, (int, float)):
            profile.append((float(x), float(u)))
OUT["profile_pts"] = len(profile)
if profile:
    OUT["profile_first"] = profile[0]
    OUT["profile_last"] = profile[-1]
    # шаг u_repr: столбец D = u_repr
    du = []
    hdr_col_d = 4
    if hdr_row:
        for r in range(hdr_row + 1, ws4.max_row + 1):
            d = ws4.cell(r, hdr_col_d).value
            if isinstance(d, (int, float)):
                du.append(float(d))
    OUT["D_col_numeric"] = len(du)
    if len(du) > 1:
        OUT["max_step_D"] = max(abs(du[i+1] - du[i]) for i in range(len(du)-1))

# --- Графики: D4 (маркер) и C4 (x_pred) ---
wsg = wb["Графики"]
OUT["gch_D4"] = wsg["D4"].value
OUT["gch_C4"] = wsg["C4"].value

with io.open(r"C:\Temp\v16_audit.json", "w", encoding="utf-8") as f:
    json.dump(OUT, f, ensure_ascii=False, indent=1, default=str)
print("OK")
