# -*- coding: utf-8 -*-
"""make_chart_check_input.py — контрольная копия LIVE для чистого chart_check.

Блок `gost13132-v17-chart-axes`. Пользователь после фикса обнулил в
EN13132_v17_LIVE.xlsx параметры результата пробы (m_пробы=0, m_ИС=0, p=0,
Y0=0, режим=2) — при таких входах Excel закономерно даёт #DIV/0!/#N/A
(K_G2PCT = m_ИС·p_ИС/m_пробы → 0/0), что НЕ связано с фиксом оси X.

Для ЧИСТОГО контрольного прогона chart_check создаём копию
v17\\tmp\\chart_check_input.xlsx, в которой в 01_Ввод записаны КОРРЕКТНЫЕ
дефолтные входы результата пробы (по INPUT_FIELDS build_en13132_v17_calc):
m_пробы=100, m_ИС=4, p_ИС=0.995, p_A=0.995, Y0=средний отклик среднего
уровня, p=2, u(m_net)=5.7735e-05, u(pA)=0.00288675, u_rel=0.001, режим=1.

Правило путей (AGENTS_PATHS_RULE.md): файлы — только в v17\\tmp\\.
"""
from __future__ import annotations

import os
import shutil
import sys

sys.stdout.reconfigure(encoding="utf-8")

import numpy as np  # noqa: E402
import openpyxl  # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import build_en13132_v17_calc as bld  # noqa: E402

V17_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIVE = os.path.join(V17_DIR, "EN13132_v17_LIVE.xlsx")
DST = os.path.join(V17_DIR, "tmp", "chart_check_input.xlsx")

print("Источник:", LIVE)
print("Копия:", DST)

os.makedirs(os.path.dirname(DST), exist_ok=True)
shutil.copy2(LIVE, DST)

wb = openpyxl.load_workbook(DST, data_only=False)
ws = wb["01_Ввод"]

# уровни из копии (заполненные)
levels = bld.read_levels(ws)
assert levels is not None and len(levels) >= 3, f"уровни: {levels!r}"
print("Уровней:", len(levels))

inputs = bld.default_inputs(levels)
print("Дефолтные входы:", {k: (round(v, 8) if isinstance(v, float) else v)
                            for k, v in inputs.items()})

# записать дефолты в ячейки 01_Ввод по меткам
for key, label, _default, _fmt, _units, _cmt in bld.INPUT_FIELDS:
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith(label):
            ws.cell(r, 2, inputs[key])
            break

# режим u_repr -> 1 (step)
for r in range(1, ws.max_row + 1):
    v = ws.cell(r, 1).value
    if isinstance(v, str) and v.startswith("Режим u_repr"):
        ws.cell(r, 2, 1)
        break

wb.save(DST)
print("Сохранено. Проверка записанных входов:")
wb2 = openpyxl.load_workbook(DST, data_only=False)
ws2 = wb2["01_Ввод"]
for r in range(18, 32):
    a = ws2.cell(r, 1).value
    b = ws2.cell(r, 2).value
    if a is not None:
        print(f"  B{r} [{str(a)[:42]}]: {b!r}")