# -*- coding: utf-8 -*-
"""probe11_chart_axes.py — ДИАГНОСТИКА блока `gost13132-v17-chart-axes`.

Гипотеза: для незаполненных уровней 7..12 формулы A39:A50/B39:B50 на листе
02_Регрессия возвращают "" (пустой текст) -> в диапазоне xVal серии точек
ScatterChart есть НЕчисловые ячейки -> Excel переключает ось X в категорийный
режим 1..n. Сетка линии регрессии (A52:A151 — чистые числа) работает, поэтому
«линия жива, а точки кривые».

ЗАДАЧА (факты):
  1) формулы A39:A50 и B39:B50 (что именно: ""? NA()? числа?);
  2) какие из A39:A50 числовые, какие текст/пусто (по кэшу и по данным);
  3) XML диаграмм листа «Графики»: xVal/yVal (numRef.f) КАЖДОЙ серии —
     подтвердить, что серия точек ссылается на A39:A50/B39:B50 и что там "".

Правило путей (AGENTS_PATHS_RULE.md): временные файлы — только в v17\diagnostics\,
НЕ в C:\\Temp.
"""
from __future__ import annotations

import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8")

import openpyxl  # noqa: E402

V17_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIVE = os.path.join(V17_DIR, "EN13132_v17_LIVE.xlsx")

print("=" * 70)
print("ПРОБА 11: диагностика оси X серии точек (калибровка)")
print("Книга:", LIVE)
print("=" * 70)

wb_f = openpyxl.load_workbook(LIVE, data_only=False)
wb_v = openpyxl.load_workbook(LIVE, data_only=True)
ws_f = wb_f["02_Регрессия"]
ws_v = wb_v["02_Регрессия"]

print("\n--- 1. Формулы A39:A50 и B39:B50 (02_Регрессия) ---")
for r in range(39, 51):
    fa = ws_f.cell(r, 1).value
    fb = ws_f.cell(r, 2).value
    va = ws_v.cell(r, 1).value
    vb = ws_v.cell(r, 2).value
    ta = type(va).__name__ if va is not None else "None"
    tb = type(vb).__name__ if vb is not None else "None"
    print(f"A{r}: формула={fa!r:65s} | кэш={va!r} ({ta})")
    print(f"B{r}: формула={fb!r:65s} | кэш={vb!r} ({tb})")

print("\n--- 2. Классификация ячеек A39:A50 (число/текст/пусто) ---")
for r in range(39, 51):
    va = ws_v.cell(r, 1).value
    if isinstance(va, (int, float)):
        cls = "ЧИСЛО"
    elif isinstance(va, str) and va == "":
        cls = "ПУСТО(текст \"\")"
    elif va is None:
        cls = "ПУСТО(None)"
    elif isinstance(va, str):
        cls = f"ТЕКСТ {va!r}"
    else:
        cls = f"{type(va).__name__} {va!r}"
    print(f"A{r}: {cls}")

print("\n--- 3. XML диаграмм листа «Графики»: numRef.f всех серий ---")
with zipfile.ZipFile(LIVE) as z:
    names = [n for n in z.namelist() if re.match(r"xl/charts/chart\d+\.xml$", n)]
    print("файлы диаграмм:", names)
    for name in sorted(names, key=lambda s: int(re.search(r"\d+", s).group())):
        xml = z.read(name).decode("utf-8")
        print(f"\n== {name} ==")
        # для каждой серии вывести tx (title), xVal и yVal
        for m in re.finditer(r"<c:ser>.*?</c:ser>", xml, re.S):
            ser = m.group(0)
            tx = re.search(r"<c:tx>.*?<c:v>(.*?)</c:v>.*?</c:tx>", ser, re.S)
            title = tx.group(1) if tx else "(без названия)"
            xv = re.search(r"<c:xVal><c:numRef><c:f>(.*?)</c:f>", ser)
            yv = re.search(r"<c:yVal><c:numRef><c:f>(.*?)</c:f>", ser)
            print(f"  серия {title!r}")
            print(f"    xVal: {xv.group(1) if xv else 'НЕТ'}")
            print(f"    yVal: {yv.group(1) if yv else 'НЕТ'}")
            # число кэшей в numRef
            xc = re.search(r"<c:xVal><c:numRef>.*?<c:ptCount val=\"(\d+)\"", ser, re.S)
            yc = re.search(r"<c:yVal><c:numRef>.*?<c:ptCount val=\"(\d+)\"", ser, re.S)
            print(f"    ptCount: x={xc.group(1) if xc else '?'}, y={yc.group(1) if yc else '?'}")
            # cached значения (первые 8)
            xcache = re.findall(r"<c:pt idx=\"\d+\"><c:v>([^<]*)</c:v>", 
                                (re.search(r"<c:xVal>.*?</c:xVal>", ser, re.S) or type("", (), {"group": ""})()).group(0) or "")
            if xcache:
                print(f"    xVal cached (первые 8): {xcache[:8]}")

print("\n--- 4. Где лежит серия точек в xVal (проверка гипотезы) ---")
with zipfile.ZipFile(LIVE) as z:
    name = "xl/charts/chart1.xml"
    xml = z.read(name).decode("utf-8")
    for m in re.finditer(r"<c:ser>.*?</c:ser>", xml, re.S):
        ser = m.group(0)
        tx = re.search(r"<c:tx>.*?<c:v>(.*?)</c:v>.*?</c:tx>", ser, re.S)
        title = tx.group(1) if tx else "(без названия)"
        xv = re.search(r"<c:xVal><c:numRef><c:f>(.*?)</c:f>", ser)
        print(f"chart1 серия {title!r}: xVal={xv.group(1) if xv else None}")

print("\nГотово.")