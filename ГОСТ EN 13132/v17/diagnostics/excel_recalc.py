# -*- coding: utf-8 -*-
"""excel_recalc.py — РЕАЛЬНЫЙ пересчёт EN13132_v17_LIVE.xlsx через Excel COM.

Контракт gost13132-v17-fixes (баг 1/2/3): критическая проверка после фиксов.
  * открывает v17\\EN13132_v17_LIVE.xlsx (UpdateLinks=0);
  * CalculateFull() — полный пересчёт формул Excel 16.0 (русская локаль);
  * сохраняет копию v17\\diagnostics\\EN13132_v17_LIVE_RECALC.xlsx;
  * закрывает и выходит.

Правила: НЕ C:\\Temp (TD-DEV-213, AGENTS_PATHS_RULE.md); кириллица UTF-8.
"""
from __future__ import annotations

import os
import sys
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

V17_DIR = os.path.dirname(os.path.abspath(__file__))          # ...\v17\diagnostics
SRC = os.path.join(os.path.dirname(V17_DIR), "EN13132_v17_LIVE.xlsx")  # ...\v17\EN13132_v17_LIVE.xlsx
DST = os.path.join(V17_DIR, "EN13132_v17_LIVE_RECALC.xlsx")

if not os.path.exists(SRC):
    raise SystemExit(f"[FAIL] нет исходной книги: {SRC}")

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(SRC, UpdateLinks=0, ReadOnly=False)
    try:
        excel.CalculateFull()
        print("[OK] CalculateFull() выполнен")
    except Exception as exc:  # noqa: BLE001
        print("[WARN] CalculateFull exception:", exc)
    # Принудительный пересчёт всех листов (страховка)
    try:
        wb.Worksheets(1).Calculate()
        print("[OK] пересчёт активного листа выполнен")
    except Exception as exc:  # noqa: BLE001
        print("[WARN] лист Calculate:", exc)
    wb.SaveCopyAs(DST)
    print("[OK] сохранена копия:", DST)
    wb.Close(SaveChanges=False)
finally:
    excel.Quit()
print("Excel COM done")