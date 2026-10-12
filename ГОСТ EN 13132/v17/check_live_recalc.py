#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_live_recalc.py — пересчёт EN13132_v17_LIVE.xlsx библиотекой formulas
и сверка ключевых ячеек с эталоном ядра (step, TD-DEV-208).

Использование: python check_live_recalc.py [путь_к_книге]
Если путь не задан: берёт EN13132_v17_LIVE.xlsx рядом со скриптом;
при недоступности (блокировка Excel) — C:\\Temp\\EN13132_v17_LIVE_fixed_tmp.xlsx.
"""
from __future__ import annotations

import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")

import numpy as np  # noqa: E402
import formulas  # noqa: E402

V17_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT = os.path.join(V17_DIR, "EN13132_v17_LIVE.xlsx")
TMP = r"C:\Temp\EN13132_v17_LIVE_fixed_tmp.xlsx"

# Эталоны step (TD-DEV-208; соответствуют ядру anova_gum.py/build_en13132_v17_calc):
#   U(06_Итог!B25) = 0.07984086 г/г  (U = k·u_c, step)
#   a (02_Регрессия!B10) = 0.96099776; R² (B15) = 0.99974
#   X0 (06_Итог!B19) = 0.99662495 г/г; w0 (06_Итог!B20) = 3.9665673 % мас.
REF = {
    "06_Итог!B19": 0.99662495,
    "06_Итог!B20": 3.9665673,
    "06_Итог!B25": 0.07984086,
    "02_Регрессия!B10": 0.96099776,
    "02_Регрессия!B15": 0.99974,
}
RTOL = 1e-3

_KEY_RE = re.compile(r"^'?\[(?P<book>[^\]]+)\](?P<sheet>[^'!]+)'?!(?P<addr>[A-Z]+\d+)$")

# Блок gost13132-v17-chart-axes: таблица точек для графика (02_Регрессия
# A39:A50/B39:B50). ФИКС (Вариант А): для незаполненных уровней 7..12 формулы
# возвращают NA() вместо "" — иначе "" в xVal серии точек ScatterChart
# переключает ось X в категорийный режим 1..n («точки кривые, линия жива»).
# Проверяем, что в формулах A..F таблицы точек НЕТ пустых текстов "".
POINTS_FIRST = 39
POINTS_LAST = 50
POINTS_COLS = (1, 2, 3, 4, 5, 6)


def pick_path():
    if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
        return sys.argv[1]
    if os.path.exists(DEFAULT):
        return DEFAULT
    if os.path.exists(TMP):
        return TMP
    raise SystemExit("нет доступной книги: ни LIVE.xlsx, ни временной копии")


def main():
    path = pick_path()
    print(f"Книга: {path}")
    try:
        model = formulas.ExcelModel().loads(path).finish()
    except PermissionError:
        # канонический заблокирован — откат на временную копию
        if os.path.exists(TMP):
            print("Канонический LIVE заблокирован (Excel). Откат на временную копию.")
            path = TMP
            model = formulas.ExcelModel().loads(path).finish()
        else:
            raise
    except Exception:
        if os.path.exists(TMP):
            print("Не удалось прочитать канонический LIVE. Откат на временную копию.")
            path = TMP
            model = formulas.ExcelModel().loads(path).finish()
        else:
            raise
    res = model.calculate()

    values = {}
    for key, ranges in res.items():
        m = _KEY_RE.match(str(key))
        if not m:
            continue
        sheet = m.group("sheet").upper()
        addr = m.group("addr")
        values[(sheet, addr)] = ranges.value

    # --- Блок gost13132-v17-chart-axes: таблица точек без пустых текстов "" ---
    # (формулы A39:F50 должны возвращать NA() для незаполненных уровней, НЕ "")
    import openpyxl  # noqa: PLC0415

    wb_raw = openpyxl.load_workbook(path, data_only=False)
    ws_reg = wb_raw["02_Регрессия"]
    bad_blanks = []
    for r in range(POINTS_FIRST, POINTS_LAST + 1):
        for c in POINTS_COLS:
            v = ws_reg.cell(r, c).value
            if isinstance(v, str) and v.startswith("=") and ',"")' in v:
                bad_blanks.append(
                    f"{openpyxl.utils.get_column_letter(c)}{r}")
    ok_pts = not bad_blanks
    print(f"[{'PASS' if ok_pts else 'FAIL'}] V9: таблица точек A{POINTS_FIRST}:F{POINTS_LAST} "
          f"без \"\" в формулах (NA() для пустых уровней)"
          + (f" — найдено: {bad_blanks}" if bad_blanks else ""))
    all_ok = ok_pts

    print("=" * 60)
    all_ok = True
    for cell, ref in REF.items():
        sheet_name, addr = cell.split("!")
        key = (sheet_name.upper(), addr.upper())
        if key not in values:
            print(f"[FAIL] {cell} — формула не посчитана (None/отсутствует)")
            all_ok = False
            continue
        val = values[key]
        if val is None:
            print(f"[FAIL] {cell} — формула вернула None")
            all_ok = False
            continue
        try:
            val_f = float(np.asarray(val).reshape(-1)[0])
        except (TypeError, ValueError):
            print(f"[FAIL] {cell} — не число: {val!r}")
            all_ok = False
            continue
        ok = abs(val_f - ref) <= RTOL * abs(ref)
        tag = "PASS" if ok else "FAIL"
        if not ok:
            all_ok = False
        print(f"[{tag}] {cell} = {val_f:.10g} (эталон {ref:.10g}, rtol={RTOL})")

    print("=" * 60)
    print("ИТОГ:", "PASS" if all_ok else "FAIL")
    return 0 if all_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())