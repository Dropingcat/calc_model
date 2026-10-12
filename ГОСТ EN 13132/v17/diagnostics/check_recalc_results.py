# -*- coding: utf-8 -*-
"""check_recalc_results.py — проверка пересчитанной Excel-книги (v17, контракт fixes).

Читает v17\\diagnostics\\EN13132_v17_LIVE_RECALC.xlsx (openpyxl, data_only=True)
и проверяет:
  1) 02_Регрессия!A52:A151 — 100 УНИКАЛЬНЫХ значений (MIN→MAX, не константа);
  2) 06_Итог!A30 — текст «w(MTBE) = ...», НЕ #VALUE!;
  3) 06_Итог!B19/B20/B22/B25 — числа (X0, w, u_c, U);
  4) ВСЯ книга — ни одной ячейки с #DIV/0!, #REF!, #VALUE!, #N/A, #NAME?, #NUM!, #NULL!.
"""
from __future__ import annotations

import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
import openpyxl  # noqa: E402

V17_DIR = os.path.dirname(os.path.abspath(__file__))  # ...\v17\diagnostics
RECALC = os.path.join(V17_DIR, "EN13132_v17_LIVE_RECALC.xlsx")

ERR = ["#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NAME?", "#NUM!", "#NULL!"]

wb = openpyxl.load_workbook(RECALC, data_only=True)
print("Книга:", RECALC)
print("Листы:", wb.sheetnames)
print("=" * 70)

# ---- 1. Сетка X ----
ws = wb["02_Регрессия"]
vals = [ws.cell(r, 1).value for r in range(52, 152)]
n_ok = sum(1 for v in vals if isinstance(v, (int, float)))
uniq = len({round(float(v), 12) for v in vals if isinstance(v, (int, float))})
print("[1] 02_Регрессия!A52:A151:")
print("    числовых =", n_ok, "/ 100; уникальных (с 12 знаками) =", uniq)
for r in (52, 53, 75, 151):
    print("    A%d = %.6f" % (r, float(ws.cell(r, 1).value)))
mn = min(float(v) for v in vals if isinstance(v, (int, float)))
mx = max(float(v) for v in vals if isinstance(v, (int, float)))
print("    MIN = %.6f, MAX = %.6f" % (mn, mx))
grid_ok = (n_ok == 100 and uniq == 100 and mx > mn)
print("[%s] сетка X уникальна MIN→MAX" % ("PASS" if grid_ok else "FAIL"))

# ---- 2. A30 ----
ws6 = wb["06_Итог"]
a30 = ws6["A30"].value
a30_ok = isinstance(a30, str) and "w(MTBE)" in a30 and not str(a30).startswith("#")
print()
print("[2] 06_Итог!A30 =", repr(a30))
print("[%s] A30 — текст протокола без ошибки" % ("PASS" if a30_ok else "FAIL"))

# ---- 3. Ключевые числа ----
print()
print("[3] Ключевые числа 06_Итог:")
checks = [
    ("B19 X0", 0.9966, 0.01),
    ("B20 w", 3.967, 0.01),
    ("B22 u_c", 0.0402, 0.01),
    ("B25 U", 0.0798, 0.01),
]
key_ok = True
for label, ref, tol in checks:
    v = ws6[label.split()[0]].value
    ok = isinstance(v, (int, float)) and abs(v - ref) <= tol
    key_ok = key_ok and ok
    print("    06_Итог!%s = %r  [%s] (эталон ~%.4f)" % (label.split()[0], v, "PASS" if ok else "FAIL", ref))

# ---- 4. Скан ошибок по всей книге ----
print()
print("[4] Скан ошибок по ВСЕЙ книге:")
err_cells = []
for sname in wb.sheetnames:
    s = wb[sname]
    cnt = 0
    for row in s.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v in ERR:
                err_cells.append("%s!%s = %s" % (sname, c.coordinate, repr(v)))
                cnt += 1
    if cnt:
        print("   %s: ошибок %d" % (sname, cnt))
err_ok = len(err_cells) == 0
print("[%s] всего ячеек с ошибками = %d" % ("PASS" if err_ok else "FAIL", len(err_cells)))
if err_cells:
    for e in err_cells[:50]:
        print("    ", e)

print("=" * 70)
all_ok = grid_ok and a30_ok and key_ok and err_ok
print("ИТОГ:", "PASS" if all_ok else "FAIL")
sys.exit(0 if all_ok else 1)