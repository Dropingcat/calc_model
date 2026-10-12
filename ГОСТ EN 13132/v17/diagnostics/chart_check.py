# -*- coding: utf-8 -*-
"""chart_check.py — РЕАЛЬНЫЙ Excel COM-пересчёт и проверка оси X серии точек.

Блок `gost13132-v17-chart-axes` (баг: «линия регрессии жива, а точки кривые —
ось X ломается, вместо значений Excel ставит 1,2,3,4,5,6»).

Что делает:
  1) открывает v17\\EN13132_v17_LIVE.xlsx через Excel COM (UpdateLinks=0);
  2) CalculateFull() — полный пересчёт формул (русская локаль);
  3) сохраняет копию v17\\diagnostics\\EN13132_v17_LIVE_RECALC2.xlsx;
  4) через openpyxl data_only=True проверяет:
       * серия точек: xVal = 6 ЧИСЛОВЫХ значений (0.026..3.14, НЕ 1..6);
       * кэш xVal в XML диаграммы chart1 (пересчитанный Excel-ом);
       * ВСЯ книга: ошибок #DIV/0!, #REF!, #VALUE! — 0; #N/A — ТОЛЬКО в
         сознательно введённых ячейках таблицы точек для незаполненных
         уровней 7..12 (02_Регрессия!A45:F50), больше нигде;
       * 06_Итог B19/B20/B22/B25 — посчитаны.

Правила (AGENTS_PATHS_RULE.md): временные файлы — ТОЛЬКО в v17\\diagnostics\\
и v17\\tmp\\; НЕ в C:\\Temp. Кириллица UTF-8.
"""
from __future__ import annotations

import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8")

import openpyxl  # noqa: E402
import win32com.client  # noqa: E402

DIAG = os.path.dirname(os.path.abspath(__file__))
V17_DIR = os.path.dirname(DIAG)
SRC = os.path.join(V17_DIR, "EN13132_v17_LIVE.xlsx")
DST = os.path.join(DIAG, "EN13132_v17_LIVE_RECALC2.xlsx")

EXPECT_X = [0.02618453865336658, 0.09184966216216219, 0.4847844463229078,
            0.9904679376083189, 2.22334429309535, 3.144321093082835]

# Путь книги можно передать аргументом (например, контрольная копия
# с корректными входами): python chart_check.py <path_to_xlsx>
if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
    SRC = os.path.abspath(sys.argv[1])
    base = os.path.splitext(os.path.basename(SRC))[0]
    DST = os.path.join(DIAG, base + "_RECALC2.xlsx")

print("=" * 70)
print("chart_check: Excel COM-пересчёт EN13132_v17_LIVE.xlsx")
print("=" * 70)

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
    wb.SaveCopyAs(DST)
    print("[OK] сохранена копия:", DST)
    wb.Close(SaveChanges=False)
finally:
    excel.Quit()
print("[OK] Excel закрыт\n")

# ---------------------------------------------------------------------------
# Проверка 1: 06_Итог B19/B20/B22/B25 + ключевые ячейки регрессии
# ---------------------------------------------------------------------------
wb_v = openpyxl.load_workbook(DST, data_only=True)
ws_it = wb_v["06_Итог"]
ws_reg = wb_v["02_Регрессия"]

key_cells = {
    "06_Итог!B19 (X0)": ws_it["B19"].value,
    "06_Итог!B20 (w0)": ws_it["B20"].value,
    "06_Итог!B22 (u_c)": ws_it["B22"].value,
    "06_Итог!B25 (U)": ws_it["B25"].value,
}
print("--- Проверка ключевых ячеек ---")
all_ok = True
for name, val in key_cells.items():
    ok = isinstance(val, (int, float))
    all_ok = all_ok and ok
    print(f"[{'PASS' if ok else 'FAIL'}] {name} = {val!r}")

# ---------------------------------------------------------------------------
# Проверка 2: серия точек xVal = A39:A50 после пересчёта
# ---------------------------------------------------------------------------
print("\n--- Серия точек: xVal (02_Регрессия!A39:A50, data_only) ---")
xvals = []
for r in range(39, 51):
    v = ws_reg.cell(r, 1).value
    xvals.append(v)
    if isinstance(v, (int, float)):
        print(f"A{r}: ЧИСЛО {v:.6f}")
    elif isinstance(v, str) and v.startswith("#"):
        print(f"A{r}: ОШИБКА {v}")
    else:
        print(f"A{r}: {v!r}")

nums = [v for v in xvals if isinstance(v, (int, float))]
ok_nums = len(nums) == 6 and all(
    abs(a - b) <= 1e-9 for a, b in zip(sorted(nums), sorted(EXPECT_X))
)
# НЕ 1..6 (категории) и НЕ 6 числовых равных 1..6
ok_not_cat = not (len(nums) == 6 and all(abs(nums[i] - (i + 1)) < 1e-9 for i in range(6)))
all_ok = all_ok and ok_nums and ok_not_cat
print(f"[{'PASS' if ok_nums else 'FAIL'}] xVal = 6 числовых значений (0.026..3.14): {sorted(round(v, 4) for v in nums)}")
print(f"[{'PASS' if ok_not_cat else 'FAIL'}] xVal НЕ категории 1..6")

# Пустых текстовых "" в диапазоне быть не должно
blank_txt = [f"A{r}" for r in range(39, 51)
             if isinstance(ws_reg.cell(r, 1).value, str) and ws_reg.cell(r, 1).value == ""]
ok_no_blank = not blank_txt
all_ok = all_ok and ok_no_blank
print(f"[{'PASS' if ok_no_blank else 'FAIL'}] в xVal нет пустых текстов \"\" {blank_txt or ''}")

# ---------------------------------------------------------------------------
# Проверка 3: кэш серии точек в XML диаграммы chart1 (пересчитанный Excel)
# ---------------------------------------------------------------------------
print("\n--- XML chart1.xml: кэш xVal серии «точки (X, Y_ср)» ---")
with zipfile.ZipFile(DST) as z:
    xml = z.read("xl/charts/chart1.xml").decode("utf-8")
for m in re.finditer(r"<c:ser>.*?</c:ser>", xml, re.S):
    ser = m.group(0)
    tx = re.search(r"<c:tx>.*?<c:v>(.*?)</c:v>.*?</c:tx>", ser, re.S)
    title = tx.group(1) if tx else "?"
    if "точки" not in title:
        continue
    xc = re.search(r"<c:xVal><c:numRef>.*?<c:ptCount val=\"(\d+)\"", ser, re.S)
    xcache = re.findall(r"<c:pt idx=\"\d+\"><c:v>([^<]*)</c:v>",
                        (re.search(r"<c:xVal>.*?</c:xVal>", ser, re.S) or
                         type("", (), {"group": lambda s, e: ""})()).group(0) or "")
    print(f"серия {title!r}: ptCount.x = {xc.group(1) if xc else '?'}")
    print(f"кэш xVal: {xcache}")
    # Excel кэширует ПОЛНЫЙ диапазон xVal (12 точек): первые 6 — реальные X
    # заполненных уровней, остальные — #N/A (Excel ПРОПУСКАЕТ их при отрисовке).
    # Важно: НЕ 1..6 (категории) и НЕТ пустых "" в кэше.
    try:
        cached_nums = [float(v) for v in xcache
                       if v not in ("#N/A", "#DIV/0!", "#REF!", "#VALUE!")]
        cached_na = [v for v in xcache if v == "#N/A"]
        no_blank = "" not in xcache
        ok_cache = (
            len(cached_nums) == 6
            and all(abs(a - b) <= 1e-6 for a, b in zip(sorted(cached_nums), sorted(EXPECT_X)))
            and len(cached_na) == 6
            and no_blank
        )
    except ValueError:
        ok_cache = False
    all_ok = all_ok and ok_cache
    print(f"[{'PASS' if ok_cache else 'FAIL'}] кэш xVal серии точек: 6 реальных X (0.026..3.14) "
          f"+ 6 #N/A (пропускаются), не 1..6")

# ---------------------------------------------------------------------------
# Проверка 4: ошибки по ВСЕЙ книге
# ---------------------------------------------------------------------------
print("\n--- Ошибки по всей книге ---")
ERR_PAT = re.compile(r"^#(DIV/0!|REF!|VALUE!|N/A|NAME\?|NUM!|NULL!)$")
hard_errs = []   # #DIV/0!, #REF!, #VALUE! — должны быть 0
na_errs = []     # #N/A — допускаются ТОЛЬКО в таблице точек незаполненных уровней
ALLOWED_NA = {(f"02_Регрессия", c, r) for r in range(45, 51) for c in range(1, 7)}
for sheet in wb_v.worksheets:
    for row in sheet.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and ERR_PAT.match(v):
                addr = (sheet.title, cell.column, cell.row)
                if v == "#N/A":
                    if addr in ALLOWED_NA:
                        continue
                    na_errs.append(f"{sheet.title}!{cell.coordinate}={v}")
                else:
                    hard_errs.append(f"{sheet.title}!{cell.coordinate}={v}")
ok_hard = not hard_errs
ok_na = not na_errs
all_ok = all_ok and ok_hard and ok_na
print(f"[{'PASS' if ok_hard else 'FAIL'}] #DIV/0!, #REF!, #VALUE! по всей книге: {len(hard_errs)}"
      + (f" — {hard_errs[:10]}" if hard_errs else ""))
print(f"[{'PASS' if ok_na else 'FAIL'}] #N/A вне таблицы точек (допустимы только A45:F50): {len(na_errs)}"
      + (f" — {na_errs[:10]}" if na_errs else ""))

# ---------------------------------------------------------------------------
print("\n" + "=" * 70)
print("ИТОГ:", "PASS" if all_ok else "FAIL")
sys.exit(0 if all_ok else 1)