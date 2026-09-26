# -*- coding: utf-8 -*-
"""TD-07 test: добавление точки в конец калибровки — модель должна адаптироваться."""
import os, time
import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(os.path.abspath("Metrology_Core_EURACHEM_v4.xlsx"), UpdateLinks=0, ReadOnly=True)
    inp = wb.Sheets("Ввод")
    # добавляем 7-ю точку x=1.2, y ~ 0.33 (идеальная прямая b1=0.275)
    inp.Cells(15, 2).Value = 1.2
    inp.Cells(15, 3).Value = 0.330
    inp.Cells(15, 4).Value = 0.328
    inp.Cells(15, 5).Value = 0.332
    excel.CalculateFull()
    time.sleep(0.5)

    reg = wb.Sheets("Регрессия")
    n = reg.Cells(2, 2).Value
    b1 = reg.Cells(7, 2).Value
    b0 = reg.Cells(8, 2).Value
    r2 = reg.Cells(12, 2).Value
    print(f"n={n}, b1={b1:.6f}, b0={b0:.6f}, R2={r2:.6f}")
    print(f"Ожидание: n=7, b1~0.275, b0~0.001")
    ok = (n == 7) and (abs(b1 - 0.275) < 0.01) and (abs(b0) < 0.01)
    print("TD-07 (динамический диапазон):", "OK" if ok else "FAIL")
    wb.Close(False)
finally:
    excel.Quit()