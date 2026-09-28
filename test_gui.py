# -*- coding: utf-8 -*-
"""Тест логики GUI: создаём реальное Tk-окно (без mainloop), заполняем, считаем."""
import os, sys
os.environ["CAL_APP_BATCH"] = "1"
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
import gui

root = tk.Tk()
root.withdraw()  # скрыть окно, но корень есть — StringVar работает
app = gui.CalibGUI(root)
# Заполним поля вручную
app.var["mode"].delete(0, "end"); app.var["mode"].insert(0, "3")
app.var["alpha"].delete(0, "end"); app.var["alpha"].insert(0, "0.95")
yobs = [0.148, 0.152, 0.150]
for i in range(6):
    app.var[f"yobs{i}"].delete(0, "end")
    if i < len(yobs):
        app.var[f"yobs{i}"].insert(0, str(yobs[i]))
# калибровка
cal = [
    (0.0, 0.002, 0.001, 0.003),
    (0.2, 0.055, 0.058, 0.054),
    (0.4, 0.110, 0.112, 0.108),
    (0.6, 0.165, 0.162, 0.168),
    (0.8, 0.220, 0.225, 0.218),
    (1.0, 0.275, 0.272, 0.278),
]
for i, row in enumerate(cal):
    for c, v in enumerate(row):
        app.cal_vars[i][c].delete(0, "end")
        app.cal_vars[i][c].insert(0, str(v))
# стандарты
app.var["u_stock"].delete(0, "end"); app.var["u_stock"].insert(0, "0.005")
app.var["u_dil"].delete(0, "end"); app.var["u_dil"].insert(0, "0.005")
app.var["r_corr"].delete(0, "end"); app.var["r_corr"].insert(0, "0.9")
# бюджет
app.var["rec"].delete(0, "end"); app.var["rec"].insert(0, "0.9")
app.var["u_rec"].delete(0, "end"); app.var["u_rec"].insert(0, "0.048")
app.var["dil"].delete(0, "end"); app.var["dil"].insert(0, "1")
app.var["u_dil_d"].delete(0, "end"); app.var["u_dil_d"].insert(0, "0")

rc = app.calculate()
print("calculate rc:", rc)
print("статус:", app.status.get())
print("\n--- результат ---")
print(app.res_text.get("1.0", "end")[:600])
assert rc == 0, "Расчёт не удался"
print("\n=== GUI-логика OK ===")