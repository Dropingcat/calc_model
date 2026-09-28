#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gui.py — графический интерфейс (tkinter) для Metrology Core EURACHEM.

Вкладки:
  Базовые   — режим весов, доверительная вероятность, повторности пробы
  Калибровка — таблица точек: x, y1, y2, y3
  Стандарты — u(x_cal)/x (маточник, разведение), Recovery
  Бюджет    — Recovery, разбавление, навеска/объём/неоднородность
  Результат — x_pred ± U, k, ν_eff, кнопки «Расчёт» и «Отчёт Word»

Данные записываются в зелёные ячейки xlsx (тот же формат, что ручной ввод),
затем считается calc_engine (автономно, без Excel COM).
"""
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox

import numpy as np

import calc_engine

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "Metrology_Core_EURACHEM_v4.xlsx")
DOCX = os.path.join(HERE, "Отчёт_по_калибровке.docx")

DEFAULT_CAL = [
    (0.0, 0.002, 0.001, 0.003),
    (0.2, 0.055, 0.058, 0.054),
    (0.4, 0.110, 0.112, 0.108),
    (0.6, 0.165, 0.162, 0.168),
    (0.8, 0.220, 0.225, 0.218),
    (1.0, 0.275, 0.272, 0.278),
]


class CalibGUI:
    def __init__(self, root):
        self.root = root
        root.title("Metrology Core EURACHEM — калибровка")
        root.geometry("900x650")

        # Строка состояния
        self.status = tk.StringVar(value="Готов. Заполните данные и нажмите «Рассчитать».")

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=6)

        self.var = {}

        self.tab_basic()
        self.tab_calib()
        self.tab_std()
        self.tab_budget()
        self.tab_result()

        # Нижняя панель
        bar = ttk.Frame(root)
        bar.pack(fill="x", padx=6, pady=4)
        ttk.Label(bar, textvariable=self.status, foreground="#1a5").pack(side="left")
        ttk.Button(bar, text="Сохранить и рассчитать", command=self.calculate).pack(side="right")
        ttk.Button(bar, text="Отчёт Word", command=self.make_report).pack(side="right", padx=4)

        # Загрузить текущие данные из xlsx (если есть)
        self.load_from_xlsx()

    # ---------- Вкладки ----------
    def _label_entry(self, parent, row, col, text, key, width=12):
        ttk.Label(parent, text=text).grid(row=row, column=col, sticky="w", padx=4, pady=2)
        e = ttk.Entry(parent, width=width)
        e.grid(row=row, column=col + 1, sticky="w", padx=4, pady=2)
        self.var[key] = e
        return e

    def tab_basic(self):
        t = ttk.Frame(self.notebook)
        self.notebook.add(t, text=" Базовые ")
        ttk.Label(t, text="Параметры расчёта", font=("Arial", 11, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", padx=6, pady=6)
        self._label_entry(t, 1, 0, "Режим весов (1=равные, 2=1/x, 3=1/x²)", "mode", 6)
        self._label_entry(t, 2, 0, "Доверительная вероятность (0.95)", "alpha", 6)
        ttk.Label(t, text="Повторности пробы y_obs (до 6):", font=("Arial", 10)).grid(row=3, column=0, columnspan=4, sticky="w", padx=6, pady=6)
        for i in range(6):
            self._label_entry(t, 4 + i // 2, (i % 2) * 2, f"y_obs[{i + 1}]", f"yobs{i}", 10)

    def tab_calib(self):
        t = ttk.Frame(self.notebook)
        self.notebook.add(t, text=" Калибровка ")
        ttk.Label(t, text="Калибровочные точки (x, y1, y2, y3). Пустая строка x — конец.", font=("Arial", 10)).grid(row=0, column=0, columnspan=6, sticky="w", padx=6, pady=6)
        hdr = ("№", "x", "y1", "y2", "y3")
        for c, h in enumerate(hdr):
            ttk.Label(t, text=h, font=("Arial", 9, "bold")).grid(row=1, column=c, padx=2)
        self.cal_vars = []
        for i in range(12):
            row_vars = []
            ttk.Label(t, text=str(i + 1)).grid(row=2 + i, column=0, padx=2)
            for c in range(1, 5):
                e = ttk.Entry(t, width=10)
                e.grid(row=2 + i, column=c, padx=2, pady=1)
                row_vars.append(e)
            self.cal_vars.append(row_vars)

    def tab_std(self):
        t = ttk.Frame(self.notebook)
        self.notebook.add(t, text=" Стандарты ")
        ttk.Label(t, text="Неопределённость эталонов u(x_cal)/x", font=("Arial", 11, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", padx=6, pady=6)
        self._label_entry(t, 1, 0, "u_rel маточника (из сертификата)", "u_stock", 8)
        self._label_entry(t, 2, 0, "u_rel разведения (пипетки/колбы)", "u_dil", 8)
        self._label_entry(t, 3, 0, "Коэфф. корреляции r (0..1)", "r_corr", 6)
        ttk.Label(t, text="(итог u(x_cal)/x считается по тем же правилам, что лист «Стандарты»)", font=("Arial", 9), foreground="#777").grid(row=4, column=0, columnspan=4, sticky="w", padx=6, pady=4)

    def tab_budget(self):
        t = ttk.Frame(self.notebook)
        self.notebook.add(t, text=" Бюджет ")
        ttk.Label(t, text="Компоненты бюджета неопределённости (отн. u/x)", font=("Arial", 11, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", padx=6, pady=6)
        self._label_entry(t, 1, 0, "Recovery Rec (0..1, 1=100%)", "rec", 8)
        self._label_entry(t, 2, 0, "u(Rec)/Rec", "u_rec", 8)
        self._label_entry(t, 3, 0, "Разбавление d", "dil", 8)
        self._label_entry(t, 4, 0, "u(d)/d", "u_dil_d", 8)
        self._label_entry(t, 5, 0, "Навеска u(m)/m", "u_m", 8)
        self._label_entry(t, 6, 0, "Объём u(V)/V", "u_v", 8)
        self._label_entry(t, 7, 0, "Неоднородность u(H)/H", "u_h", 8)

    def tab_result(self):
        t = ttk.Frame(self.notebook)
        self.notebook.add(t, text=" Результат ")
        ttk.Label(t, text="Результат измерения", font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", padx=6, pady=8)
        self.res_text = tk.Text(t, width=90, height=22, font=("Consolas", 10))
        self.res_text.grid(row=1, column=0, columnspan=3, padx=6, pady=4)
        self.res_text.insert("1.0", "Нажмите «Сохранить и рассчитать»...")

    # ---------- Логика ----------
    def _g(self, key, default=0.0):
        try:
            return float(self.var[key].get())
        except Exception:
            return default

    def _read_cal(self):
        rows = []
        for row in self.cal_vars:
            vals = [e.get().strip() for e in row]
            if not vals[0]:
                continue
            try:
                x = float(vals[0])
                ys = [float(v) if v else float("nan") for v in vals[1:]]
                rows.append((x, ys))
            except ValueError:
                pass
        return rows

    def save_to_xlsx(self):
        """Записывает данные GUI в зелёные ячейки xlsx (формат cal_v4)."""
        import openpyxl
        # Создаём свежую книгу формулами (как в Excel-ветке), затем заполняем ввод
        from cal_v4 import create_metrology_excel_v4
        create_metrology_excel_v4(XLSX)
        wb = openpyxl.load_workbook(XLSX)
        inp = wb["Ввод"]
        # режим и вероятность
        inp["B6"] = int(self._g("mode", 3))
        inp["B7"] = self._g("alpha", 0.95)
        # повторности пробы
        yobs = [self._g(f"yobs{i}") for i in range(6)]
        yobs = [y for y in yobs if y != 0.0]
        for i in range(6):
            inp.cell(113 + i, 2, yobs[i] if i < len(yobs) else "")
        # калибровочные точки
        cal = self._read_cal()
        for r in range(9, 109):
            idx = r - 9
            if idx < len(cal):
                x, ys = cal[idx]
                inp.cell(r, 2, x)
                for c, v in enumerate(ys):
                    inp.cell(r, 3 + c, v if not np.isnan(v) else "")
            else:
                inp.cell(r, 2, "")
                for c in range(3, 6):
                    inp.cell(r, c, "")
        # Стандарты: маточник
        std = wb["Стандарты"]
        std["B11"] = self._g("u_stock", 0.005)
        std["B13"] = self._g("u_dil", 0.005)
        std["B121"] = self._g("r_corr", 0.9)
        # Бюджет
        bud = wb["Бюджет"]
        bud["B5"] = self._g("rec", 0.9)
        bud["B6"] = self._g("u_rec", 0.048)
        bud["B11"] = self._g("dil", 1.0)
        bud["B12"] = self._g("u_dil_d", 0.0)
        wb.save(XLSX)
        return XLSX

    def calculate(self):
        try:
            self.save_to_xlsx()
            import openpyxl
            wb = openpyxl.load_workbook(XLSX)  # data_only=False, свежие значения ввода
            d = calc_engine.read_input(wb)
            res = calc_engine.compute_all(d)
            self.last_res = res
            # Вывод
            txt = []
            txt.append("=== РЕЗУЛЬТАТ ===\n")
            txt.append(f"x_pred = {res['x_pred']:.6f}  ±  U = {res['U']:.6f}")
            txt.append(f"  k = {res['k']:.3f}, ν_eff = {res['nu_eff']:.1f}, P = 0.95")
            txt.append(f"  u_c = {res['u_c']:.6f}")
            txt.append(f"  LOD = {res['LOD']:.6f}, LOQ = {res['LOQ']:.6f}")
            txt.append(f"\nКоэффициенты:")
            txt.append(f"  ОМНК: b1 = {res['b1']:.9g}, b0 = {res['b0']:.9g}, R² = {res['r2']:.6f}")
            if d["mode"] != 1:
                txt.append(f"  WLS: b1_w = {res['b1w']:.9g}, b0_w = {res['b0w']:.9g}")
            txt.append(f"\nВклад эталонов u(x_cal)/x = {d['u_rel']:.4%}")
            txt.append(f"\nБюджет (u_rel): калибровка={res['u_model']/res['x_pred']:.4%}, "
                       f"Recovery={self._g('u_rec',0.048):.4%}")
            self.res_text.delete("1.0", "end")
            self.res_text.insert("1.0", "\n".join(txt))
            self.status.set("Расчёт выполнен. Результат во вкладке «Результат».")
            return 0
        except Exception as e:
            if not os.environ.get("CAL_APP_BATCH"):
                messagebox.showerror("Ошибка", str(e))
            self.status.set(f"Ошибка: {e}")
            return 1

    def make_report(self):
        try:
            if not hasattr(self, "last_res"):
                if self.calculate() != 0:
                    return
            import report_v4
            d = report_v4.read_from_py(XLSX)
            report_v4.build_report(d, DOCX)
            self.status.set(f"Отчёт сохранён: {DOCX}")
            if not os.environ.get("CAL_APP_BATCH"):
                messagebox.showinfo("Отчёт", f"Отчёт создан:\n{DOCX}")
        except Exception as e:
            if not os.environ.get("CAL_APP_BATCH"):
                messagebox.showerror("Ошибка", str(e))

    def load_from_xlsx(self):
        """Загружает текущие значения из xlsx в поля (если файл есть и читается)."""
        import openpyxl
        if not os.path.exists(XLSX):
            return
        try:
            wb = openpyxl.load_workbook(XLSX)
            inp = wb["Ввод"]
            self.var["mode"].delete(0, "end"); self.var["mode"].insert(0, str(inp["B6"].value or 3))
            self.var["alpha"].delete(0, "end"); self.var["alpha"].insert(0, str(inp["B7"].value or 0.95))
            for i in range(6):
                v = inp.cell(113 + i, 2).value
                self.var[f"yobs{i}"].delete(0, "end")
                if v is not None:
                    self.var[f"yobs{i}"].insert(0, str(v))
            # калибровка
            for i in range(12):
                r = 9 + i
                x = inp.cell(r, 2).value
                if x is None:
                    continue
                self.cal_vars[i][0].delete(0, "end"); self.cal_vars[i][0].insert(0, str(x))
                for c in range(1, 4):
                    v = inp.cell(r, 2 + c).value
                    self.cal_vars[i][c].delete(0, "end")
                    if v is not None:
                        self.cal_vars[i][c].insert(0, str(v))
            std = wb["Стандарты"]
            self.var["u_stock"].delete(0, "end"); self.var["u_stock"].insert(0, str(std["B11"].value or 0.005))
            self.var["u_dil"].delete(0, "end"); self.var["u_dil"].insert(0, str(std["B13"].value or 0.005))
            self.var["r_corr"].delete(0, "end"); self.var["r_corr"].insert(0, str(std["B121"].value or 0.9))
            bud = wb["Бюджет"]
            self.var["rec"].delete(0, "end"); self.var["rec"].insert(0, str(bud["B5"].value or 0.9))
            self.var["u_rec"].delete(0, "end"); self.var["u_rec"].insert(0, str(bud["B6"].value or 0.048))
        except Exception:
            pass


def main():
    if os.environ.get("CAL_APP_BATCH"):
        # Заглушка для тестов без реального GUI-цикла
        return 0
    root = tk.Tk()
    app = CalibGUI(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())