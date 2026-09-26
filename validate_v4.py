#!/usr/bin/env python3
"""Кросс-валидация cal_v4.xlsx (TD-06) через COM + numpy.

Для КАЖДОГО режима весов (1=ОМНК, 2=1/x, 3=1/x²):
  - переключает B6 на «Вводе», пересчитывает книгу;
  - читает значения Excel через COM;
  - считает независимый numpy-эталон;
  - сверяет коэффициенты WLS, SSE, φ, var/cov, x_pred, u², u_c, ν_eff, k.

Отличие от openpyxl-подхода: читаются РЕАЛЬНЫЕ пересчитанные значения,
а не stale-cache. Расхождение ≤ TOL → PASS.
"""
import math
import os
import sys
import time

import numpy as np
import win32com.client
from scipy import stats

FILE = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "Metrology_Core_EURACHEM_v4.xlsx")
TOL = 1e-5  # относительный допуск на коэффициенты (Excel fp vs numpy)

MODES = {1: "ОМНК/равные", 2: "1/x", 3: "1/x²"}


def t_inv_2t(alpha, df):
    return stats.t.ppf(1 - alpha / 2, max(df, 1))


def main():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(FILE, UpdateLinks=0, ReadOnly=True)
        inp = wb.Sheets("Ввод")
        wls = wb.Sheets("Взвешенная регрессия")
        unc = wb.Sheets("Неопределенность")

        y_obs = inp.Cells(3, 2).Value
        p = inp.Cells(4, 2).Value
        u_rel = inp.Cells(5, 2).Value
        alpha = 1 - inp.Cells(7, 2).Value

        # данные с «Ввода»: сырые отклики C..E -> среднее
        xs, ys = [], []
        for r in range(9, 109):
            x = inp.Cells(r, 2).Value
            vals = [inp.Cells(r, c).Value for c in (3, 4, 5)]
            vals = [v for v in vals if isinstance(v, (int, float))]
            if x is not None and vals:
                xs.append(x)
                ys.append(sum(vals) / len(vals))
        x = np.array(xs, float)
        y = np.array(ys, float)
        n = len(x)
        print(f"n={n}, x={x}, y={y}")

        # ===== ОМНК (не зависит от режима) =====
        xb, yb = x.mean(), y.mean()
        Sxx = ((x - xb) ** 2).sum()
        Sxy = ((x - xb) * (y - yb)).sum()
        b1 = Sxy / Sxx
        b0 = yb - b1 * xb
        r2 = Sxy ** 2 / (Sxx * ((y - yb) ** 2).sum())
        s2 = ((y - b0 - b1 * x) ** 2).sum() / (n - 2)
        vb1 = s2 / Sxx
        vb0 = s2 * (1 / n + xb ** 2 / Sxx)
        cov = -s2 * xb / Sxx
        reg = wb.Sheets("Регрессия")
        ols = {"b1": b1, "b0": b0, "s2": s2, "vb1": vb1, "vb0": vb0, "cov": cov}
        ols_ok = True
        for i, (name, nv) in enumerate(ols.items()):
            ev = reg.Cells(7 if name == "b1" else
                           8 if name == "b0" else
                           6 if name == "s2" else
                           9 if name == "vb1" else
                           10 if name == "vb0" else 11, 2).Value
            ok = abs(nv - ev) <= TOL * max(1, abs(ev))
            ols_ok &= ok
            print(f"  ОМНК {name}: {'OK' if ok else 'FAIL'} excel={ev:.9g} numpy={nv:.9g}")

        all_ok = ols_ok

        # ===== WLS по режимам =====
        for mode in (1, 2, 3):
            inp.Cells(6, 2).Value = mode
            excel.CalculateFull()
            time.sleep(0.2)

            # веса как в Excel
            w = np.ones(n)
            if mode == 2:
                w = np.array([1.0 / xi if xi > 0 else np.nan for xi in x])
            elif mode == 3:
                w = np.array([1.0 / (xi ** 2) if xi > 0 else np.nan for xi in x])
            I = ~np.isnan(w)
            w = np.where(I, w, 0.0)
            n_eff = int(I.sum())

            Sw = w.sum()
            Swx = (w * x).sum()
            Swy = (w * y).sum()
            Sxxw = (w * x * x).sum() - Swx ** 2 / Sw
            Sxyw = (w * x * y).sum() - Swx * Swy / Sw
            b1w = Sxyw / Sxxw
            b0w = (Swy - b1w * Swx) / Sw
            e = y - b0w - b1w * x
            SSEw = (w * e * e).sum()
            dof = n_eff - 2
            phi = SSEw / dof
            vb1w = phi / Sxxw
            sumw2x2 = (w * w * x * x).sum()
            vb0w = phi * sumw2x2 / (Sw * Sxxw)
            covw = -phi * Swx / (Sw * Sxxw)
            xbw = Swx / Sw

            # x_pred и неопределённость
            x0 = (y_obs - b0w) / b1w
            if mode == 1:
                # классическая интервальная форма EURACHEM
                u2 = s2 / b1w ** 2 * (1 / p + 1 / n + (x0 - xb) ** 2 / Sxx)
            else:
                # ковариационная форма с учётом весов (WLS)
                dx = x0 - xbw + b0w / b1w
                u2 = (phi * (1 / p + 1 / Sw) + dx ** 2 * vb1w + vb0w
                      + 2 * dx * covw) / b1w ** 2
            u2_cal = (x0 * u_rel) ** 2
            u2_c = u2 + u2_cal
            nu_model = dof
            nu_cal = 50
            nu_eff = (u2_c ** 2 / (u2 ** 2 / nu_model + u2_cal ** 2 / nu_cal)
                      if u2_c > 0 else np.nan)
            k = t_inv_2t(alpha, max(round(nu_eff), 1))

            # чтение Excel
            row_of = {"Sw": 115, "Swx": 116, "Sxxw": 118, "Sxyw": 119, "n_eff": 120,
                      "b1w": 122, "b0w": 123, "SSEw": 124, "phi": 125,
                      "vb1w": 126, "vb0w": 127, "covw": 128, "sumw2x2": 132}
            exp = {"Sw": Sw, "Swx": Swx, "Sxxw": Sxxw, "Sxyw": Sxyw, "n_eff": n_eff,
                   "b1w": b1w, "b0w": b0w, "SSEw": SSEw, "phi": phi,
                   "vb1w": vb1w, "vb0w": vb0w, "covw": covw, "sumw2x2": sumw2x2}
            e_x0 = unc.Cells(1, 2).Value
            e_u2 = unc.Cells(2, 2).Value
            e_uc = unc.Cells(5, 2).Value
            e_nu = unc.Cells(11, 2).Value
            e_k = unc.Cells(12, 2).Value

            print(f"\n=== WLS режим {mode} ({MODES[mode]}) ===")
            mode_ok = True
            for name, nv in exp.items():
                ev = wls.Cells(row_of[name], 2).Value
                ok = abs(nv - ev) <= TOL * max(1, abs(ev))
                mode_ok &= ok
                print(f"    {name:<8}: {'OK' if ok else 'FAIL'} excel={ev:.9g} numpy={nv:.9g}")
            for name, ev, nv in [("x_pred", e_x0, x0), ("u2_model", e_u2, u2),
                                 ("u_c", e_uc, math.sqrt(u2_c)), ("nu_eff", e_nu, nu_eff),
                                 ("k", e_k, k)]:
                ok = abs(nv - ev) <= TOL * max(1, abs(ev))
                mode_ok &= ok
                print(f"    {name:<8}: {'OK' if ok else 'FAIL'} excel={ev:.9g} numpy={nv:.9g}")
            all_ok &= mode_ok

        print("\n=== ИТОГ CROSS-VALIDATION:", "ВСЁ СОВПАЛО ✅" if all_ok else "ЕСТЬ РАСХОЖДЕНИЯ ❌", "===")
        return 0 if all_ok else 1

    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        excel.Quit()


if __name__ == "__main__":
    sys.exit(main())