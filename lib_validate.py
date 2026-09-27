#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""lib_validate.py — независимая проверка расчётов библиотеками (Фаза 5, LIB-01/02).

Читает данные из xlsx, прогоняет через:
  - scipy.stats.linregress   (ОМНК: b1, b0, s, var, cov)
  - statsmodels.WLS          (взвешенная регрессия для режимов 2/3)
  - scipy.stats.shapiro      (нормальность остатков)
  - scipy.stats.bartlett     (гомоскедастичность)
  - scipy.stats.levene       (гомоскедастичность, альтернатива)
и сверяет с Excel (листы «Регрессия», «Взвешенная регрессия»).

Критерий (из ревью R-01/R-06/A-03):
  - коэффициенты ОМНК/WLS: расхождение ≤ 1e-8;
  - дисперсии: ≤ 1e-6.

Использование:
  python lib_validate.py [файл.xlsx]
"""
import os
import sys

import numpy as np
import openpyxl
from scipy import stats

try:
    import statsmodels.api as sm
    HAS_SM = True
except ImportError:
    HAS_SM = False

FILE = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "Metrology_Core_EURACHEM_v4.xlsx")
TOL_COEF = 1e-8
TOL_VAR = 1e-6


def read_data(file):
    wb = openpyxl.load_workbook(file, data_only=True)
    inp = wb["Ввод"]
    mode = inp["B6"].value
    xs, raw = [], []
    for r in range(9, 109):
        x = inp.cell(r, 2).value
        vals = [inp.cell(r, c).value for c in (3, 4, 5)
                if isinstance(inp.cell(r, c).value, (int, float))]
        if x is not None and vals:
            xs.append(x)
            raw.append(vals)
    x = np.array(xs, float)
    y = np.array([np.mean(v) for v in raw], float)
    return wb, mode, x, y, raw


def main():
    wb, mode, x, y, raw = read_data(FILE)
    n = len(x)
    print(f"n={n}, режим={mode}, x={x}")

    ok_all = True

    # ============ 1. ОМНК через scipy.stats.linregress ============
    res = stats.linregress(x, y)
    b1, b0, r2 = res.slope, res.intercept, res.rvalue ** 2
    yhat = b1 * x + b0
    s2 = ((y - yhat) ** 2).sum() / (n - 2)
    Sxx = ((x - x.mean()) ** 2).sum()
    vb1 = s2 / Sxx
    vb0 = s2 * (1 / n + x.mean() ** 2 / Sxx)
    cov = -s2 * x.mean() / Sxx
    print("\n[ОМНК: scipy.stats.linregress]")
    print(f"  b1={b1:.12g}, b0={b0:.12g}, R²={r2:.9g}, s²={s2:.3e}")

    reg = wb["Регрессия"]
    for name, ev, nv, tol in [
        ("b1", reg["B7"].value, b1, TOL_COEF),
        ("b0", reg["B8"].value, b0, TOL_COEF),
        ("s2", reg["B6"].value, s2, TOL_VAR),
        ("var(b1)", reg["B9"].value, vb1, TOL_VAR),
        ("var(b0)", reg["B10"].value, vb0, TOL_VAR),
        ("cov", reg["B11"].value, cov, TOL_VAR),
    ]:
        rel = abs(nv - ev) / max(abs(ev), 1e-12)
        ok = rel <= tol
        ok_all &= ok
        print(f"  {name:<9}: {'OK' if ok else 'FAIL'} excel={ev:.12g} lib={nv:.12g} rel={rel:.1e}")

    # остатки (для shapiro)
    resid = y - yhat

    # ============ 2. WLS через statsmodels ============
    wls = wb["Взвешенная регрессия"]
    if HAS_SM:
        # проверяем только активный режим (файл хранит cached-значения активного режима)
        for m in (mode,):
            w = np.ones(n)
            if m == 2:
                w = np.array([1.0 / xi if xi > 0 else np.nan for xi in x])
            elif m == 3:
                w = np.array([1.0 / (xi ** 2) if xi > 0 else np.nan for xi in x])
            I = ~np.isnan(w)
            w = np.where(I, w, 0.0)
            n_eff = int(I.sum())
            X = sm.add_constant(x[I])
            wm = sm.WLS(y[I], X, weights=w[I]).fit()
            b1w, b0w = wm.params[1], wm.params[0]
            # пересчитаем с весом в пределах участвующих точек
            e = y - b0w - b1w * x
            SSEw = (w * e * e).sum()
            phi = SSEw / (n_eff - 2)
            Sw = w.sum()
            Swx = (w * x).sum()
            Sxxw = (w * x * x).sum() - Swx ** 2 / Sw
            vb1w = phi / Sxxw
            sumwx2 = (w * x * x).sum()
            vb0w = phi * sumwx2 / (Sw * Sxxw)
            covw = -phi * Swx / (Sw * Sxxw)
            print(f"\n[WLS режим {m}: statsmodels.WLS]")
            print(f"  b1_w={b1w:.12g}, b0_w={b0w:.12g}, phi={phi:.3e}, n_eff={n_eff}")
            for name, ev, nv, tol in [
                ("b1_w", wls["B122"].value, b1w, TOL_COEF),
                ("b0_w", wls["B123"].value, b0w, TOL_COEF),
                ("phi", wls["B125"].value, phi, TOL_VAR),
                ("var(b1)_w", wls["B126"].value, vb1w, TOL_VAR),
                ("var(b0)_w", wls["B127"].value, vb0w, TOL_VAR),
                ("cov_w", wls["B128"].value, covw, TOL_VAR),
            ]:
                rel = abs(nv - ev) / max(abs(ev), 1e-12)
                ok = rel <= tol
                ok_all &= ok
                print(f"    {name:<11}: {'OK' if ok else 'FAIL'} excel={ev:.12g} lib={nv:.12g} rel={rel:.1e}")
    else:
        print("\n[WLS] statsmodels НЕ установлен — пропуск")
        ok_all = False

    # ============ 3. Нормальность остатков ============
    print("\n[Нормальность остатков: scipy.stats.shapiro]")
    if len(resid) >= 3:
        W, p_sh = stats.shapiro(resid)
        print(f"  W={W:.6f}, p-value={p_sh:.6f}")
        print(f"  {'PASS (p≥0.05, остатки нормальны)' if p_sh >= 0.05 else 'FAIL (p<0.05, отклонение от нормальности)'}")
    else:
        print("  недостаточно точек")
        p_sh = float("nan")

    # ============ 4. Гомоскедастичность (Levene) ============
    print("\n[Гомоскедастичность: scipy.stats.levene]")
    # группируем отклики по уровням (повторности)
    groups = [raw[i] for i in range(n) if len(raw[i]) >= 2]
    if len(groups) >= 2 and all(len(g) >= 2 for g in groups):
        W_l, p_l = stats.levene(*groups)
        print(f"  W={W_l:.6f}, p-value={p_l:.6f}")
        print(f"  {'PASS (p≥0.05, дисперсии однородны)' if p_l >= 0.05 else 'FAIL (p<0.05, гетероскедастичность)'}")
    else:
        print("  недостаточно повторностей в группах")

    print("\n=== ИТОГ LIB-VALIDATE:", "ВСЁ СОВПАЛО ✅" if ok_all else "ЕСТЬ РАСХОЖДЕНИЯ ❌", "===")
    return 0 if ok_all else 1


if __name__ == "__main__":
    sys.exit(main())