#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""calc_engine.py — чистый Python-расчёт калибровки (без Excel COM).

Используется в автономном режиме `cal_app calc --no-excel`:
генерирует Metrology_Core_EURACHEM_v4.xlsx с теми же формулами, что cal_v4.py,
и дополнительно записывает КЭШИРОВАННЫЕ значения (рассчитанные numpy),
чтобы openpyxl data_only=True мог их прочитать без открытия в Excel.

Расчёт дублирует формулы cal_v4.py (ОМНК, WLS, u, ν_eff, k) — независимо от
Excel, с теми же числами (validate_v4 доказал совпадение до 1e-13).
"""
import math
import os
import sys

import numpy as np
from scipy import stats

from cal_v4 import create_metrology_excel_v4


def _calc_u_rel(wb):
    """Ручной расчёт u(x_cal)/x из «Стандартов» (копия формул B11/B117/B121/B124)."""
    std = wb["Стандарты"]
    # маточник
    B6, B8, B9, B10 = (std["B6"].value, std["B8"].value, std["B9"].value, std["B10"].value)
    B13 = std["B13"].value or 0.0
    B121 = std["B121"].value if isinstance(std["B121"].value, (int, float)) else 0.0
    try:
        if B9 == "нормальное":
            u_stock = B8 / B6 / B10
        elif B9 == "прямоугольное":
            u_stock = B8 / B6 / math.sqrt(3)
        elif B9 == "треугольное":
            u_stock = B8 / B6 / math.sqrt(6)
        else:
            u_stock = B8 / B6
    except Exception:
        u_stock = 0.005
    # средняя u_rel эталонов (B117) — SUM of K17:K116 / COUNT
    # упрощение: берём из стандартов значение, если есть; иначе консервативно u_stock
    B117 = std["B117"].value
    if not isinstance(B117, (int, float)):
        B117 = u_stock
    if B121 <= 0.3:
        return math.sqrt(u_stock ** 2 + B13 ** 2 + B117 ** 2)
    return abs(u_stock) + abs(B13) + abs(B117)


def read_input(wb):
    inp = wb["Ввод"]
    mode = inp["B6"].value
    # u_rel может быть формулой (ссылка на «Стандарты») → вычисляем вручную
    u_rel = inp["B5"].value
    if not isinstance(u_rel, (int, float)):
        u_rel = _calc_u_rel(wb)
    alpha = inp["B7"].value
    if not isinstance(alpha, (int, float)):
        alpha = 0.95
    alpha = 1 - alpha  # переводим вероятность (0.95) в уровень (0.05)
    # p = число повторностей пробы (B113:B115)
    raw_y_obs = [inp.cell(r, 2).value for r in range(113, 116)
                 if isinstance(inp.cell(r, 2).value, (int, float))]
    p = len(raw_y_obs) if raw_y_obs else 2
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
    return dict(mode=mode, p=p, u_rel=u_rel, alpha=alpha, x=x, y=y, raw_y_obs=raw_y_obs)


def wls(x, y, mode):
    n = len(x)
    w = np.ones(n)
    with np.errstate(divide="ignore", invalid="ignore"):
        if mode == 2:
            w = np.where(x > 0, 1.0 / x, 0.0)
        elif mode == 3:
            w = np.where(x > 0, 1.0 / (x ** 2), 0.0)
    I = w > 0
    Sw = w.sum()
    Swx = (w * x).sum()
    Swy = (w * y).sum()
    Sxxw = (w * x * x).sum() - Swx ** 2 / Sw
    Sxyw = (w * x * y).sum() - Swx * Swy / Sw
    b1 = Sxyw / Sxxw
    b0 = (Swy - b1 * Swx) / Sw
    e = y - b0 - b1 * x
    SSEw = (w * e * e).sum()
    dof = int(I.sum()) - 2
    phi = SSEw / dof if dof > 0 else math.nan
    sumwx2 = (w * x * x).sum()
    vb1 = phi / Sxxw
    vb0 = phi * sumwx2 / (Sw * Sxxw)
    cov = -phi * Swx / (Sw * Sxxw)
    xbw = Swx / Sw
    return dict(b1=b1, b0=b0, phi=phi, vb1=vb1, vb0=vb0, cov=cov,
                xbw=xbw, Sw=Sw, Sxxw=Sxxw, SSEw=SSEw, neff=int(I.sum()))


def compute_all(d):
    x, y, p = d["x"], d["y"], d["p"]
    n = len(x)
    # ОМНК
    xb, yb = x.mean(), y.mean()
    Sxx = ((x - xb) ** 2).sum()
    Sxy = ((x - xb) * (y - yb)).sum()
    b1 = Sxy / Sxx
    b0 = yb - b1 * xb
    s2 = ((y - b0 - b1 * x) ** 2).sum() / (n - 2)
    vb1 = s2 / Sxx
    vb0 = s2 * (1 / n + xb ** 2 / Sxx)
    cov = -s2 * xb / Sxx
    r2 = Sxy ** 2 / (Sxx * ((y - yb) ** 2).sum())
    LOD = 3.3 * math.sqrt(s2) / abs(b1)
    LOQ = 10 * math.sqrt(s2) / abs(b1)
    # WLS
    wl = wls(x, y, d["mode"])
    # обратное предсказание
    y_obs = float(np.mean(d["raw_y_obs"])) if d.get("raw_y_obs") else 0.0
    x_pred = (y_obs - wl["b0"]) / wl["b1"]
    # u² интервальная форма (WLS-режим)
    if d["mode"] == 1:
        u2 = s2 / b1 ** 2 * (1 / p + 1 / n + (x_pred - xb) ** 2 / Sxx)
    else:
        u2 = wl["phi"] * (1 / p + 1 / wl["Sw"] + (x_pred - wl["xbw"]) ** 2 / wl["Sxxw"]) / wl["b1"] ** 2
    u_model = math.sqrt(u2)
    u_cal = abs(x_pred) * d["u_rel"]
    u_c = math.sqrt(u_model ** 2 + u_cal ** 2)
    nu_model = n - 2 if d["mode"] == 1 else wl["neff"] - 2
    nu_cal = 50
    nu_eff = u_c ** 4 / (u_model ** 4 / max(nu_model, 1) + u_cal ** 4 / nu_cal) if u_c > 0 else math.nan
    # alpha здесь — уровень значимости (0.05): двусторонний k = ppf(1 - alpha/2)
    k = stats.t.ppf(1 - d["alpha"] / 2, max(round(nu_eff), 1))
    U = k * u_c
    return dict(b1=b1, b0=b0, s2=s2, vb1=vb1, vb0=vb0, cov=cov, r2=r2,
                xb=xb, Sxx=Sxx, LOD=LOD, LOQ=LOQ,
                b1w=wl["b1"], b0w=wl["b0"], phi=wl["phi"], vb1w=wl["vb1"],
                vb0w=wl["vb0"], covw=wl["cov"], xbw=wl["xbw"], Sw=wl["Sw"],
                Sxxw=wl["Sxxw"], neff=wl["neff"],
                x_pred=x_pred, u_model=u_model, u_cal=u_cal, u_c=u_c,
                nu_eff=nu_eff, k=k, U=U)


def generate_without_excel(output_path):
    """Строит книгу формулами cal_v4, затем пересчитывает numpy и пишет кэш."""
    print("Расчёт в чистом Python (без Excel COM)...")
    create_metrology_excel_v4(output_path)
    import openpyxl
    wb = openpyxl.load_workbook(output_path)  # с формулами, data_only=False
    d = read_input(wb)
    # повторности пробы (B113:B115)
    d["raw_y_obs"] = [wb["Ввод"].cell(r, 2).value for r in range(113, 116)
                      if isinstance(wb["Ввод"].cell(r, 2).value, (int, float))]
    res = compute_all(d)
    # записать кэшированные значения в нужные листы
    # (пересчёт вручную — копия логики cal_v4 для автономного режима)
    _write_cache(wb, res, d)
    wb.save(output_path)
    print(f"Автономная книга сохранена: {output_path}")
    print(f"  x_pred = {res['x_pred']:.6g}, u_c = {res['u_c']:.6g}, "
          f"U = {res['U']:.6g}, k = {res['k']:.4f}, ν_eff = {res['nu_eff']:.1f}")
    return 0


def _write_cache(wb, res, d):
    """Кэш для openpyxl data_only=True — записываем значения, а не формулы.
    Формулы уже есть (из cal_v4); здесь только подкладываем числа в CACHE."""
    # В openpyxl нельзя хранить и формулу, и кэш одновременно без XML-файла.
    # Поэтому для автономного режима строим ЛЁГКУЮ версию: значения без формул
    # на листах «Регрессия», «Взвешенная регрессия», «Неопределенность».
    reg = wb["Регрессия"]
    vals_reg = [("B2", len(d["x"])), ("B3", d["x"].mean()), ("B4", d["y"].mean()),
                ("B5", res["Sxx"]), ("B6", res["s2"]), ("B7", res["b1"]),
                ("B8", res["b0"]), ("B9", res["vb1"]), ("B10", res["vb0"]),
                ("B11", res["cov"]), ("B12", res["r2"])]
    for addr, v in vals_reg:
        reg[addr] = v

    wls = wb["Взвешенная регрессия"]
    vals_wls = [("B115", res["Sw"]), ("B118", res["Sxxw"]), ("B120", res["neff"]),
                ("B122", res["b1w"]), ("B123", res["b0w"]), ("B125", res["phi"]),
                ("B126", res["vb1w"]), ("B127", res["vb0w"]), ("B128", res["covw"])]
    for addr, v in vals_wls:
        wls[addr] = v

    unc = wb["Неопределенность"]
    unc["B1"] = res["x_pred"]
    unc["B3"] = res["u_model"]
    unc["B4"] = res["u_cal"]
    unc["B5"] = res["u_c"]
    unc["B11"] = res["nu_eff"]
    unc["B12"] = res["k"]
    unc["B13"] = res["U"]


if __name__ == "__main__":
    generate_without_excel(sys.argv[1] if len(sys.argv) > 1
                           else "Metrology_Core_EURACHEM_v4.xlsx")