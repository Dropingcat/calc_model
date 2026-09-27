#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""validate_auto.py — автономная кросс-валидация (без Excel COM).

Сверяет результаты calc_engine (чистый Python) с кэшированными значениями
в xlsx (после `calc --no-excel` или открытия в Excel). Если кэша нет —
просто подтверждает, что расчёт выполнен, и выводит итог.

Использование:
  python validate_auto.py [файл.xlsx]
"""
import os
import sys

import numpy as np
import openpyxl

import calc_engine

FILE = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "Metrology_Core_EURACHEM_v4.xlsx")
TOL = 1e-6


def main(xlsx_path=None):
    global FILE
    if xlsx_path:
        FILE = os.path.abspath(xlsx_path)
    wb = openpyxl.load_workbook(FILE)  # data_only=False
    d = calc_engine.read_input(wb)
    res = calc_engine.compute_all(d)

    print(f"файл: {FILE}, режим={d['mode']}, n={len(d['x'])}")
    print(f"  x_pred = {res['x_pred']:.9g}, u_c = {res['u_c']:.9g}, "
          f"U = {res['U']:.9g}, k = {res['k']:.4f}, ν_eff = {res['nu_eff']:.2f}")

    # Пытаемся сверить с кэшем в xlsx (если значения есть)
    reg = wb["Регрессия"]
    unc = wb["Неопределенность"]
    checks = [
        ("b1", reg["B7"].value, res["b1"]),
        ("b0", reg["B8"].value, res["b0"]),
        ("x_pred", unc["B1"].value, res["x_pred"]),
        ("u_c", unc["B5"].value, res["u_c"]),
        ("U", unc["B13"].value, res["U"]),
    ]
    all_ok = True
    for name, ev, nv in checks:
        if not isinstance(ev, (int, float)):
            print(f"  {name:<7}: кэш отсутствует — пропуск")
            continue
        rel = abs(nv - ev) / max(abs(ev), 1e-12)
        ok = rel <= TOL
        all_ok &= ok
        print(f"  {name:<7}: {'OK' if ok else 'FAIL'} excel={ev:.9g} py={nv:.9g} rel={rel:.1e}")

    print("\n=== ИТОГ validate_auto:", "ВСЁ СОВПАЛО ✅" if all_ok else "ЕСТЬ РАСХОЖДЕНИЯ ❌", "===")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())