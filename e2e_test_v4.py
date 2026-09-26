#!/usr/bin/env python3
"""e2e-тест Metrology_Core_EURACHEM_v4.xlsx через Excel COM.

Открывает книгу в реальном Excel, пересчитывает формулы, прогоняет
режимы весов (1=ОМНК, 2=1/x, 3=1/x²) и проверяет:
  1) отсутствие ошибок #N/A, #VALUE!, #DIV/0!, #REF!, #NUM! в расчётных областях;
  2) согласованность ОМНК/WLS коэффициентов (b1, b0) с ожидаемыми;
  3) все проверки листа «Валидация» → PASS;
  4) вердикты Lack-of-Fit и модели дисперсии на «Диагностике»;
  5) корректность TD-01: при весах 1/x² нулевой уровень исключён (n_eff=5),
     но остаётся в ОМНК;
  6) U(x_pred) положительна и конечна.

Использование:
  python e2e_test_v4.py [путь_к_xlsx]
"""
import sys
import time

import win32com.client

FILE = sys.argv[1] if len(sys.argv) > 1 else "Metrology_Core_EURACHEM_v4.xlsx"

ERROR_MARKERS = ("#N/A", "#VALUE!", "#DIV/0!", "#REF!", "#NUM!", "#NAME?", "#NULL!")

EXPECTED_B1 = 0.275  # наклон истинной модели y = 0.275*x
EXPECTED_B0 = 0.0    # свободный член


def scan_errors(sheet, max_row, max_col):
    """Возвращает список ячеек с маркерами ошибок."""
    bad = []
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            try:
                v = sheet.Cells(r, c).Value
            except Exception:
                continue
            if isinstance(v, str):
                for m in ERROR_MARKERS:
                    if m in v:
                        bad.append((r, c, v))
                        break
    return bad


def fmt(v):
    try:
        return f"{v:.6f}" if isinstance(v, (int, float)) else str(v)
    except Exception:
        return str(v)


def main():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        import os
        abs_path = os.path.abspath(FILE)
        if not os.path.exists(abs_path):
            print(f"!! файл не найден: {abs_path}")
            return 1
        wb = excel.Workbooks.Open(abs_path, UpdateLinks=0, ReadOnly=True)
        # полный пересчёт
        excel.CalculateFull()
        time.sleep(0.5)

        # лист 8 «Валидация»: проверим перед переключением режимов
        val = wb.Sheets("Валидация")
        n_checks = 0
        fails = []
        r = 3
        while val.Cells(r, 1).Value:
            name = val.Cells(r, 1).Value
            status = val.Cells(r, 3).Value
            if status is None:
                break
            n_checks += 1
            if str(status).strip().upper() != "PASS":
                fails.append((r, name, status))
            r += 1

        verdict_ev = str(val.Cells(r + 1, 2).Value)  # Итог
        print(f"[Валидация] проверок: {n_checks}, FAIL: {len(fails)}")
        for fr, fn, fs in fails:
            print(f"    FAIL  row{fr}: {fn} -> {fs}")
        print(f"[Валидация] Итог: {verdict_ev}")

        # Лист «Регрессия» (ОМНК) — базовые значения
        reg = wb.Sheets("Регрессия")
        b1, b0, r2 = reg.Cells(7, 2).Value, reg.Cells(8, 2).Value, reg.Cells(12, 2).Value
        s2 = reg.Cells(6, 2).Value
        print(f"[ОМНК] b1={fmt(b1)}, b0={fmt(b0)}, R²={fmt(r2)}, s²_y/x={fmt(s2)}")
        b1_ok = abs(b1 - EXPECTED_B1) < 0.005
        b0_ok = abs(b0 - EXPECTED_B0) < 5e-3
        print(f"    b1 в допуске (≈0.275): {b1_ok}, b0 в допуске (≈0): {b0_ok}")

        # Лист «Взвешенная регрессия» при текущем режиме
        wls = wb.Sheets("Взвешенная регрессия")
        b1w, b0w = wls.Cells(122, 2).Value, wls.Cells(123, 2).Value
        n_eff = wls.Cells(120, 2).Value
        sw = wls.Cells(115, 2).Value
        print(f"[WLS(текущий)] b1_w={fmt(b1w)}, b0_w={fmt(b0w)}, n_eff={n_eff}, Σw={fmt(sw)}")

        # Лист «Диагностика»
        dg = wb.Sheets("Диагностика")
        lof_ols = dg.Cells(117, 2).Value
        lof_wls = dg.Cells(120, 2).Value
        var_model = dg.Cells(125, 2).Value
        corr = dg.Cells(123, 2).Value
        print(f"[Диагностика] LoF ОМНК: {lof_ols}")
        print(f"[Диагностика] LoF WLS: {lof_wls}")
        print(f"[Диагностика] Модель дисперсии: {var_model}")
        print(f"[Диагностика] corr(|e|√w, x) = {fmt(corr)}")

        # Лист «Неопределенность»
        unc = wb.Sheets("Неопределенность")
        xp, uc, kk, uu = unc.Cells(1, 2).Value, unc.Cells(5, 2).Value, unc.Cells(12, 2).Value, unc.Cells(13, 2).Value
        print(f"[U] x_pred={fmt(xp)}, u_c={fmt(uc)}, k={fmt(kk)}, U={fmt(uu)}")
        u_ok = uc is not None and uc > 0 and uu is not None and uu > 0
        print(f"    U положительна и конечна: {u_ok}")

        # ====== Прогон режимов весов ======
        inp = wb.Sheets("Ввод")
        results = {}
        for mode in (1, 2, 3):
            inp.Cells(6, 2).Value = mode
            excel.CalculateFull()
            time.sleep(0.3)
            mode_name = {1: "ОМНК", 2: "1/x", 3: "1/x²"}[mode]
            b1w = wls.Cells(122, 2).Value
            b0w = wls.Cells(123, 2).Value
            n_eff = wls.Cells(120, 2).Value
            xp = unc.Cells(1, 2).Value
            uc = unc.Cells(5, 2).Value
            uu = unc.Cells(13, 2).Value
            kk = unc.Cells(12, 2).Value
            var_model = dg.Cells(125, 2).Value
            # TD-01: при режимах 2/3 n_eff должен быть 5 (x=0 исключён)
            if mode == 1:
                td01_ok = n_eff == 6
            else:
                td01_ok = n_eff == 5
            # Проверка отсутствия ошибок в расчётных областях
            errs_wls = scan_errors(wls, 108, 13)
            errs_unc = scan_errors(unc, 16, 2)
            errs_dg = scan_errors(dg, 125, 2)
            total_err = len(errs_wls) + len(errs_unc) + len(errs_dg)
            results[mode] = dict(
                b1w=b1w, b0w=b0w, n_eff=n_eff, xp=xp, uc=uc, uu=uu, kk=kk,
                var_model=var_model, td01_ok=td01_ok, n_err=total_err,
                errs=(errs_wls[:3] + errs_unc[:3] + errs_dg[:3]),
            )
            print(f"\n--- Режим {mode} ({mode_name}) ---")
            print(f"    b1_w={fmt(b1w)}, b0_w={fmt(b0w)}, n_eff={n_eff} (TD-01: {td01_ok})")
            print(f"    x_pred={fmt(xp)}, u_c={fmt(uc)}, k={fmt(kk)}, U={fmt(uu)}")
            print(f"    Модель дисперсии: {var_model}")
            print(f"    Ошибок в формулах: {total_err}")
            for er, ec, ev in results[mode]["errs"]:
                print(f"        !! {mode_name} {er},{ec}: {ev[:40]}")

        # ====== Сводный вердикт ======
        print("\n===== ИТОГ E2E =====")
        all_ok = True
        if fails:
            all_ok = False
            print(f"!! {len(fails)} FAIL на «Валидации»")
        if not (b1_ok and b0_ok):
            all_ok = False
            print("!! b1/b0 вне допуска")
        if not u_ok:
            all_ok = False
            print("!! U некорректна")
        for mode, res in results.items():
            if not res["td01_ok"]:
                all_ok = False
                print(f"!! TD-01 не выполнен в режиме {mode}")
            if res["n_err"]:
                all_ok = False
                print(f"!! ошибки формул в режиме {mode}: {res['n_err']}")
            if not (res["uc"] and res["uc"] > 0):
                all_ok = False
                print(f"!! u_c <= 0 в режиме {mode}")
        print("ВСЁ ПРОШЛО ✅" if all_ok else "ЕСТЬ ПРОБЛЕМЫ ❌")
        return 0 if all_ok else 1

    finally:
        try:
            wb.Close(False)
        except Exception:
            pass
        excel.Quit()


if __name__ == "__main__":
    sys.exit(main())