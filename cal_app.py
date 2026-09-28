#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cal_app.py — единый портативный конвейер калибровки (Metrology Core EURACHEM).

ДВЕ ВЕТКИ:
  1) EXCEL-ветка (нужен установленный Microsoft Excel): генерация xlsx с формулами.
  2) PYTHON/MC-ветка (не нужен Excel): расчёт в чистом Python (calc_engine),
     Monte Carlo, отчёт из результата. РАБОТАЕТ ВЕЗДЕ (в т.ч. с флешки).

Все вызовы — внутрипроцессные (важно для PyInstaller onefile).

Использование:
  cal_app run                      — автономный полный цикл: calc(py) → validate → MC → отчёт
  cal_app run --excel              — то же, но с генерацией через Excel (если установлен)
  cal_app calc [--no-excel]        — только расчёт/генерация xlsx
  cal_app validate                 — кросс-валидация (numpy, без Excel)
  cal_app mc [N]                   — Monte Carlo (N итераций)
  cal_app report [--from-py]       — отчёт Word
  cal_app run --xlsx FILE          — костыль: использовать внешний xlsx (его данные)
"""
import os
import sys

# Принудительный UTF-8 для консоли (Windows cp1251 ломает кириллицу/стрелки)
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
os.environ.setdefault("PYTHONIOENCODING", "utf-8")

# Явные импорты, чтобы PyInstaller включил модули в один exe
import calc_engine  # noqa: F401
import validate_auto  # noqa: F401
import mc_v4  # noqa: F401
import report_v4  # noqa: F401
import cal_v4  # noqa: F401

# Папка исполняемого файла (exe) или скрипта — туда пишем артефакты.
# ВАЖНО для PyInstaller onefile: __file__ указывает на временную распаковку,
# поэтому берём sys.executable (сам exe), а не __file__.
if getattr(sys, "frozen", False):
    HERE = os.path.dirname(sys.executable)
else:
    HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "Metrology_Core_EURACHEM_v4.xlsx")


def has_excel():
    try:
        import win32com.client  # noqa
        return True
    except Exception:
        return False


def _parse(args, _xlsx=None):
    """Извлекает --xlsx FILE и --excel флаг."""
    xlsx = _xlsx or XLSX
    use_excel = False
    rest = []
    i = 0
    while i < len(args):
        a = args[i]
        if a == "--xlsx" and i + 1 < len(args):
            xlsx = os.path.abspath(args[i + 1]); i += 2; continue
        if a == "--excel":
            use_excel = True
        elif a == "--no-excel":
            use_excel = False
        else:
            rest.append(a)
        i += 1
    return xlsx, use_excel, rest


def cmd_calc(args):
    xlsx, use_excel, _ = _parse(args)
    if use_excel and has_excel():
        # Excel-ветка: cal_v4 генерирует xlsx с формулами
        return cal_v4.create_metrology_excel_v4(xlsx)
    # Python-ветка: автономный расчёт
    return calc_engine.generate_without_excel(xlsx)


def cmd_validate(args):
    xlsx, _, _ = _parse(args)
    return validate_auto.main(xlsx)


def cmd_mc(args):
    xlsx, _, rest = _parse(args)
    n = int(rest[0]) if rest else 100000
    return mc_v4.run(xlsx, n)


def cmd_report(args):
    xlsx, _, rest = _parse(args)
    from_py = "--from-py" in rest
    if from_py or not has_excel():
        # автономная ветка: отчёт из Python-расчёта
        d = report_v4.read_from_py(xlsx)
        out = os.path.join(HERE, "Отчёт_по_калибровке.docx")
        print(f"файл: {xlsx}, вердикт: {d['verdict']}, x_pred={d['x_pred']}")
        report_v4.build_report(d, out)
        return 0
    return report_v4.main()


def cmd_run(args):
    xlsx, use_excel, _ = _parse(args)
    if use_excel and has_excel():
        if cmd_calc(["--excel", "--xlsx", xlsx]) != 0:
            return 1
    else:
        if cmd_calc(["--no-excel", "--xlsx", xlsx]) != 0:
            return 1
    if cmd_validate(["--xlsx", xlsx]) != 0:
        return 1
    if cmd_mc(["50000", "--xlsx", xlsx]) != 0:
        return 1
    return cmd_report(["--from-py", "--xlsx", xlsx])


def main():
    if len(sys.argv) < 2:
        # Двойной клик по exe (без аргументов): сразу полный цикл, без help.
        print("Запуск по двойному клику — выполняю полный цикл (run).")
        print("Подсказка: `cal_app.exe help` — список команд.\n")
        rc = cmd_run([])
        _pause()
        return rc
    cmd = sys.argv[1]
    args = sys.argv[2:]
    if cmd in ("help", "-h", "--help"):
        print(__doc__)
        _pause()
        return 0
    dispatch = {
        "run": cmd_run,
        "calc": cmd_calc,
        "validate": cmd_validate,
        "mc": cmd_mc,
        "report": cmd_report,
    }
    fn = dispatch.get(cmd)
    if fn is None:
        print(__doc__)
        return 2
    return fn(args)


def _pause():
    try:
        input("\nНажмите Enter для закрытия...")
    except Exception:
        pass


if __name__ == "__main__":
    sys.exit(main())