"""Генератор Excel-контракта из fraction_lib (TD-2177-001 / TD-D86-001 book).

Единый источник истины формул расчётного слоя. Для каждой книги v3.12
создаёт производную книгу v3.13 с листом «Контракт_v313», куда записывает:

  1. Нормативные константы, вшитые в библиотеку (диапазон Табл.4, лимит
     экстраполяции, политики) — чтобы ревизор видел, что книга и код сверяются
     по одним и тем же числам;
  2. Канонические формулы крутизны S(p) по п.5.6.3 ГОСТ 2177 / A4.3 ASTM D86
     (центральная разность ±5%, для 90% — 85..95, для 95% — 90..100);
  3. Эталонный расчёт fraction_lib на данных листа «Ввод» книги (T_corr, r, R,
     slope по точкам) — независимые числа, которые можно сравнить с пересчитанной
     книгой (Excel F9) без LibreOffice;
  4. Флаги политики (PRESSURE_CLAMPED / EXTRAPOLATED_96_98 / NOMOGRAM_CLAMPED_HIGH),
     если они сработали на эталонных данных.

Принципы:
  * Исходные книги НЕ модифицируются — читаются только лист «Ввод».
  * Новый лист не конфликтует с существующими именами листов.
  * Все значения сериализуются как числа/строки (без формул) — лист пригоден
    для машинного diff (golden-тесты test_excel_contract.py).

Использование:
    python -m fraction_lib.excel_generator            # все известные книги
    python -m fraction_lib.excel_generator --dry-run  # только расчёт, без записи
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
from pathlib import Path

import openpyxl

from .constants import (
    GOST2177_PRESSURE_RANGE,
    GOST2177_EXTRAPOLATION_MAX_PCT,
    GOST2177_TABLE4_POLICY_DEFAULT,
    GOST2177_TABLE4,
)
from .gost2177 import GOST2177Calculator
from .d86 import D86Calculator
from .models import FractionInput

REPO_ROOT = Path(__file__).resolve().parent.parent

# Книги-источники (v3.12) и имена продуктовых книг (v3.13)
BOOKS = [
    {
        "id": "G2177",
        "method": "gost2177",
        "src": REPO_ROOT / "new_versions_v4.1" / "ГОСТ_2177_метод_А_v3.12_нормативно_аудированный.xlsx",
        "out": REPO_ROOT / "new_versions_v4.2" / "ГОСТ_2177_метод_А_v3.13_контракт_с_генератором.xlsx",
    },
    {
        "id": "D86",
        "method": "d86",
        "src": REPO_ROOT / "new_versions_v4.1" / "ASTM_D86_ручной_метод_v3.12_нормативно_аудированный.xlsx",
        "out": REPO_ROOT / "new_versions_v4.2" / "ASTM_D86_ручной_метод_v3.13_контракт_с_генератором.xlsx",
    },
]

CONTRACT_SHEET = "Контракт_v313"


def read_book_inputs(ws) -> dict:
    """Разбирает лист «Ввод» книги формата v3.11/v3.12 (фиксированная сетка).

    Ячейки: B1 — давление мм рт.ст.; группа/тип — в ячейке справа от подписи
    «Группа» (ГОСТ v3.12: подпись D1 / значение E1; ASTM v3.12: C1 / D1);
    G1 — объём образца мл; строки 4..24: B — V мл, C — T1, D — T2;
    E25/F25 — потери/остаток %.
    """
    p_atm = ws["B1"].value
    group = None
    for col in range(1, min(ws.max_column, 8) + 1):
        cell = ws.cell(row=1, column=col)
        if isinstance(cell.value, str) and "рупп" in cell.value:  # «Группа»/«группы»
            right = ws.cell(row=1, column=col + 1).value
            group = right if right is not None else cell.value
            break
    if group is None:
        group = ws["E1"].value
    vol_sample = ws["G1"].value or 100.0
    loss_pct = ws["E25"].value
    residue_pct = ws["F25"].value

    t_obs: dict[float, float] = {}
    for row in range(4, 25):
        v_ml = ws.cell(row=row, column=2).value
        t1 = ws.cell(row=row, column=3).value
        if v_ml is None or t1 is None:
            continue
        pct = round(float(v_ml) / float(vol_sample) * 100.0, 1)
        t_obs[pct] = float(t1)

    return {
        "P_atm": float(p_atm),
        "group": str(int(group)) if isinstance(group, (int, float)) else str(group),
        "vol_sample": float(vol_sample),
        "loss_pct": float(loss_pct) if loss_pct is not None else None,
        "residue_pct": float(residue_pct) if residue_pct is not None else None,
        "t_obs": t_obs,
    }


def build_input(book_id: str, data: dict) -> FractionInput:
    inp = FractionInput(
        P_atm=data["P_atm"],
        object_type=data["group"],
        method="gost2177" if book_id == "G2177" else "d86",
        T_obs={k: v for k, v in sorted(data["t_obs"].items())},
        loss_pct=data["loss_pct"],
        residue_pct=data["residue_pct"],
    )
    return inp


def canonical_slope_formula(pct: float) -> str:
    """Каноническая формула крутизны п.5.6.3 (равноотстоящие соседи ±5%).

    В терминах строк листа «Расчёт» v3.12 (A=2..22 = 0..100% шагом 5):
      S(p) = ABS(F{r+1} - F{r-1}) / (A{r+1} - A{r-1}),  где r = p/5 + 2.
    Особые точки: IBP — первый интервал 0→5; КК/FBP — последний интервал.
    """
    if pct <= 0:
        return "=ABS(F3-F2)/(A3-A2)   # IBP: интервал 0->5"
    if pct >= 100:
        return "=ABS(F22-F21)/(A22-A21)   # КК: последний интервал 95->100"
    r = int(round(pct / 5.0)) + 2
    lo, hi = pct - 5.0, pct + 5.0
    note = ""
    if pct == 90.0:
        note = "   # исправлено v3.12: было 85-90, по 5.6.3 нужно 85-95"
    if pct == 95.0:
        note = "   # исправлено v3.12: было 90-95, по 5.6.3 нужно 90-100"
    return f"=ABS(F{r + 1}-F{r - 1})/(A{r + 1}-A{r - 1})   # соседи {lo:g}..{hi:g}%{note}"


def run_reference(book_id: str, inp: FractionInput):
    calc = GOST2177Calculator() if book_id == "G2177" else D86Calculator()
    return calc.calculate(inp)


def write_contract_sheet(wb_out, book: dict, data: dict, res) -> None:
    ws = wb_out.create_sheet(CONTRACT_SHEET)
    r = 1

    def put(*vals):
        nonlocal r
        for c, v in enumerate(vals, start=1):
            ws.cell(row=r, column=c, value=v)
        r += 1

    put("КОНТРАКТ fraction_lib ↔ книга (сгенерировано excel_generator)")
    put("Дата:", _dt.date.today().isoformat())
    put("Метод:", book["method"], "Книга-источник:", book["src"].name)
    put("")

    put("== 1. Нормативные константы библиотеки ==")
    put("GOST2177_PRESSURE_RANGE (мм рт.ст.):", GOST2177_PRESSURE_RANGE[0], GOST2177_PRESSURE_RANGE[1])
    put("GOST2177_EXTRAPOLATION_MAX_PCT:", GOST2177_EXTRAPOLATION_MAX_PCT)
    put("GOST2177_TABLE4_POLICY_DEFAULT:", GOST2177_TABLE4_POLICY_DEFAULT)
    put("Строк Таблицы 4:", len(GOST2177_TABLE4))
    put("")

    put("== 2. Входные данные листа «Ввод» (эталон прогона) ==")
    put("P, мм рт.ст.:", data["P_atm"], "| Группа/тип:", data["group"],
        "| Объём пробы, мл:", data["vol_sample"])
    put("Потери, %:", data["loss_pct"], "| Остаток, %:", data["residue_pct"])
    put("% отгона", "T_набл, °C")
    for pct, t in sorted(data["t_obs"].items()):
        put(pct, t)
    put("")

    put("== 3. Канонические формулы крутизны S(p) (п.5.6.3 / A4.3) ==")
    put("% отгона", "Формула (строчки листа Расчёт v3.12)", "S_ref (fraction_lib)")
    for pct in [0.0, 5.0, 10.0, 15.0, 20.0, 25.0, 30.0, 35.0, 40.0, 45.0, 50.0,
                55.0, 60.0, 65.0, 70.0, 75.0, 80.0, 85.0, 90.0, 95.0, 100.0]:
        s_val = res.Metrics.slope.get(str(pct)) or res.Metrics.slope.get(f"{pct:g}")
        put(pct, canonical_slope_formula(pct), s_val)
    put("")

    put("== 4. Эталонный результат fraction_lib ==")
    put("Loss_скорр, %:", round(res.Loss, 6), "| Residue, %:", res.Residue)
    put("% отгона", "T_corr_ref", "slope_ref", "r_ref", "R_ref")
    keys = sorted((set(res.T_corr) | set(res.Metrics.r)),
                  key=lambda k: (0, float(k)) if str(k).replace(".", "").isdigit() else (1, str(k)))
    for k in keys:
        put(k, res.T_corr.get(k), res.Metrics.slope.get(k),
            res.Metrics.r.get(k), res.Metrics.R.get(k))
    put("")

    put("== 5. Флаги политики на эталонных данных ==")
    put("Flags:", ", ".join(res.Flags) if res.Flags else "(нет — данные в каноне)")
    put("")
    put("Как сверять: открыть книгу в Excel, F9 (полный пересчёт), сравнить")
    put("столбец T_corr листа «Расчёт» (F2:F22) со столбцом T_corr_ref здесь;")
    put("расхождение > 0.5 °C в любой точке = контракт нарушен (регистрация в Техдолг).")


def generate(book: dict, dry_run: bool = False) -> dict:
    src = Path(book["src"])
    wb_in = openpyxl.load_workbook(src, data_only=False)
    data = read_book_inputs(wb_in["Ввод"])
    inp = build_input(book["id"], data)
    res = run_reference(book["id"], inp)

    summary = {
        "book": book["id"],
        "method": res.method,
        "n_points": len(data["t_obs"]),
        "flags": list(res.Flags),
        "loss_corr": res.Loss,
        "ok": True,
    }

    if not dry_run:
        out = Path(book["out"])
        out.parent.mkdir(parents=True, exist_ok=True)
        wb_out = openpyxl.load_workbook(src)
        if CONTRACT_SHEET in wb_out.sheetnames:
            del wb_out[CONTRACT_SHEET]
        write_contract_sheet(wb_out, book, data, res)
        wb_out.save(out)
        try:
            summary["written"] = str(out.relative_to(REPO_ROOT))
        except ValueError:  # продукт вне репозитория (tmp_path в тестах)
            summary["written"] = str(out)
    return summary


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--dry-run", action="store_true", help="не записывать книги")
    args = ap.parse_args(argv)

    results = []
    for book in BOOKS:
        try:
            results.append(generate(book, dry_run=args.dry_run))
        except Exception as e:  # noqa: BLE001
            results.append({"book": book["id"], "ok": False, "error": f"{type(e).__name__}: {e}"})
    print(json.dumps(results, ensure_ascii=False, indent=2))
    return 0 if all(x.get("ok") for x in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
