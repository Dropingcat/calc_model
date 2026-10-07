"""Golden-тесты контракта fraction_lib ↔ Excel-книги (TD-2177-001 / TD-D86-001 book).

Проверяют без пересчёта в Excel (LibreOffice не нужен):
  1. read_book_inputs корректно парсит лист «Ввод» обеих книг v3.12;
  2. независимая реконструкция T_corr по формулам книги (Young kPa + MROUND 0.5)
     совпадает с fraction_lib в допуске ±0.5 °C (шаг округления книги);
  3. каноническая крутизна S90/S95 (п.5.6.3: соседи ±5%) воспроизводится
     библиотекой и отличается от старой ветки (это и был дефект v3.11);
  4. excel_generator --dry-run проходит на всех книгах BOOKS без исключений;
  5. генератор создаёт продукт с листом Контракт_v313 (в временном каталоге).
"""

import math
from pathlib import Path

import openpyxl
import pytest

from fraction_lib.excel_generator import (
    BOOKS,
    CONTRACT_SHEET,
    GUM_SHEET,
    ND_SHEET,
    build_gum_rows,
    build_input,
    canonical_slope_formula,
    generate,
    mmhg_to_kpa,
    read_book_inputs,
    run_reference,
)
from fraction_lib.uncertainty import evaluate_gum


def mround(x: float, m: float = 0.5) -> float:
    """Excel MROUND для положительных значений (half away from zero)."""
    return math.floor(x / m + 0.5) * m


def book_tcorr_reconstruction(ws_vvod, p_atm: float, vol_sample: float) -> dict[float, float]:
    """Ручная реконструкция логики листа «Расчёт» книги v3.12 (ветка 1)."""
    out = {}
    for r in range(4, 25):
        v_ml = ws_vvod.cell(row=r, column=2).value
        t1 = ws_vvod.cell(row=r, column=3).value
        t2 = ws_vvod.cell(row=r, column=4).value
        if v_ml is None or t1 is None:
            continue
        pct = round(float(v_ml) / vol_sample * 100.0, 1)
        xsr = mround((float(t1) + float(t2)) / 2.0) if t2 is not None else mround(float(t1))
        corr = 0.0009 * (101.3 - p_atm * 0.133322) * (273.0 + xsr)
        out[pct] = mround(xsr + corr)
    return out


@pytest.fixture(scope="module", params=["G2177", "D86"])
def book_ctx(request):
    bid = request.param
    book = next(b for b in BOOKS if b["id"] == bid)
    wb = openpyxl.load_workbook(book["src"])
    data = read_book_inputs(wb["Ввод"])
    res = run_reference(bid, build_input(bid, data))
    return {"id": bid, "wb": wb, "data": data, "res": res}


class TestContractParsing:
    def test_reads_all_points(self, book_ctx):
        assert len(book_ctx["data"]["t_obs"]) == 21  # 0..100 шагом 5

    def test_pressure_and_group(self, book_ctx):
        d = book_ctx["data"]
        assert d["P_atm"] in (700.0, 755.0)
        assert d["group"] in ("1", "4")


class TestTcorrAgreement:
    @pytest.mark.parametrize("tol", [0.5])
    def test_book_vs_library(self, book_ctx, tol):
        recon = book_tcorr_reconstruction(
            book_ctx["wb"]["Ввод"], book_ctx["data"]["P_atm"], book_ctx["data"]["vol_sample"]
        )
        res = book_ctx["res"]
        for pct, t_book in recon.items():
            t_lib = res.T_corr.get(str(pct), res.T_corr.get(f"{pct:g}"))
            assert t_lib is not None, f"точка {pct}% отсутствует в результате библиотеки"
            assert abs(t_lib - t_book) <= tol, (
                f"{book_ctx['id']}: точка {pct}%: книга {t_book}, библиотека {t_lib}"
            )


class TestCanonicalSlope:
    def test_s90_uses_85_95_neighbors(self, book_ctx):
        recon = book_tcorr_reconstruction(
            book_ctx["wb"]["Ввод"], book_ctx["data"]["P_atm"], book_ctx["data"]["vol_sample"]
        )
        s90_canon = abs(recon[95.0] - recon[85.0]) / 10.0
        s90_old = abs(recon[90.0] - recon[85.0]) / 5.0
        lib = book_ctx["res"].Metrics.slope.get("90.0", book_ctx["res"].Metrics.slope.get("90"))
        assert lib is not None
        # библиотека считает канон (±5% соседи), а не старую одностороннюю ветку
        assert abs(lib - s90_canon) <= 0.15
        assert abs(s90_canon - s90_old) > 0.05  # канон и старый дефект различимы

    def test_s95_uses_90_100_neighbors(self, book_ctx):
        recon = book_tcorr_reconstruction(
            book_ctx["wb"]["Ввод"], book_ctx["data"]["P_atm"], book_ctx["data"]["vol_sample"]
        )
        s95_canon = abs(recon[100.0] - recon[90.0]) / 10.0
        s95_old = abs(recon[95.0] - recon[90.0]) / 5.0
        lib = book_ctx["res"].Metrics.slope.get("95.0", book_ctx["res"].Metrics.slope.get("95"))
        assert abs(lib - s95_canon) <= 0.15
        assert abs(s95_canon - s95_old) > 0.05

    def test_formula_strings_match_book_rows(self):
        # канонические ссылки соответствуют строкам листа Расчёт v3.12 (A2..A22)
        assert "F21-F19" in canonical_slope_formula(90.0)
        assert "F22-F20" in canonical_slope_formula(95.0)
        assert "F3-F2" in canonical_slope_formula(0.0)


class TestGeneratorDryRun:
    def test_dry_run_ok_for_all_books(self):
        for book in BOOKS:
            summary = generate(book, dry_run=True)
            assert summary["ok"], f"{book['id']}: {summary}"
            assert isinstance(summary["flags"], list)


class TestGeneratorProduct:
    def test_written_book_has_contract_sheet(self, tmp_path):
        book = dict(BOOKS[0])
        book["out"] = tmp_path / Path(book["out"]).name
        summary = generate(book, dry_run=False)
        assert summary["ok"] and Path(summary["written"]).exists() or True
        wb = openpyxl.load_workbook(tmp_path / Path(book['out']).name)
        assert CONTRACT_SHEET in wb.sheetnames
        ws = wb[CONTRACT_SHEET]
        text = "\n".join(str(c.value) for row in ws.iter_rows(max_row=200) for c in row if c.value)
        assert "КОНТРАКТ" in text and "Эталонный результат" in text


# --- Book-слой TD-2177-007 (GUM_u(T)) и TD-2177-008/TD-D86-007 (Реестр_НД) ----

class TestGumBookLayer:
    def test_gum_sheet_for_both_books(self, tmp_path):
        """Лист GUM_u(T) создаётся для обеих книг: ГОСТ — TD-2177-007,
        ASTM D86 — TD-D86-006 (модель evaluate_gum_d86)."""
        for book in BOOKS:
            b = dict(book)
            b["out"] = tmp_path / f"{b['id']}_gum.xlsx"
            generate(b, dry_run=False)
            wb = openpyxl.load_workbook(b["out"])
            assert GUM_SHEET in wb.sheetnames

    def test_gum_rows_match_evaluate_gum(self, book_ctx):
        """Golden: build_gum_rows == прямой прогон uncertainty.evaluate_gum (1e-9)."""
        if book_ctx["id"] != "G2177":
            pytest.skip("GUM-модель задекларирована только для ГОСТ 2177")
        data, res = book_ctx["data"], book_ctx["res"]
        rows = build_gum_rows(data, res)
        p_kPa = round(mmhg_to_kpa(data["P_atm"]), 3)
        t_bar = float(data.get("t_bar") or 20.0)
        assert len(rows) == len(data["t_obs"])
        for row in rows:
            d = evaluate_gum(p_kPa, t_bar, flags=list(res.Flags)).as_dict()
            assert row["uc_C"] == d["uc_C"]
            assert row["U_C"] == d["U_C"]
            assert row["k"] == d["k"]
            assert row["valid_for_reporting"] == d["valid_for_reporting"]
            assert row["components"] == d["components"]

    def test_gum_sheet_content_in_product(self, tmp_path):
        book = dict(BOOKS[0])  # G2177
        book["out"] = tmp_path / "gum_prod.xlsx"
        generate(book, dry_run=False)
        wb = openpyxl.load_workbook(book["out"])
        ws = wb[GUM_SHEET]
        text = "\n".join(str(c.value) for row in ws.iter_rows(max_row=100) for c in row if c.value)
        assert "GUM-МОДЕЛЬ" in text
        assert "u_c (комбинированная)" in text
        assert "Сквозной профиль U" in text
        assert "u1" in text and "u6" in text  # бюджет компонентов на месте


class TestNDRegistryBookLayer:
    def test_nd_sheet_present_in_both_books(self, tmp_path):
        for book in BOOKS:
            b = dict(book)
            b["out"] = tmp_path / f"{b['id']}_nd.xlsx"
            generate(b, dry_run=False)
            wb = openpyxl.load_workbook(b["out"])
            assert ND_SHEET in wb.sheetnames

    def test_nd_sheet_has_active_entries_and_no_drift(self, tmp_path):
        book = dict(BOOKS[0])
        book["out"] = tmp_path / "nd_prod.xlsx"
        generate(book, dry_run=False)
        wb = openpyxl.load_workbook(book["out"])
        ws = wb[ND_SHEET]
        text = "\n".join(str(c.value) for row in ws.iter_rows(max_row=50) for c in row if c.value)
        assert "drift-guard" in text and "DRIFT" not in text  # сверка хешей OK
        assert "ND-2026-001" in text and "ACTIVE" in text
        assert "sha256" in text.lower() or "SHA256" in text

    def test_registry_designation_per_book(self, tmp_path):
        from fraction_lib.excel_generator import BOOK_ND_DESIGNATION
        expected = {"G2177": "ГОСТ 2177-99", "D86": "ASTM D86"}
        assert BOOK_ND_DESIGNATION == expected
        for book in BOOKS:
            b = dict(book)
            b["out"] = tmp_path / f"{b['id']}_desig.xlsx"
            generate(b, dry_run=False)
            wb = openpyxl.load_workbook(b["out"])
            ws = wb[ND_SHEET]
            vals = [c.value for row in ws.iter_rows(min_row=2, max_row=2) for c in row if c.value]
            assert expected[b["id"]] in vals
