"""Тесты TD-D86-002 (Table 5) и TD-D86-006 (GUM D86): код + book-слой."""

from __future__ import annotations

import math

import pytest
from pathlib import Path

from fraction_lib.d86_procedure import (
    D86TestConditions, check_conditions, table5_reference_rows,
)
from fraction_lib.uncertainty import evaluate_gum_d86, UncertaintyComponent
from fraction_lib.constants import (
    D86_TABLE5_GROUPS, D86_YOUNG_K_MMHG, LOSS_CORR_MMHG_STD,
)


def ok_cond(group=1, **kw):
    base = dict(group=group, sample_volume_ml=100.0, condense_rate_ml_min=4.5,
                ibp_to_5_min=8.0, time_to_ibp_min=7.0, bath_preheat_ok=True,
                apparatus_group_ok=True)
    base.update(kw)
    return D86TestConditions(**base)


class TestTable5CheckTD_D86_002:
    def test_pass_all_groups(self):
        for g in sorted(D86_TABLE5_GROUPS):
            kw = {}
            if g in (3, 4):   # IBP→5% нормативно не лимитирован для гр.3–4
                kw["ibp_to_5_min"] = None
            c = ok_cond(group=g, **kw)
            v = check_conditions(c)
            assert v["verdict"] == "PASS", (g, v)

    def test_missing_fields_fail(self):
        c = D86TestConditions(group=1)   # всё незаполнено
        v = check_conditions(c)
        assert v["verdict"] == "FAIL"
        assert any(x["id"] == "CONDENSE_RATE" and x["severity"] == "MISSING"
                   for x in v["violations"])

    def test_wrong_group_invalid(self):
        v = check_conditions(ok_cond(group=9))
        assert v["verdict"] == "FAIL"
        assert "TABLE5_INVALID:GROUP" in v["flags"]

    def test_condensation_out_of_range_invalid(self):
        c = ok_cond(condense_rate_ml_min=6.0)
        v = check_conditions(c)
        assert v["verdict"] == "FAIL"
        assert "TABLE5_INVALID:CONDENSE_RATE" in v["flags"]

    def test_sample_volume_out_of_range_invalid(self):
        c = ok_cond(sample_volume_ml=95.0)
        v = check_conditions(c)
        assert "TABLE5_INVALID:SAMPLE_VOL" in v["flags"]

    def test_group4_slow_condensation_allowed(self):
        # нижняя граница для гр.4 нормативно не задана — 2 мл/мин допустимо
        c = ok_cond(group=4, condense_rate_ml_min=2.0, ibp_to_5_min=None)
        assert check_conditions(c)["verdict"] == "PASS"

    def test_group4_too_fast_invalid(self):
        c = ok_cond(group=4, condense_rate_ml_min=6.0, ibp_to_5_min=None)
        assert "TABLE5_INVALID:CONDENSE_RATE" in check_conditions(c)["flags"]

    def test_time_to_ibp_is_warning_only(self):
        c = ok_cond(time_to_ibp_min=20.0)
        v = check_conditions(c)
        assert v["verdict"] == "CONDITIONAL"
        assert "WARNING:TABLE5:TIME_TO_IBP" in v["flags"]

    def test_reference_rows_shape(self):
        rows = table5_reference_rows()
        assert len(rows) == 4
        assert all(len(row) == 6 for row in rows)


class TestGumD86TD_D86_006:
    def test_budget_components(self):
        gum = evaluate_gum_d86(760.0, 20.0, 1.0)
        names = [c.name for c in gum.components]
        assert names == ["u1", "u2", "u3", "u4", "u5", "u6", "u7"]
        assert gum.uc > 0
        assert abs(gum.U - gum.k * gum.uc) < 1e-12

    def test_sensitivities_match_young(self):
        p, t = 750.0, 25.0
        gum = evaluate_gum_d86(p, t, 2.0)
        c_p = -D86_YOUNG_K_MMHG * (273.0 + t)
        c_t = D86_YOUNG_K_MMHG * (LOSS_CORR_MMHG_STD - p)
        d = {c.name: c for c in gum.components}
        assert d["u2"].sensitivity == pytest.approx(c_p)
        assert d["u3"].sensitivity == pytest.approx(c_t)

    def test_volume_channel_uses_slope(self):
        g_lo = evaluate_gum_d86(760.0, 20.0, 0.1)
        g_hi = evaluate_gum_d86(760.0, 20.0, 5.0)
        assert g_hi.uc > g_lo.uc   # крутая кривая ⇒ больший вклад объёма

    def test_type_a_reduces_u1(self):
        gum = evaluate_gum_d86(760.0, 20.0, 1.0, s_rep=0.2, n_repeats=4)
        u1 = gum.components[0]
        assert u1.kind == "A"
        assert u1.halfwidth == pytest.approx(0.2 / 2.0)
        assert math.isfinite(gum.dof_eff) and gum.dof_eff > 0

    def test_policy_flags_block_reporting(self):
        for flag in ("LOSS_CORR_UNSTABLE:P=820mm", "TABLE5_INVALID:CONDENSE_RATE",
                     "INVALID_TEST"):
            gum = evaluate_gum_d86(760.0, 20.0, 1.0, flags=[flag])
            assert gum.valid_for_reporting is False, flag
            assert gum.notes

    def test_clean_flags_ok(self):
        gum = evaluate_gum_d86(760.0, 20.0, 1.0, flags=["WARNING:TABLE5:TIME_TO_IBP"])
        assert gum.valid_for_reporting is True

    def test_uc_independent_sum(self):
        gum = evaluate_gum_d86(760.0, 20.0, 1.0)
        manual = math.sqrt(sum(c.u ** 2 for c in gum.components))
        assert gum.uc == pytest.approx(manual, rel=1e-12)


class TestBookLayerTD_D86_002_006:
    """Книжный слой: генератор пишет GUM_u(T) и Условия_Table5 в ASTM-книгу."""

    def test_generated_astm_book_has_sheets(self, tmp_path):
        from fraction_lib import excel_generator as EG

        book = dict(EG.BOOKS[1])          # D86
        out = tmp_path / Path(book["out"]).name
        book["out"] = out
        summary = EG.generate(book)
        assert summary["ok"], summary
        import openpyxl
        wb = openpyxl.load_workbook(out)
        assert EG.GUM_SHEET in wb.sheetnames
        assert EG.TABLE5_SHEET in wb.sheetnames
        ws = wb[EG.GUM_SHEET]
        text = "\n".join(str(c.value) for row in ws.iter_rows(max_row=6)
                         for c in row if c.value)
        assert "TD-D86-006" in text
        ws5 = wb[EG.TABLE5_SHEET]
        text5 = "\n".join(str(c.value) for row in ws5.iter_rows(max_row=40)
                          for c in row if c.value)
        assert "TD-D86-002" in text5
        assert "Вердикт:" in text5

    def test_gum_d86_rows_match_model(self):
        """Golden: строки генератора == прямой прогон evaluate_gum_d86.as_dict()."""
        from fraction_lib import excel_generator as EG

        book = EG.BOOKS[1]
        import openpyxl
        wb = openpyxl.load_workbook(EG.Path(book["src"]), data_only=False)
        data = EG.read_book_inputs(wb["Ввод"])
        inp = EG.build_input(book["id"], data)
        res = EG.run_reference(book["id"], inp)
        rows = EG.build_gum_d86_rows(data, res)
        p_mm = float(data["P_atm"])
        t_bar = float(data.get("t_bar") or 20.0)
        flask = float(data.get("vol_sample") or 100.0)
        for row in rows:
            slope = res.Metrics.slope.get(str(row["V_pct"])) \
                or res.Metrics.slope.get(f"{row['V_pct']:g}") or 0.0
            direct = evaluate_gum_d86(p_mm, t_bar, slope,
                                               flask_volume_ml=flask,
                                               flags=list(res.Flags)).as_dict()
            assert row["uc_C"] == direct["uc_C"]
            assert row["U_C"] == direct["U_C"]
            assert row["components"] == direct["components"]
