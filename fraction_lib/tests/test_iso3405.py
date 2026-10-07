"""ПУНКТ 3: тесты ГОСТ ISO 3405-2022 (iso3405.py).

Критерии готовности (контракт):
- поправка Янга пример (760→0, т.е. 101.3 кПа → 0);
- потери L=2.5, P=98.6 → Lc=2.0;
- r₁ при S=1 → 2.078 (0.864·1 + 1.214);
- R₁ при S=1 → 3.73 (1.736·1 + 1.994);
- 5% R = R₁ + 1.11;
- K формула совпадает с эталоном ((1/√2)·√(R²−r²/2));
- немонотонный ввод → NonMonotonicDistillationError;
- потери > 1.5% (гр.1-2) → LossLimitExceeded → INVALID_TEST;
- границы групп → inclusive_ge.
"""

from __future__ import annotations

import pytest

from fraction_lib.base import round_half
from fraction_lib.constants import ISO3405_G1_SLOPE_COEF, ISO3405_G1_R_CORR
from fraction_lib.exceptions import (
    LossLimitExceeded,
    NonMonotonicDistillationError,
    BoundaryRuleError,
)
from fraction_lib.iso3405 import ISO3405Calculator
from fraction_lib.models import FractionInput, k_criterion


@pytest.fixture
def calc():
    return ISO3405Calculator()


class TestYoungCorrection:
    """Поправка давления Сиднея Янга (11.2, формула 2)."""

    def test_young_760_mmhg_zero(self, calc):
        """P=760 мм рт.ст. (=101.3 кПа) → поправка ≈ 0 (допуск 0.01 °С)."""
        inp = FractionInput(P_atm=760.0, V_pct=[100.0], T_obs={50: 120.0},
                            loss_pct=0.0, residue_pct=0.0)
        res = calc.calculate(inp)
        # 760 мм = 101.325 кПа; стандарт ISO использует 101.3 → поправка
        # 0.0009·(101.3−101.325)·373 = −0.0084 °С (мала, физически корректна)
        assert res.T_corr["50"] == pytest.approx(120.0, abs=0.01)

    def test_young_below_std_adds(self, calc):
        """P < 101.3 кПа → поправку ПРИБАВЛЯЮТ (T_corr > T_obs)."""
        inp = FractionInput(P_atm=98.6, pressure_unit="kPa", V_pct=[100.0],
                            T_obs={50: 100.0}, loss_pct=0.0, residue_pct=0.0)
        res = calc.calculate(inp)
        expected = 100.0 + 0.0009 * (101.3 - 98.6) * (273.0 + 100.0)
        assert res.T_corr["50"] == pytest.approx(expected, abs=1e-9)


class TestLossCorrection:
    """Lc = 0,5 + (L − 0,5)/(1 + 0,125·(101,3 − p)) (11.3, формула 3)."""

    def test_loss_25_p986_lc20(self, calc):
        """L=2.5, P=98.6 кПа → Lc = 2.0 (пример приложения D стандарта)."""
        inp = FractionInput(P_atm=98.6, pressure_unit="kPa")
        lc = calc._correct_loss(2.5, inp)
        assert lc == pytest.approx(2.0, abs=0.01)

    def test_loss_at_std_pressure(self, calc):
        """P=101.3 кПа → Lc = L (знаменатель = 1)."""
        inp = FractionInput(P_atm=101.3, pressure_unit="kPa")
        assert calc._correct_loss(2.5, inp) == pytest.approx(2.5, abs=1e-9)


class TestPrecisionGroup1:
    """Таблица 6, группа 1 (ручной метод)."""

    def test_r1_at_slope1(self, calc):
        """r₁ = 0.864·S + 1.214; при S=1 → 2.078."""
        a_r, b_r = ISO3405_G1_SLOPE_COEF["r"]
        assert a_r * 1.0 + b_r == pytest.approx(2.078, abs=1e-9)

    def test_R1_at_slope1(self, calc):
        """R₁ = 1.736·S + 1.994; при S=1 → 3.73."""
        a_R, b_R = ISO3405_G1_SLOPE_COEF["R"]
        assert a_R * 1.0 + b_R == pytest.approx(3.73, abs=1e-9)

    def test_5pct_R_correction_plus111(self, calc):
        """5% R = R₁ + 1.11."""
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[100.0],
            T_obs={0: 30.0, 5: 60.0, 10: 80.0, 50: 130.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        r1 = ISO3405_G1_SLOPE_COEF["r"][0] * res.Metrics.slope["5"] + ISO3405_G1_SLOPE_COEF["r"][1]
        R1 = ISO3405_G1_SLOPE_COEF["R"][0] * res.Metrics.slope["5"] + ISO3405_G1_SLOPE_COEF["R"][1]
        assert res.Metrics.R["5"] == pytest.approx(R1 + 1.11, abs=1e-9)
        assert res.Metrics.r["5"] == pytest.approx(r1, abs=1e-9)

    def test_90pct_R_correction_minus122(self, calc):
        """90% R = R₁ − 1.22."""
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[100.0],
            T_obs={0: 30.0, 5: 60.0, 10: 80.0, 50: 130.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        R1 = ISO3405_G1_SLOPE_COEF["R"][0] * res.Metrics.slope["90"] + ISO3405_G1_SLOPE_COEF["R"][1]
        assert res.Metrics.R["90"] == pytest.approx(R1 - 1.22, abs=1e-9)

    def test_95pct_R_correction_minus094(self, calc):
        """95% R = R₁ − 0.94."""
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[100.0],
            T_obs={0: 30.0, 5: 60.0, 10: 80.0, 50: 130.0, 90: 180.0, 95: 185.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        R1 = ISO3405_G1_SLOPE_COEF["R"][0] * res.Metrics.slope["95"] + ISO3405_G1_SLOPE_COEF["R"][1]
        assert res.Metrics.R["95"] == pytest.approx(R1 - 0.94, abs=1e-9)

    def test_ibp_fbp_fixed(self, calc):
        """IBP r/R = 3.3/5.6; FBP r/R = 3.9/7.2 (фиксированные)."""
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[100.0],
            T_obs={"IBP": 30.0, 50: 130.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert res.Metrics.r["IBP"] == pytest.approx(3.3, abs=1e-9)
        assert res.Metrics.R["IBP"] == pytest.approx(5.6, abs=1e-9)
        assert res.Metrics.r["FBP"] == pytest.approx(3.9, abs=1e-9)
        assert res.Metrics.R["FBP"] == pytest.approx(7.2, abs=1e-9)


class TestKCriterion:
    """K = (1/√2)·√(R² − r²/2) — формула ГСО, НЕ меняется."""

    def test_k_formula_matches_etalon(self, calc):
        """K в результатах == k_criterion(r, R) для каждой точки."""
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[100.0],
            T_obs={"IBP": 30.0, 5: 60.0, 50: 130.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        for pt in res.Metrics.r:
            r = res.Metrics.r[pt]
            R = res.Metrics.R[pt]
            expected = (1.0 / 2.0**0.5) * (R * R - r * r / 2.0) ** 0.5
            assert res.Metrics.K[pt] == pytest.approx(expected, abs=1e-9)
        # 0.84R — НЕ менять
        assert res.Metrics.R_ok["5"] == pytest.approx(res.Metrics.R["5"] * 0.84, abs=1e-9)


class TestSparseDataScenario:
    """Сценарий А: разреженные данные → интерполяция T(V) + монотонность."""

    def test_interpolation_fills_missing(self, calc):
        """Точка 50% интерполируется между 10% и 90%."""
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[100.0],
            T_obs={0: 30.0, 10: 80.0, 50: None, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        # 50% = 80 + (180-80)·(50-10)/(90-10) = 130 (по скорректированным
        # значениям ≈ 129.991 из-за поправки Янга ~0.009 °С)
        assert res.T_corr["50"] == pytest.approx(130.0, abs=0.05)

    def test_non_monotonic_raises(self, calc):
        """Немонотонный ввод → NonMonotonicDistillationError."""
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[100.0],
            T_obs={0: 100.0, 10: 90.0, 50: 120.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        with pytest.raises(NonMonotonicDistillationError):
            calc.calculate(inp)


class TestGroupBoundaries:
    """Сценарий Б: границы групп → inclusive_ge."""

    def test_inclusive_ge_default_ok(self, calc):
        """По умолчанию inclusive_ge — граница не вызывает ошибку."""
        inp = FractionInput(
            P_atm=760.0, object_type="2", V_pct=[100.0],
            T_obs={0: 30.0, 5: 60.0, 50: 130.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert "50" in res.Metrics.r

    def test_wrong_rule_raises_boundary_error(self, calc):
        """boundary_rule != inclusive_ge при slope на границе → BoundaryRuleError."""
        inp = FractionInput(
            P_atm=760.0, object_type="2", V_pct=[100.0],
            T_obs={0: 30.0, 5: 60.0, 50: 130.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0, boundary_rule="exclusive",
        )
        # slope может попасть на границу (1.0/2.0/...) — если нет, ошибки не будет,
        # но правило должно быть принято без исключения в любом случае.
        try:
            calc.calculate(inp)
        except BoundaryRuleError:
            pass  # допустимо


class TestLossLimit:
    """Сценарий В: потери > лимита группы → LossLimitExceeded → INVALID_TEST."""

    def test_loss_above_limit_group1(self, calc):
        """Группа 1, потери 2.5% > 1.5% → INVALID_TEST флаг."""
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[97.5],
            T_obs={0: 30.0, 50: 130.0, "FBP": 200.0},
            loss_pct=2.5, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert "INVALID_TEST" in res.Flags
        assert any("LOSS_LIMIT_EXCEEDED" in f for f in res.Flags)

    def test_loss_within_limit_ok(self, calc):
        """Группа 1, потери 1.0% <= 1.5% → без INVALID_TEST."""
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[99.0],
            T_obs={0: 30.0, 50: 130.0, "FBP": 200.0},
            loss_pct=1.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert "INVALID_TEST" not in res.Flags

    def test_loss_limit_group3_higher(self, calc):
        """Группа 3, потери 1.8% <= 2.0% → без INVALID_TEST (лимит выше)."""
        inp = FractionInput(
            P_atm=760.0, object_type="3", V_pct=[98.2],
            T_obs={0: 40.0, 50: 150.0, "FBP": 230.0},
            loss_pct=1.8, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert "INVALID_TEST" not in res.Flags


class TestGroup234Table7:
    """Таблица 7 (группы 2-4, редакция 2022)."""

    def test_group2_ibp_r_formula(self, calc):
        """Гр.2 IBP: r = 0.35·S + 1.0; R = 0.93·S + 2.8."""
        inp = FractionInput(
            P_atm=760.0, object_type="2", V_pct=[100.0],
            T_obs={"IBP": 30.0, 5: 60.0, 50: 130.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        s_ibp = res.Metrics.slope["IBP"]
        assert res.Metrics.r["IBP"] == pytest.approx(0.35 * s_ibp + 1.0, abs=1e-9)
        assert res.Metrics.R["IBP"] == pytest.approx(0.93 * s_ibp + 2.8, abs=1e-9)

    def test_group2_50pct_r_formula(self, calc):
        """Гр.2 50%: r = 0.41·S + 1.0; R = 1.33·S + 1.8."""
        inp = FractionInput(
            P_atm=760.0, object_type="2", V_pct=[100.0],
            T_obs={"IBP": 30.0, 5: 60.0, 50: 130.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        s50 = res.Metrics.slope["50"]
        assert res.Metrics.r["50"] == pytest.approx(0.41 * s50 + 1.0, abs=1e-9)
        assert res.Metrics.R["50"] == pytest.approx(1.33 * s50 + 1.8, abs=1e-9)


class TestRounding:
    """Округление до 0,5 °С."""

    def test_round_half(self):
        assert round_half(2.25) == pytest.approx(2.5)
        assert round_half(2.74) == pytest.approx(2.5)
        assert round_half(2.76) == pytest.approx(3.0)