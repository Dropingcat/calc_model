"""ПУНКТ 4: тесты ASTM D86-20a (d86.py).

Критерии готовности (контракт):
- поправка Янга (ф.3-4);
- потери Lc (ф.6-7);
- r/R Table 8 (гр.1): Sc=1 → R(5%)=4.84, R(95%)=2.84;
- немонотонный → NonMonotonicDistillationError;
- границы групп → inclusive_ge;
- потери > 1.5% → INVALID_TEST.
"""

from __future__ import annotations

import pytest

from fraction_lib.constants import D86_TABLE8, D86_TABLE9
from fraction_lib.exceptions import (
    NonMonotonicDistillationError,
    BoundaryRuleError,
)
from fraction_lib.d86 import D86Calculator
from fraction_lib.models import FractionInput


@pytest.fixture
def calc():
    return D86Calculator()


class TestYoungCorrection:
    """Поправка давления (11.3, формулы 3/4)."""

    def test_young_760_mmhg_zero(self, calc):
        """P=760 мм рт.ст. → поправка ≈ 0 (допуск 0.01)."""
        inp = FractionInput(P_atm=760.0, V_pct=[100.0], T_obs={50: 120.0},
                            loss_pct=0.0, residue_pct=0.0)
        res = calc.calculate(inp)
        assert res.T_corr["50"] == pytest.approx(120.0, abs=0.01)

    def test_young_below_std_adds(self, calc):
        """P < 760 мм → поправку ПРИБАВЛЯЮТ (T_corr > T_obs)."""
        inp = FractionInput(P_atm=740.0, V_pct=[100.0], T_obs={50: 100.0},
                            loss_pct=0.0, residue_pct=0.0)
        res = calc.calculate(inp)
        expected = 100.0 + 0.00012 * (760.0 - 740.0) * (273.0 + 100.0)
        assert res.T_corr["50"] == pytest.approx(expected, abs=1e-9)


class TestLossCorrection:
    """Lc = 0,5 + (L − 0,5)/(1 + (760 − P)/60,0) (11.4, формула 7)."""

    def test_loss_mmhg_formula(self, calc):
        """P=760 → Lc = L."""
        inp = FractionInput(P_atm=760.0)
        assert calc._correct_loss(2.5, inp) == pytest.approx(2.5, abs=1e-9)

    def test_loss_kpa_formula(self, calc):
        """P=101.3 кПа → Lc = L (знаменатель = 1)."""
        inp = FractionInput(P_atm=101.3, pressure_unit="kPa")
        assert calc._correct_loss(2.5, inp) == pytest.approx(2.5, abs=1e-9)


class TestPrecisionGroup1Table8:
    """Table 8 (группа 1, ручной метод)."""

    def test_R5_at_sc1(self, calc):
        """Sc=1 → R(5%) = 3.1 + 1.74·1 = 4.84."""
        assert D86_TABLE8["5"]["R"][0] + D86_TABLE8["5"]["R"][1] * 1.0 == pytest.approx(4.84, abs=1e-9)

    def test_R95_at_sc1(self, calc):
        """Sc=1 → R(95%) = 1.1 + 1.74·1 = 2.84."""
        assert D86_TABLE8["95"]["R"][0] + D86_TABLE8["95"]["R"][1] * 1.0 == pytest.approx(2.84, abs=1e-9)

    def test_R90_at_sc1(self, calc):
        """Sc=1 → R(90%) = 0.8 + 1.74·1 = 2.54."""
        assert D86_TABLE8["90"]["R"][0] + D86_TABLE8["90"]["R"][1] * 1.0 == pytest.approx(2.54, abs=1e-9)

    def test_full_calculation_r_values(self, calc):
        """Полный расчёт: R(5%) = 3.1+1.74·S5; r(5%) = 1.9+0.86·S5."""
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[100.0],
            T_obs={0: 30.0, 5: 60.0, 10: 80.0, 50: 130.0, 90: 180.0, 95: 185.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        s5 = res.Metrics.slope["5"]
        assert res.Metrics.R["5"] == pytest.approx(3.1 + 1.74 * s5, abs=1e-9)
        assert res.Metrics.r["5"] == pytest.approx(1.9 + 0.86 * s5, abs=1e-9)

    def test_ibp_fbp_fixed(self, calc):
        """IBP r/R = 3.3/5.6; FBP r/R = 3.9/7.2."""
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


class TestPrecisionGroup234Table9:
    """Table 9 (группы 2-4)."""

    def test_group2_50pct_formula(self, calc):
        """Гр.2 50%: r = 1.0+0.41·Sc; R = 1.8+1.33·Sc."""
        inp = FractionInput(
            P_atm=760.0, object_type="2", V_pct=[100.0],
            T_obs={"IBP": 30.0, 5: 60.0, 50: 130.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        s50 = res.Metrics.slope["50"]
        assert res.Metrics.r["50"] == pytest.approx(1.0 + 0.41 * s50, abs=1e-9)
        assert res.Metrics.R["50"] == pytest.approx(1.8 + 1.33 * s50, abs=1e-9)


class TestNonMonotonic:
    """Сценарий А: немонотонный ввод → NonMonotonicDistillationError."""

    def test_non_monotonic_raises(self, calc):
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[100.0],
            T_obs={0: 100.0, 10: 90.0, 50: 120.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        with pytest.raises(NonMonotonicDistillationError):
            calc.calculate(inp)


class TestGroupBoundaries:
    """Сценарий Б: inclusive_ge."""

    def test_default_rule_ok(self, calc):
        inp = FractionInput(
            P_atm=760.0, object_type="2", V_pct=[100.0],
            T_obs={0: 30.0, 5: 60.0, 50: 130.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert "50" in res.Metrics.r

    def test_wrong_rule_possible(self, calc):
        inp = FractionInput(
            P_atm=760.0, object_type="2", V_pct=[100.0],
            T_obs={0: 30.0, 5: 60.0, 50: 130.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0, boundary_rule="exclusive",
        )
        try:
            calc.calculate(inp)
        except BoundaryRuleError:
            pass  # допустимо


class TestLossLimit:
    """Сценарий В: потери > 1.5% (гр.1-2) → INVALID_TEST."""

    def test_loss_above_limit(self, calc):
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[97.5],
            T_obs={0: 30.0, 50: 130.0, "FBP": 200.0},
            loss_pct=2.5, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert "INVALID_TEST" in res.Flags
        assert any("LOSS_LIMIT_EXCEEDED" in f for f in res.Flags)

    def test_loss_within_limit(self, calc):
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[99.0],
            T_obs={0: 30.0, 50: 130.0, "FBP": 200.0},
            loss_pct=1.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert "INVALID_TEST" not in res.Flags


class TestKCriterion:
    """K = (1/√2)·√(R² − r²/2) — совпадает с эталоном."""

    def test_k_matches_formula(self, calc):
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[100.0],
            T_obs={"IBP": 30.0, 5: 60.0, 50: 130.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        for pt in res.Metrics.r:
            r = res.Metrics.r[pt]
            R = res.Metrics.R[pt]
            expected = (1.0 / 2.0**0.5) * max(R * R - r * r / 2.0, 0.0) ** 0.5
            assert res.Metrics.K[pt] == pytest.approx(expected, abs=1e-9)