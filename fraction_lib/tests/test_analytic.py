"""ПУНКТ 6 (Способ 1): аналитическая валидация.

Синтетические данные. Проверки:
- материальный баланс ΣV = 100 ± 0,2;
- монотонность T(V_{i+1}) >= T(V_i);
- гладкость (2-я разность без скачков > 3·средняя);
- границы T_НК <= T_10% <= ... <= T_КК.
"""

from __future__ import annotations

import pytest

from fraction_lib.iso3405 import ISO3405Calculator
from fraction_lib.d86 import D86Calculator
from fraction_lib.gost2177 import GOST2177Calculator
from fraction_lib.models import FractionInput
from fraction_lib.validation import (
    check_mass_balance,
    check_monotonicity,
    check_smoothness,
    check_temperature_bounds,
)


def synthetic_iso_input() -> FractionInput:
    """Синтетические данные ISO 3405 (группа 1, P=760)."""
    return FractionInput(
        P_atm=760.0,
        object_type="1",
        V_pct=[99.5],  # отгон 99.5 + потери 0.5 + остаток 0 = 100 (баланс)
        T_obs={
            "IBP": 35.0, 5: 55.0, 10: 70.0, 20: 90.0, 30: 105.0,
            40: 118.0, 50: 130.0, 60: 142.0, 70: 155.0, 80: 170.0,
            90: 188.0, 95: 195.0, "FBP": 210.0,
        },
        loss_pct=0.5,
        residue_pct=0.0,
    )


class TestAnalyticMassBalance:
    """Материальный баланс: Σ = 100 ± 0,2."""

    def test_mass_balance_iso_synthetic(self):
        res = ISO3405Calculator().calculate(synthetic_iso_input())
        assert check_mass_balance(res)

    def test_mass_balance_broken_detected(self):
        """Баланс 98% → обнаружен валидатором (исключение при расчёте)."""
        inp = FractionInput(P_atm=760.0, V_pct=[98.0], T_obs={50: 120.0},
                            loss_pct=0.0, residue_pct=0.0)
        from fraction_lib.exceptions import MassBalanceViolation
        with pytest.raises(MassBalanceViolation):
            GOST2177Calculator().calculate(inp)

    def test_mass_balance_check_function(self):
        class FakeRes:
            Loss = 2.5
            Residue = 1.0
        # 100 - 2.5 - 1.0 = 96.5; 96.5+2.5+1.0 = 100 → ок
        assert check_mass_balance(FakeRes())


class TestAnalyticMonotonicity:
    """Монотонность T(V_{i+1}) >= T(V_i)."""

    def test_monotonic_synthetic_iso(self):
        res = ISO3405Calculator().calculate(synthetic_iso_input())
        assert check_monotonicity(res.T_corr)

    def test_monotonic_broken_detected(self):
        """Немонотонный ряд → False."""
        assert not check_monotonicity({"10": 100.0, "20": 90.0, "30": 110.0})

    def test_monotonic_ok(self):
        assert check_monotonicity({"10": 90.0, "20": 100.0, "30": 110.0})


class TestAnalyticSmoothness:
    """Гладкость: 2-я разность без скачков > 3·средняя."""

    def test_smooth_iso_curve(self):
        res = ISO3405Calculator().calculate(synthetic_iso_input())
        assert check_smoothness(res.T_corr)

    def test_smooth_linear_curve(self):
        """Линейная кривая → вторая разность 0 → гладкая."""
        t_corr = {str(i): 100.0 + i * 2.0 for i in range(0, 101, 10)}
        assert check_smoothness(t_corr)

    def test_jumpy_curve_detected(self):
        """Скачок в середине → негладкая."""
        t_corr = {str(i): 100.0 + i for i in range(0, 40, 10)}
        t_corr["50"] = 300.0  # скачок
        t_corr.update({str(i): 305.0 + (i - 50) for i in range(60, 101, 10)})
        assert not check_smoothness(t_corr)


class TestAnalyticBounds:
    """Границы: T_НК <= T_10% <= ... <= T_КК."""

    def test_bounds_iso(self):
        res = ISO3405Calculator().calculate(synthetic_iso_input())
        assert check_temperature_bounds(res.T_corr)

    def test_bounds_gost(self):
        inp = FractionInput(
            P_atm=760.0, V_pct=[100.0],
            T_obs={"IBP": 40.0, 10: 70.0, 50: 120.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = GOST2177Calculator().calculate(inp)
        assert check_temperature_bounds(res.T_corr)

    def test_bounds_d86(self):
        inp = FractionInput(
            P_atm=760.0, object_type="1", V_pct=[100.0],
            T_obs={"IBP": 35.0, 10: 70.0, 50: 130.0, 90: 188.0, "FBP": 210.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = D86Calculator().calculate(inp)
        assert check_temperature_bounds(res.T_corr)


class TestAnalyticResultSanity:
    """Общая проверка результата расчёта на синтетике."""

    def test_all_metrics_present(self):
        res = ISO3405Calculator().calculate(synthetic_iso_input())
        assert res.Metrics.r
        assert res.Metrics.R
        assert res.Metrics.slope
        assert res.Metrics.K
        assert res.Metrics.R_ok
        for pt in res.Metrics.r:
            assert res.Metrics.K[pt] >= 0
            assert res.Metrics.R_ok[pt] == pytest.approx(res.Metrics.R[pt] * 0.84, abs=1e-9)

    def test_flags_no_invalid(self):
        res = ISO3405Calculator().calculate(synthetic_iso_input())
        assert "INVALID_TEST" not in res.Flags