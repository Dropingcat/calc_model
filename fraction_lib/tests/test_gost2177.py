"""ПУНКТ 2: тесты ГОСТ 2177-99 (gost2177.py).

Критерии готовности (контракт):
- P=760 → V_K = L;
- L=2.5, P=740 → V_K = 2.0;
- P=500 → PressureOutOfRangeError;
- баланс 98% → MassBalanceViolation;
- |X1−X2| > r → RepeatabilityViolation (сценарий В);
- поправка давления Сиднея Янга: C = 0,00012·(760 − P)·(273 + t).
"""

from __future__ import annotations

import pytest

from fraction_lib.constants import GOST2177_TABLE4, GOST2177_NOMOGRAM, interpolate_table
from fraction_lib.exceptions import (
    PressureOutOfRangeError,
    MassBalanceViolation,
    RepeatabilityViolation,
)
from fraction_lib.gost2177 import GOST2177Calculator
from fraction_lib.models import FractionInput


@pytest.fixture
def calc():
    return GOST2177Calculator()


class TestLossCorrection:
    """V_K = A·L + B (таблица 4, линейная интерполяция)."""

    def test_p760_vk_equals_loss(self, calc):
        """P=760 → V_K = L (A=1.000, B=0.000)."""
        inp = FractionInput(P_atm=760.0)
        assert calc._correct_loss(2.5, inp) == pytest.approx(2.5, abs=1e-9)

    def test_p740_loss25_vk20(self, calc):
        """L=2.5, P=740 → V_K = 0.750·2.5 + 0.125 = 2.0 (пример контракта)."""
        inp = FractionInput(P_atm=740.0)
        assert calc._correct_loss(2.5, inp) == pytest.approx(2.0, abs=1e-9)

    def test_p560_vk_table4_first_row(self, calc):
        """P=560 (граница) → A=0.231, B=0.384: V_K = 0.231·L + 0.384."""
        inp = FractionInput(P_atm=560.0)
        assert calc._correct_loss(10.0, inp) == pytest.approx(0.231 * 10.0 + 0.384, abs=1e-9)

    def test_p_interpolation_linear(self, calc):
        """Линейная интерполяция между 740 и 750: P=745 → среднее A/B."""
        A740, B740 = 0.750, 0.125
        A750, B750 = 0.857, 0.071
        inp = FractionInput(P_atm=745.0)
        vk = calc._correct_loss(1.0, inp)
        expected_A = (A740 + A750) / 2.0
        expected_B = (B740 + B750) / 2.0
        assert vk == pytest.approx(expected_A * 1.0 + expected_B, abs=1e-9)

    def test_interpolate_table_boundaries(self):
        """interpolate_table: границы 560 и 760 возвращают крайние строки."""
        A, B = interpolate_table(560.0, GOST2177_TABLE4)
        assert A == pytest.approx(0.231)
        assert B == pytest.approx(0.384)
        A, B = interpolate_table(760.0, GOST2177_TABLE4)
        assert A == pytest.approx(1.000)
        assert B == pytest.approx(0.000)


class TestPressureScenario:
    """Сценарий А: давление вне [560, 760] → PressureOutOfRangeError."""

    def test_p500_out_of_range(self, calc):
        inp = FractionInput(P_atm=500.0, V_pct=[100.0], T_obs={50: 120.0})
        with pytest.raises(PressureOutOfRangeError):
            calc.calculate(inp)

    def test_p800_out_of_range(self, calc):
        inp = FractionInput(P_atm=800.0, V_pct=[100.0], T_obs={50: 120.0})
        with pytest.raises(PressureOutOfRangeError):
            calc.calculate(inp)

    def test_correct_loss_raises_out_of_range(self, calc):
        inp = FractionInput(P_atm=500.0)
        with pytest.raises(PressureOutOfRangeError):
            calc._correct_loss(2.5, inp)


class TestTable4PolicyTD2177002:
    """TD-2177-002: формальная политика A/B для нетабличного давления."""

    FULL_CURVE = {"IBP": 40.0, 10: 60.0, 50: 90.0, 90: 120.0, "FBP": 140.0}

    def _inp(self, p_mmhg):
        # баланс: отгон 98 + потери 2 + остаток 0 = 100
        return FractionInput(
            P_atm=p_mmhg, V_pct=[98.0], T_obs=dict(self.FULL_CURVE),
            loss_pct=2.0, residue_pct=0.0,
        )

    def test_default_policy_is_strict(self):
        c = GOST2177Calculator()
        assert c.table4_policy == "strict"

    def test_unknown_policy_rejected(self):
        with pytest.raises(ValueError):
            GOST2177Calculator(table4_policy="clamp_high")

    def test_strict_p500_still_fails(self):
        c = GOST2177Calculator(table4_policy="strict")
        with pytest.raises(PressureOutOfRangeError):
            c.calculate(self._inp(500.0))

    def test_clamp_low_p500_uses_row_560_with_flag(self):
        c = GOST2177Calculator(table4_policy="clamp_low")
        res = c.calculate(self._inp(500.0))
        assert "PRESSURE_CLAMPED" in res.Flags
        # Vk = 0.231*L + 0.384 при L=2.0 → 0.846 (res.Loss — скорректированные потери)
        assert res.Loss == pytest.approx(0.231 * 2.0 + 0.384, abs=1e-9)

    def test_clamp_low_in_range_no_flag(self):
        c = GOST2177Calculator(table4_policy="clamp_low")
        res = c.calculate(self._inp(700.0))
        assert "PRESSURE_CLAMPED" not in res.Flags

    def test_clamp_low_above_760_always_fails(self):
        c = GOST2177Calculator(table4_policy="clamp_low")
        with pytest.raises(PressureOutOfRangeError):
            c.calculate(self._inp(800.0))


class TestMassBalanceScenario:
    """Сценарий Б: отгон+остаток+потери != 100 ± 0.2 → MassBalanceViolation."""

    def test_balance_98_violation(self, calc):
        """Баланс 98% → MassBalanceViolation."""
        inp = FractionInput(
            P_atm=760.0,
            V_pct=[98.0],
            T_obs={50: 120.0},
            loss_pct=0.0,
            residue_pct=0.0,
        )
        with pytest.raises(MassBalanceViolation):
            calc.calculate(inp)

    def test_balance_100_ok(self, calc):
        """Баланс 100% → расчёт проходит."""
        inp = FractionInput(
            P_atm=760.0,
            V_pct=[100.0],
            T_obs={50: 120.0},
            loss_pct=0.0,
            residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert res.Loss == pytest.approx(0.0, abs=1e-9)

    def test_balance_auto_normalize(self, calc):
        """auto_normalize=True: баланс 98% → нормализация + флаг."""
        inp = FractionInput(
            P_atm=760.0,
            V_pct=[98.0],
            T_obs={50: 120.0},
            loss_pct=0.0,
            residue_pct=0.0,
            auto_normalize=True,
        )
        res = calc.calculate(inp)
        assert "MASS_BALANCE_NORMALIZED" in res.Flags


class TestRepeatabilityScenario:
    """Сценарий В: |X1 − X2| > r → RepeatabilityViolation."""

    def test_pairs_above_r(self, calc):
        """Два определения, |X1−X2| > r → RepeatabilityViolation."""
        inp = FractionInput(
            P_atm=760.0,
            V_pct=[100.0],
            T_obs={50: [100.0, 115.0]},  # |15| > r (r≈3 при slope=0)
            loss_pct=0.0,
            residue_pct=0.0,
        )
        with pytest.raises(RepeatabilityViolation):
            calc.calculate(inp)

    def test_pairs_within_r_ok(self, calc):
        """Два определения, |X1−X2| <= r → расчёт проходит."""
        inp = FractionInput(
            P_atm=760.0,
            V_pct=[100.0],
            T_obs={50: [100.0, 101.0]},
            loss_pct=0.0,
            residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert res.T_corr["50"] is not None


class TestPressureCorrection:
    """Поправка давления Сиднея Янга (метод А, формула 1)."""

    def test_young_correction_formula(self, calc):
        """C = 0,00012·(760 − P)·(273 + t). P=740, t=100 → C=0.8952."""
        inp = FractionInput(P_atm=740.0)
        # прямой вызов приватного метода — юнит-проверка формулы
        corr = calc._apply_pressure_correction(100.0, 740.0)
        expected = 100.0 + 0.00012 * (760.0 - 740.0) * (273.0 + 100.0)
        assert corr == pytest.approx(expected, abs=1e-9)

    def test_young_at_std_pressure_zero(self, calc):
        """P=760 → поправка = 0 (T_corr == T_obs)."""
        inp = FractionInput(
            P_atm=760.0,
            V_pct=[100.0],
            T_obs={"IBP": 40.0, 50: 120.0, "FBP": 190.0},
            loss_pct=0.0,
            residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert res.T_corr["IBP"] == pytest.approx(40.0, abs=1e-9)
        assert res.T_corr["50"] == pytest.approx(120.0, abs=1e-9)
        assert res.T_corr["FBP"] == pytest.approx(190.0, abs=1e-9)

    def test_full_calculation_result_structure(self, calc):
        """Полный расчёт: структура результата и метод."""
        inp = FractionInput(
            P_atm=760.0,
            V_pct=[0, 10, 50, 90, 100],
            T_obs={0: 40.0, 10: 70.0, 50: 120.0, 90: 180.0, "FBP": 200.0},
            loss_pct=0.0,
            residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert res.method == "gost2177"
        assert "50" in res.T_corr
        assert "50" in res.Metrics.r
        assert "50" in res.Metrics.R
        assert "50" in res.Metrics.K
        assert res.Metrics.K["50"] > 0
        assert res.Metrics.R_ok["50"] == pytest.approx(res.Metrics.R["50"] * 0.84, abs=1e-9)


class TestMethodB:
    """Метод Б — фиксированная точность (6.4.1/6.4.2)."""

    def test_method_b_fixed_precision(self, calc):
        inp = FractionInput(
            P_atm=760.0,
            object_type="B",
            V_pct=[100.0],
            T_obs={"IBP": 200.0, 10: 260.0, 50: 320.0, 90: 380.0, "FBP": 400.0},
            loss_pct=0.0,
            residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert res.raw["extra"]["method"] == "B"
        # фиксированные значения: R(50)=3.0, r(50)=2.0
        assert res.Metrics.r["50"] == pytest.approx(2.0, abs=1e-9)
        assert res.Metrics.R["50"] == pytest.approx(3.0, abs=1e-9)

class TestExtrapolationBanTD2177004:
    """TD-2177-004 / ГОСТ 2177 п.5.5.8: жёсткий запрет экстраполяции T(V)."""

    def test_interpolation_below_95_allowed(self, calc):
        """Точка 30% без наблюдения между наблюдёнными 20/50 — интерполяция, ок."""
        inp = FractionInput(
            P_atm=760.0,
            V_pct=[100.0],
            T_obs={20: 60.0, 50: 80.0, 90: 110.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        # интерполяция внутри диапазона наблюдений — без ошибок и флагов
        assert "EXTRAPOLATED_96_98" not in res.Flags

    def test_extrapolated_96_allowed_with_flag(self, calc):
        """96% без наблюдения — легализованная экстраполяция (ветка потерь>=2%) + флаг."""
        inp = FractionInput(
            P_atm=760.0,
            V_pct=[100.0],
            T_obs={"IBP": 40.0, 10: 60.0, 50: 90.0, 90: 120.0, 95: 130.0, 96: None},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert "EXTRAPOLATED_96_98" in res.Flags

    def test_extrapolation_beyond_98_forbidden(self, calc):
        """FBP (100%) без наблюдения — запрещённая экстраполяция > 98%."""
        inp = FractionInput(
            P_atm=760.0,
            V_pct=[100.0],
            T_obs={"IBP": 40.0, 10: 60.0, 50: 90.0, 90: 120.0, 95: 130.0, "FBP": None},
            loss_pct=0.0, residue_pct=0.0,
        )
        from fraction_lib.exceptions import InvalidTestError
        with pytest.raises(InvalidTestError) as ei:
            calc.calculate(inp)
        assert ei.value.code == "EXTRAPOLATION_FORBIDDEN"

    def test_point_left_of_first_observed_forbidden(self, calc):
        """Точка левее первой наблюдённой (< диапазона наблюдений) — запрет."""
        inp = FractionInput(
            P_atm=760.0,
            V_pct=[100.0],
            T_obs={"IBP": None, 10: 60.0, 50: 90.0, 90: 120.0, 95: 130.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        from fraction_lib.exceptions import InvalidTestError
        with pytest.raises(InvalidTestError) as ei:
            calc.calculate(inp)
        assert ei.value.code == "EXTRAPOLATION_FORBIDDEN"

    def test_fully_observed_curve_no_flags(self, calc):
        """Полностью наблюдённая кривая — без ошибок и флагов."""
        inp = FractionInput(
            P_atm=760.0,
            V_pct=[100.0],
            T_obs={"IBP": 40.0, 5: 50.0, 10: 60.0, 50: 90.0, 90: 120.0, 95: 130.0, "FBP": 140.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert res.Flags == []


class TestNomogramValidationTD2177003:
    """TD-2177-003: независимая валидация оцифрованной номограммы (рис. 6)."""

    def test_kt1_node_self_consistency(self):
        # KT-1: интерполяция в узлах сетки даёт точно табличные значения
        from fraction_lib.gost2177 import _nomogram_values
        for k, v in GOST2177_NOMOGRAM.items():
            got = _nomogram_values(k)
            assert all(abs(g - e) < 1e-12 for g, e in zip(got, v)), (k, got, v)

    def test_kt2_midpoint_linearity(self):
        # KT-2: между соседними узлами — строго линейно
        from fraction_lib.gost2177 import _nomogram_values
        ks = sorted(GOST2177_NOMOGRAM.keys())
        for k0, k1 in zip(ks[:-1], ks[1:]):
            m = (k0 + k1) / 2
            expect = tuple((a + b) / 2 for a, b in
                           zip(GOST2177_NOMOGRAM[k0], GOST2177_NOMOGRAM[k1]))
            got = _nomogram_values(m)
            assert all(abs(g - e) < 1e-9 for g, e in zip(got, expect)), (m, got, expect)

    def test_kt3_monotonicity_all_series(self):
        # KT-3: все шесть серий r/R невозрастают с ростом крутизны.
        # НАЙДЕННЫЙ ДЕФЕКТ ОЦИФРОВКИ (TD-2177-003, 2026-10-08): серии r_з и R_з
        # имеют нефизичный скачок вверх S=1.3→1.4 (r_з 2.3→2.7, R_з 5.1→6.0),
        # тогда как остальные четыре серии монотонны. До сверки с бумажной
        # номограммой рис. 6 тест фиксирует дефект списком известных нарушений
        # (whitelist) — если оцифровка правится и дефект исчезает, whitelist
        # надо обнулить; если появятся НОВЫЕ нарушения — тест красный.
        ks = sorted(GOST2177_NOMOGRAM.keys())
        series_names = ["r_нк", "r_кон", "r_з", "R_нк", "R_кон", "R_з"]
        known_defects = {("r_з", 1.3, 1.4), ("R_з", 1.3, 1.4)}
        violations = set()
        for s_idx, name in enumerate(series_names):
            vals = [GOST2177_NOMOGRAM[k][s_idx] for k in ks]
            for i in range(len(vals) - 1):
                if vals[i + 1] > vals[i] + 1e-12:
                    violations.add((name, ks[i], ks[i + 1]))
        new = violations - known_defects
        assert not new, f"новые нарушения монотонности (дефект оцифровки?): {sorted(new)}"

    def test_kt4_range_coverage(self):
        # KT-4: сетка покрывает [0.0, 5.0] с шагом 0.1 (51 узел)
        ks = sorted(GOST2177_NOMOGRAM.keys())
        assert len(ks) == 51 and abs(ks[0]) < 1e-12 and abs(ks[-1] - 5.0) < 1e-12
        assert all(abs((b - a) - 0.1) < 1e-9 for a, b in zip(ks[:-1], ks[1:]))

    def test_clamp_high_sets_flag_not_silent(self):
        # выход за S>5.0: значение = крайнее, но НЕ молча — флаг NOMOGRAM_CLAMPED_HIGH
        from fraction_lib.gost2177 import _nomogram_values
        fl: list[str] = []
        got = _nomogram_values(6.2, fl)
        assert got == GOST2177_NOMOGRAM[5.0]
        assert "NOMOGRAM_CLAMPED_HIGH" in fl

    def test_in_range_no_flag(self):
        from fraction_lib.gost2177 import _nomogram_values
        fl: list[str] = []
        _nomogram_values(2.35, fl)
        assert fl == []
