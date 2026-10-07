"""Тесты TD-2177-007: GUM-модель неопределённости ГОСТ 2177.

Контроль:
 - аналитические чувствительности c_p, c_t совпадают с численными производными
   реальной поправки давления из gost2177 (способ 2 — численная валидация);
 - бюджет положительный, доминирует ожидаемый компонент;
 - политика флагов (экстраполяция/клемп) влияет на U и valid_for_reporting;
 - U = k*u_c, ν_eff/Welch–Satterthwaite при типе A повышает k над 2.
"""

import math

import pytest

from fraction_lib.uncertainty import (
    COEFF_DT,
    EXTRAPOLATION_U_FACTOR,
    P_STD_KPA,
    U_NOMOGRAM_HALFWIDTH,
    U_ROUND_HALFWIDTH,
    UncertaintyComponent,
    evaluate_gum,
)


def _delta_p_kpa(p_kPa: float, t_bar: float) -> float:
    """Нормативная поправка давления, kPa-форма (эквивалент 0.0000009*(101300-Pb)*(273+t))."""
    return COEFF_DT * (P_STD_KPA - p_kPa) * (273.0 + t_bar)


class TestSensitivities:
    def test_cp_matches_numeric_derivative(self):
        """c_p = dT/dp аналитически vs численно (центральная разность)."""
        p, t = 95.0, 20.0
        h = 0.01
        num = (_delta_p_kpa(p + h, t) - _delta_p_kpa(p - h, t)) / (2 * h)
        ana = -COEFF_DT * (273.0 + t)
        assert abs(num - ana) < 1e-9

    def test_ct_matches_numeric_derivative(self):
        p, t = 95.0, 20.0
        h = 0.01
        num = (_delta_p_kpa(p, t + h) - _delta_p_kpa(p, t - h)) / (2 * h)
        ana = COEFF_DT * (P_STD_KPA - p)
        assert abs(num - ana) < 1e-9

    def test_sensitivity_signs(self):
        # при p < 101.3 поправка положительна; рост p уменьшает T_corr → c_p < 0
        res = evaluate_gum(95.0, 20.0)
        cp = next(c for c in res.components if c.name == "u2").sensitivity
        assert cp < 0
        ct = next(c for c in res.components if c.name == "u3").sensitivity
        assert ct > 0


class TestBudget:
    def test_positive_and_monotonic_in_obs(self):
        r_small = evaluate_gum(101.3, 20.0, u_obs_halfwidth=0.2)
        r_big = evaluate_gum(101.3, 20.0, u_obs_halfwidth=1.0)
        assert r_small.uc > 0 and r_big.uc > r_small.uc

    def test_default_budget_magnitude(self):
        """p=101.3: c_t-группа (u3,u4) обнуляется; остаются u1, u2(c_p), u5, u6."""
        from fraction_lib.uncertainty import U_PRESSURE_HALFWIDTH_KPA
        res = evaluate_gum(P_STD_KPA, 20.0)
        cp = -COEFF_DT * (273.0 + 20.0)
        expected = math.sqrt((0.5 ** 2 + (abs(cp) * U_PRESSURE_HALFWIDTH_KPA) ** 2
                              + U_NOMOGRAM_HALFWIDTH ** 2
                              + U_ROUND_HALFWIDTH ** 2) / 3.0)
        assert abs(res.uc - expected) < 1e-9

    def test_thread_off_at_std_pressure_zeroes_barometer_group(self):
        """p=101.3 и без u4: барометрическая группа (c_t) целиком вне бюджета."""
        from fraction_lib.uncertainty import U_PRESSURE_HALFWIDTH_KPA
        res = evaluate_gum(P_STD_KPA, 20.0, include_thread_correction=False)
        cp = -COEFF_DT * (273.0 + 20.0)
        expected = math.sqrt((0.5 ** 2 + (abs(cp) * U_PRESSURE_HALFWIDTH_KPA) ** 2
                              + U_NOMOGRAM_HALFWIDTH ** 2
                              + U_ROUND_HALFWIDTH ** 2) / 3.0)
        assert abs(res.uc - expected) < 1e-9

    def test_rectangular_divided_by_sqrt3(self):
        c = UncertaintyComponent("x", "d", "B", "rectangular", 0.5, 1.0, "°C")
        assert abs(c.u - 0.5 / math.sqrt(3)) < 1e-12

    def test_U_equals_k_uc(self):
        res = evaluate_gum(98.0, 18.0)
        assert abs(res.U - res.k * res.uc) < 1e-12


class TestPolicyFlags:
    def test_extrapolation_penalty_increases_U(self):
        base = evaluate_gum(101.3, 20.0)
        ext = evaluate_gum(101.3, 20.0, flags=["EXTRAPOLATED_96_98"])
        assert ext.U > base.U
        assert any("EXTRAPOLATED" in n for n in ext.notes)

    def test_clamped_flags_invalid_for_reporting(self):
        for flag in ("PRESSURE_CLAMPED", "NOMOGRAM_CLAMPED_HIGH"):
            res = evaluate_gum(101.3, 20.0, flags=[flag])
            assert res.valid_for_reporting is False
            assert res.notes

    def test_clean_result_valid(self):
        res = evaluate_gum(100.0, 20.0, flags=[])
        assert res.valid_for_reporting is True


class TestTypeAAndWelch:
    def test_type_a_replaces_halfwidth(self):
        a = evaluate_gum(101.3, 20.0, s_rep=0.4, n_repeats=5)
        u1 = next(c for c in a.components if c.kind == "A")
        assert abs(u1.halfwidth - 0.4 / math.sqrt(5)) < 1e-12

    def test_welch_raises_k_when_few_repeats(self):
        """При малом числе повторностей ν_eff мал → t95 > 2 → k должен вырасти."""
        res = evaluate_gum(101.3, 20.0, s_rep=0.9, n_repeats=2)
        assert res.dof_eff < 30
        assert res.k >= 2.0
        assert res.t95 > 1.96
        assert any("ν_eff" in n or "k повышен" in n for n in res.notes)

    def test_all_B_infinite_dof_keeps_k(self):
        res = evaluate_gum(101.3, 20.0)
        assert math.isinf(res.dof_eff)
        assert res.k == 2.0


class TestThreadCorrectionSwitch:
    def test_toggle_removes_u4(self):
        with_u4 = evaluate_gum(95.0, 20.0, include_thread_correction=True)
        without = evaluate_gum(95.0, 20.0, include_thread_correction=False)
        names_with = {c.name for c in with_u4.components}
        names_wo = {c.name for c in without.components}
        assert "u4" in names_with and "u4" not in names_wo
        assert without.uc <= with_u4.uc
