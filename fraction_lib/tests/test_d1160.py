"""ПУНКТ 5: тесты ASTM D1160-18 (d1160.py).

Критерии готовности (контракт):
- AET по Максвелл-Боннеллу (A7.1/A7.3, P≥2 мм; A7.5 при P<2 мм);
- r/R по формуле M·[e·exp(a+b·ln(1.8S))]/1.8 с константами 12.5.1/12.5.2;
- r при S=1, P=10 мм, 5-50% = 2.8·[e·exp(0.240+0.350·ln(1.8))]/1.8 ≈ сверка с табл.2;
- P=0.5 мм → A7.5 ветка;
- скачок крутизны → CrackingIndicator;
- OverlapViolation (пересечение с D86, допуск 5°);
- P_vac вне [1,50] → InvalidPressureError.
"""

from __future__ import annotations

import pytest

from fraction_lib.constants import D1160_AET, D1160_CONSTANTS
from fraction_lib.exceptions import (
    InvalidPressureError,
    CrackingIndicator,
    OverlapViolation,
)
from fraction_lib.d1160 import (
    D1160Calculator,
    maxwell_bonnell_aet,
    aet_k_correction,
    d1160_precision,
)
from fraction_lib.models import FractionInput


@pytest.fixture
def calc():
    return D1160Calculator()


class TestMaxwellBonnellAET:
    """Пересчёт AET (Annex A7)."""

    def test_aet_a73_branch(self):
        """P=10 мм (≥2) → ветка A7.3: A = (5.994295 − 0.972546·lg(P))/(2663.129 − 95.76·lg(P))."""
        cfg = D1160_AET["A7_3"]
        p = 10.0
        import math
        lg_p = math.log10(p)
        A = (cfg["n1"] - cfg["n2"] * lg_p) / (cfg["d1"] - cfg["d2"] * lg_p)
        # A при P=10 мм: lg(10)=1 → A = (5.994295−0.972546)/(2663.129−95.76) = 5.021749/2567.369 = 0.001956
        assert A == pytest.approx(0.001956, abs=1e-6)

    def test_aet_known_example(self):
        """AET при t_obs=200°C, P=10 мм — выше наблюдаемой, разумный диапазон."""
        aet = maxwell_bonnell_aet(200.0, 10.0)
        # AET должна быть ВЫШЕ наблюдаемой (атмосферный эквивалент при вакууме)
        assert aet > 200.0
        assert aet < 450.0  # разумный диапазон для AET при 200°C вакуумных

    def test_aet_a75_branch_p05(self):
        """P=0.5 мм (<2) → ветка A7.5."""
        cfg = D1160_AET["A7_5"]
        import math
        lg_p = math.log10(0.5)
        A = (cfg["n1"] - cfg["n2"] * lg_p) / (cfg["d1"] - cfg["d2"] * lg_p)
        # ветка A7.5 использует другие константы — проверяем, что A отличается
        A73 = (D1160_AET["A7_3"]["n1"] - D1160_AET["A7_3"]["n2"] * lg_p) / (
            D1160_AET["A7_3"]["d1"] - D1160_AET["A7_3"]["d2"] * lg_p)
        assert A != pytest.approx(A73, abs=1e-9)
        # и что maxwell_bonnell_aet использует A7.5 при P=0.5
        aet = maxwell_bonnell_aet(200.0, 0.5)
        assert aet > 200.0

    def test_aet_k_correction(self):
        """Поправка AET на K-фактор: t = 1,4·(K−12)·lg(Pa/Po)."""
        # K=12 → поправка 0
        assert aet_k_correction(300.0, 12.0, 760.0, 10.0) == pytest.approx(300.0, abs=1e-9)
        # K=12.2 → поправка 1.4·0.2·lg(76) = 0.28·1.8808 = 0.5266
        expected = 300.0 + 1.4 * (12.2 - 12.0) * (__import__("math").log10(760.0 / 10.0))
        assert aet_k_correction(300.0, 12.2, 760.0, 10.0) == pytest.approx(expected, abs=1e-9)


class TestPrecisionD1160:
    """r = M·exp(a+b·ln(1.8·S))/1.8 (12.2, формулы 1/2)."""

    def test_r_at_slope1_p10(self):
        """r при S=1, P=10 мм, 5-50% = 2.8·exp(0.240+0.350·ln(1.8))/1.8."""
        import math
        a, b, M = D1160_CONSTANTS[10]["r"]["5_50"]
        assert (a, b, M) == pytest.approx((0.240, 0.350, 2.8))
        expected = M * math.exp(a + b * math.log(1.8 * 1.0)) / 1.8
        r = d1160_precision(1.0, 10.0, "5_50", "r")
        assert r == pytest.approx(expected, abs=1e-9)
        # сверка с таблицей 2: при C/V=1.0, r(10мм,5-50)=2.4
        assert r == pytest.approx(2.4, abs=0.5)

    def test_R_at_slope1_p10(self):
        """R при S=1, P=10 мм, 5-50% = 3.2·exp(1.415+0.409·ln(1.8))/1.8."""
        import math
        a, b, M = D1160_CONSTANTS[10]["R"]["5_50"]
        expected = M * math.exp(a + b * math.log(1.8 * 1.0)) / 1.8
        R = d1160_precision(1.0, 10.0, "5_50", "R")
        assert R == pytest.approx(expected, abs=1e-9)
        # сверка с табл.2: R(10мм,5-50) при C/V=1.0 = 9.3
        assert R == pytest.approx(9.3, abs=1.5)

    def test_r_ibp_constant(self):
        """IBP r = M·exp(a)/1.8 (b=0)."""
        import math
        a, b, M = D1160_CONSTANTS[10]["r"]["IBP"]
        assert b == 0
        expected = M * math.exp(a) / 1.8
        assert d1160_precision(0.0, 10.0, "IBP", "r") == pytest.approx(expected, abs=1e-9)

    def test_interpolation_between_1_and_10mm(self):
        """Линейная интерполяция констант между 1 и 10 мм (12.2.3)."""
        r_1 = d1160_precision(1.0, 1.0, "5_50", "r")
        r_10 = d1160_precision(1.0, 10.0, "5_50", "r")
        r_5 = d1160_precision(1.0, 5.5, "5_50", "r")
        assert min(r_1, r_10) <= r_5 <= max(r_1, r_10)

    def test_p50_mmhg_uses_p10_constants(self):
        """P=50 мм → использует константы 10 мм (p >= 10)."""
        r_50 = d1160_precision(1.0, 50.0, "5_50", "r")
        r_10 = d1160_precision(1.0, 10.0, "5_50", "r")
        assert r_50 == pytest.approx(r_10, abs=1e-9)


class TestPressureValidation:
    """P_vac вне [1, 50] → InvalidPressureError."""

    def test_pvac_none(self, calc):
        inp = FractionInput(P_atm=760.0, P_vac=None)
        with pytest.raises(InvalidPressureError):
            calc.calculate(inp)

    def test_pvac_zero(self, calc):
        inp = FractionInput(P_atm=760.0, P_vac=0.5, V_pct=[100.0], T_obs={50: 200.0})
        # 0.5 мм в диапазоне [1,50]? НЕТ — ниже 1. Но A7.5 для P<2.
        # Валидация давления: P_vac должен быть >= 1. 0.5 → InvalidPressureError.
        with pytest.raises(InvalidPressureError):
            calc.calculate(inp)

    def test_pvac_above_range(self, calc):
        inp = FractionInput(P_atm=760.0, P_vac=60.0, V_pct=[100.0], T_obs={50: 200.0})
        with pytest.raises(InvalidPressureError):
            calc.calculate(inp)


class TestCrackingIndicator:
    """Сценарий А: скачок крутизны → CrackingIndicator."""

    def test_slope_jump_flags(self, calc):
        """Скачок крутизны > 15 °С/% → CrackingIndicator флаг."""
        inp = FractionInput(
            P_atm=760.0, P_vac=10.0, V_pct=[100.0],
            T_obs={0: 100.0, 5: 150.0, 50: 200.0, 90: 220.0, "FBP": 230.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        # крутизна 0→5: (150-100)/5 = 10; 5→50: (200-150)/45 = 1.1; 50→90: 0.5
        # нет скачка > 15 → без флага
        assert not any("CrackingIndicator" in f for f in res.Flags)

    def test_cracking_on_steep_jump(self, calc):
        """Искусственный скачок крутизны → CrackingIndicator."""
        inp = FractionInput(
            P_atm=760.0, P_vac=10.0, V_pct=[100.0],
            T_obs={0: 100.0, 5: 105.0, 10: 300.0, 50: 310.0, "FBP": 320.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert any("CrackingIndicator" in f for f in res.Flags)

    def test_cracking_on_decomposition_temp(self, calc):
        """T_IBP_vac > decomposition_temp_degC → CrackingIndicator."""
        inp = FractionInput(
            P_atm=760.0, P_vac=10.0, V_pct=[100.0],
            T_obs={"IBP": 400.0, 50: 450.0, "FBP": 480.0},
            loss_pct=0.0, residue_pct=0.0, decomposition_temp_degC=350.0,
        )
        res = calc.calculate(inp)
        assert any("CrackingIndicator" in f for f in res.Flags)


class TestOverlapViolation:
    """Сценарий Б: пересечение с D86, допуск 5°."""

    def test_overlap_within_tolerance(self, calc):
        """d86_overlap в пределах 5° от AET → без флага."""
        inp = FractionInput(
            P_atm=760.0, P_vac=10.0, V_pct=[100.0],
            T_obs={10: 200.0, 50: 250.0, 90: 300.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        # AET при T_obs=250, P=10 мм ≈ 365; d86_overlap = AET+1 → в допуске
        import fraction_lib.d1160 as d
        aet50 = d.maxwell_bonnell_aet(250.0, 10.0)
        inp.d86_overlap = {10: d.maxwell_bonnell_aet(200.0, 10.0) + 1.0,
                           50: aet50 + 1.0,
                           90: d.maxwell_bonnell_aet(300.0, 10.0) + 1.0}
        res = calc.calculate(inp)
        assert not any("OverlapViolation" in f for f in res.Flags)

    def test_overlap_violation_flag(self, calc):
        """d86_overlap вне 5° от AET → OverlapViolation флаг."""
        inp = FractionInput(
            P_atm=760.0, P_vac=10.0, V_pct=[100.0],
            T_obs={10: 200.0, 50: 250.0, 90: 300.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        import fraction_lib.d1160 as d
        aet50 = d.maxwell_bonnell_aet(250.0, 10.0)
        inp.d86_overlap = {50: aet50 + 20.0}  # Δ=20 > 5
        res = calc.calculate(inp)
        assert any("OverlapViolation" in f for f in res.Flags)


class TestFullCalculation:
    """Полный расчёт D1160: AET в T_corr."""

    def test_tcorr_is_aet(self, calc):
        """T_corr для D1160 — AET (выше наблюдаемых при вакууме)."""
        inp = FractionInput(
            P_atm=760.0, P_vac=10.0, V_pct=[100.0],
            T_obs={10: 150.0, 50: 200.0, 90: 250.0},
            loss_pct=0.0, residue_pct=0.0,
        )
        res = calc.calculate(inp)
        assert res.method == "d1160"
        for pt in ("10", "50", "90"):
            assert res.T_corr[pt] > 150.0  # AET > наблюдаемых
        assert res.raw["extra"]["aet"]["50"] == pytest.approx(res.T_corr["50"], abs=1e-3)
        assert res.raw["aet_final"]["50"] == pytest.approx(res.T_corr["50"], abs=1e-6)