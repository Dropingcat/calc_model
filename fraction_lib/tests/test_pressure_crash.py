# -*- coding: utf-8 -*-
"""Краш-тесты давления: расширенный диапазон до 800+ мм рт.ст.

Покрывает реальные дефекты, найденные Монте-Карло/краш-пробами:
  1. ISO 3405 loss-формула: полюс при ~820 мм -> Loss = 38059 (было) -> теперь loss_pct.
  2. D86 loss-формула: отрицательные потери при P>797 (было) -> теперь loss_pct.
  3. ГОСТ 2177: P вне [560,760] -> PressureOutOfRangeError (экстраполяция запрещена).
  4. D1160: P_vac вне [1,50] -> InvalidPressureError.
  5. Экстремальные P (до 2000): расчёт не даёт NaN/бесконечность/астрономию.
"""
from __future__ import annotations

import math

import pytest

from fraction_lib.iso3405 import ISO3405Calculator
from fraction_lib.d86 import D86Calculator
from fraction_lib.gost2177 import GOST2177Calculator
from fraction_lib.d1160 import D1160Calculator
from fraction_lib.models import FractionInput
from fraction_lib.exceptions import PressureOutOfRangeError, InvalidPressureError

ISO_T = {"IBP": 119.25, 5: 150.0, 10: 165.75, 20: 181.25, 30: 196.0, 40: 221.25,
         50: 242.5, 60: 260.0, 70: 276.75, 80: 295.25, 90: 313.25, 95: 324.0, "FBP": 355.0}
D86_T = {"IBP": 161.75, 5: 190.5, 10: 203.75, 20: 227.5, 30: 244.0, 40: 256.5,
         50: 268.75, 60: 280.5, 70: 293.75, 80: 309.25, 90: 329.75, 95: 345.25, "FBP": 355.0}
GOST_T = {"IBP": 131.0, 5: 152.5, 10: 158.25, 20: 168.5, 30: 173.5, 40: 179.75,
          50: 185.5, 60: 191.75, 70: 198.5, 80: 206.5, 90: 216.5, 95: 224.0, "FBP": 234.25}
D1160_T = {"IBP": 49.75, 5: 60.25, 10: 70.25, 20: 80.25, 30: 90.25, 40: 100.25,
           50: 110.25, 60: 120.25, 70: 130.25, 80: 140.25, 90: 150.25, 95: 160.25, "FBP": 170.25}


def mk_iso(P, loss=0.3):
    return FractionInput(P_atm=P, object_type="4", T_obs=ISO_T,
                         loss_pct=loss, residue_pct=2.0, V_pct=[100.0 - loss - 2.0])


def mk_d86(P, loss=0.3):
    return FractionInput(P_atm=P, object_type="4", T_obs=D86_T,
                         loss_pct=loss, residue_pct=1.5, V_pct=[98.5 - loss])


def mk_gost(P):
    return FractionInput(P_atm=P, T_obs=GOST_T,
                         loss_pct=0.5, residue_pct=1.0, V_pct=[98.5])


def mk_d1160(Pvac):
    return FractionInput(P_atm=750.0, P_vac=Pvac, T_obs=D1160_T,
                         loss_pct=0.0, residue_pct=0.0, V_pct=[100.0])


class TestISOLossPole:
    """ISO 3405: полюс loss-формулы при P~820 мм."""

    @pytest.mark.parametrize("P", [800.0, 810.0, 819.0, 819.82, 820.0, 821.0])
    def test_no_explosion_at_high_pressure(self, P):
        res = ISO3405Calculator().calculate(mk_iso(P))
        assert 0.0 <= res.Loss <= 5.0, f"Loss={res.Loss} при P={P} (аномалия)"
        assert not math.isnan(res.Loss)

    @pytest.mark.parametrize("P", [800.0, 820.0, 900.0, 1500.0, 2000.0])
    def test_loss_equals_observed_above_760(self, P):
        """При P > 760 мм коррекция потерь не применяется -> Loss = loss_pct."""
        res = ISO3405Calculator().calculate(mk_iso(P, loss=0.3))
        assert abs(res.Loss - 0.3) < 1e-9, f"Loss={res.Loss}, ожидалось 0.3"


class TestD86LossPole:
    """D86: отрицательные потери при P>797 и полюс при 820."""

    @pytest.mark.parametrize("P", [790.0, 797.0, 800.0, 810.0, 819.0, 820.0, 821.0, 900.0])
    def test_no_negative_loss(self, P):
        res = D86Calculator().calculate(mk_d86(P))
        assert 0.0 <= res.Loss <= 5.0, f"Loss={res.Loss} при P={P}"

    @pytest.mark.parametrize("P", [800.0, 900.0, 2000.0])
    def test_loss_equals_observed_above_760(self, P):
        """TD-D86-001: при P > 760 коррекция НЕ отключается (единая формула D86-23).

        Проверяем: Lc конечна, ≠ наблюдаемым потерям (коррекция применена),
        без полюса/отрицательных значений вне зоны сингулярности 810–830 мм.
        """
        res = D86Calculator().calculate(mk_d86(P, loss=0.7))
        assert not math.isnan(res.Loss) and res.Loss >= 0.0
        if not (810.0 <= P <= 830.0):
            assert abs(res.Loss - 0.7) > 1e-9, f"Lc={res.Loss}: коррекция молча не применена при P={P}"


class TestGostPressureRange:
    """ГОСТ 2177: таблица 4 применима только в [560, 760]."""

    @pytest.mark.parametrize("P", [560.0, 700.0, 760.0])
    def test_within_range_ok(self, P):
        res = GOST2177Calculator().calculate(mk_gost(P))
        # потери корректируются по таблице 4 (V_K = A·L+B); при P=760: A=1,B=0 -> 0.5
        assert 0.0 <= res.Loss <= 1.0, f"Loss={res.Loss} при P={P}"
        if P == 760.0:
            assert abs(res.Loss - 0.5) < 1e-9

    @pytest.mark.parametrize("P", [559.9, 500.0, 760.1, 800.0, 850.0])
    def test_out_of_range_raises(self, P):
        with pytest.raises(PressureOutOfRangeError):
            GOST2177Calculator().calculate(mk_gost(P))


class TestD1160VacRange:
    """D1160: P_vac в [1, 50] мм."""

    @pytest.mark.parametrize("Pvac", [1.0, 2.0, 10.0, 50.0])
    def test_within_range_ok(self, Pvac):
        res = D1160Calculator().calculate(mk_d1160(Pvac))
        assert res.T_corr.get("50") is not None

    @pytest.mark.parametrize("Pvac", [0.13, 0.5, 0.99, 50.1, 51.0, 100.0])
    def test_out_of_range_raises(self, Pvac):
        with pytest.raises(InvalidPressureError):
            D1160Calculator().calculate(mk_d1160(Pvac))


class TestExtremePressureNoNaN:
    """Экстремальные давления: нет NaN/бесконечности/астрономических значений."""

    @pytest.mark.parametrize("P", [800.0, 900.0, 1000.0, 1500.0, 2000.0])
    def test_iso_extreme(self, P):
        res = ISO3405Calculator().calculate(mk_iso(P))
        for v in res.T_corr.values():
            if v is not None:
                assert math.isfinite(v), f"T_corr={v} при P={P}"
        assert math.isfinite(res.Loss)

    @pytest.mark.parametrize("P", [800.0, 900.0, 1000.0, 1500.0, 2000.0])
    def test_d86_extreme(self, P):
        res = D86Calculator().calculate(mk_d86(P))
        for v in res.T_corr.values():
            if v is not None:
                assert math.isfinite(v), f"T_corr={v} при P={P}"
        assert math.isfinite(res.Loss)