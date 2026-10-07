"""ПУНКТ 1: тесты каркаса — models, exceptions, constants.

Критерии готовности (контракт):
- dataclasses создаются с типизацией;
- исключения наследуют Exception (через FractionLibError);
- таблицы имеют ожидаемые размеры (Табл.4: 21 строка, A[560]=0.231, A[760]=1.000).
"""

from __future__ import annotations

import dataclasses

import pytest

from fraction_lib.constants import (
    GOST2177_TABLE4,
    GOST2177_PRESSURE_RANGE,
    ISO3405_G1_SLOPE_COEF,
    ISO3405_G1_R_CORR,
    ISO3405_G234,
    D86_TABLE8,
    D86_TABLE9,
    D1160_CONSTANTS,
    D1160_PRESSURE_RANGE,
    K_CRITERION_COEFF,
)
from fraction_lib.exceptions import (
    PressureOutOfRangeError,
    MassBalanceViolation,
    NonMonotonicDistillationError,
    RepeatabilityViolation,
    LossLimitExceeded,
    InvalidTestError,
    BoundaryRuleError,
    CrackingIndicator,
    OverlapViolation,
    FractionLibError,
)
from fraction_lib.models import (
    FractionInput,
    FractionResult,
    Metrics,
    k_criterion,
)


# ---------------------------------------------------------------------------
# Модели (dataclass с типизацией)
# ---------------------------------------------------------------------------
class TestModels:
    def test_fraction_input_typed_fields(self):
        """FractionInput создаётся и поля типизированы."""
        inp = FractionInput(P_atm=760.0)
        assert inp.P_atm == 760.0
        assert inp.pressure_unit == "mmHg"
        assert inp.P_vac is None
        assert inp.object_type == "1"
        assert inp.method == "iso3405"
        assert inp.V_ml == []
        assert inp.V_pct == []
        assert inp.T_obs == {}
        assert inp.loss_ml is None
        assert inp.residue_ml is None
        assert inp.auto_normalize is False
        assert inp.boundary_rule == "inclusive_ge"
        assert dataclasses.is_dataclass(FractionInput)

    def test_fraction_input_pressure_conversion(self):
        """pressure_mmhg/pressure_kpa: 101.325 кПа == 760 мм рт.ст."""
        inp = FractionInput(P_atm=101.325, pressure_unit="kPa")
        assert inp.pressure_mmhg() == pytest.approx(760.0, abs=0.01)
        inp2 = FractionInput(P_atm=760.0, pressure_unit="mmHg")
        assert inp2.pressure_kpa() == pytest.approx(101.325, abs=0.01)

    def test_fraction_result_typed(self):
        """FractionResult создаётся со всеми полями."""
        res = FractionResult()
        assert res.T_corr == {}
        assert res.V_corr == {}
        assert res.Loss == 0.0
        assert res.Residue == 0.0
        assert isinstance(res.Metrics, Metrics)
        assert res.Flags == []
        assert dataclasses.is_dataclass(FractionResult)

    def test_metrics_typed(self):
        m = Metrics(r={"50": 2.0}, R={"50": 5.0}, slope={"50": 1.5}, K={"50": 3.0}, R_ok={"50": 4.2})
        assert m.r["50"] == 2.0
        assert m.R_ok["50"] == 4.2

    def test_k_criterion_formula_not_changed(self):
        """K = (1/√2)·√(R² − r²/2) — формула ГСО, НЕ меняется."""
        r, R = 2.0, 5.0
        expected = (1.0 / 2.0**0.5) * (R * R - r * r / 2.0) ** 0.5
        assert k_criterion(r, R) == pytest.approx(expected, abs=1e-12)
        # частный случай: r=0 → K = R/√2
        assert k_criterion(0.0, 5.0) == pytest.approx(5.0 / 2**0.5, abs=1e-12)

    def test_to_dict_json_friendly(self):
        """to_dict() возвращает JSON-совместимые структуры."""
        inp = FractionInput(P_atm=740.0, loss_ml=2.5)
        d = inp.to_dict()
        assert d["P_atm"] == 740.0
        assert d["loss_ml"] == 2.5


# ---------------------------------------------------------------------------
# Исключения
# ---------------------------------------------------------------------------
class TestExceptions:
    @pytest.mark.parametrize(
        "exc_cls, kwargs",
        [
            (PressureOutOfRangeError, {"pressure_mmhg": 500.0}),
            (MassBalanceViolation, {"distillate": 90.0, "residue": 5.0, "loss": 3.0}),
            (NonMonotonicDistillationError, {"point": "50", "t_prev": 100.0, "t_next": 90.0,
                                             "v_prev": 40.0, "v_next": 50.0}),
            (RepeatabilityViolation, {"point": "50", "x1": 100.0, "x2": 105.0, "r": 2.0}),
            (LossLimitExceeded, {"loss": 2.5, "limit": 1.5, "group": 1}),
            (InvalidTestError, {"message": "test invalid", "reason": "loss"}),
            (BoundaryRuleError, {"slope": 1.0, "boundary": 1.0, "rule": "inclusive_ge"}),
            (CrackingIndicator, {"message": "crack"}),
            (OverlapViolation, {"pct": 10.0, "t_d86": 100.0, "t_d1160": 110.0}),
        ],
    )
    def test_all_exceptions_inherit_exception(self, exc_cls, kwargs):
        """Все исключения наследуют Exception (через FractionLibError)."""
        assert issubclass(exc_cls, Exception)
        assert issubclass(exc_cls, FractionLibError)
        exc = exc_cls(**kwargs)
        assert isinstance(exc, Exception)
        assert exc.code  # машиночитаемый код
        assert exc.to_dict()["code"] == exc.code

    def test_fraction_lib_error_details(self):
        exc = PressureOutOfRangeError(500.0)
        assert exc.details["pressure_mmhg"] == 500.0
        assert "500" in str(exc)


# ---------------------------------------------------------------------------
# Константы: таблицы НД
# ---------------------------------------------------------------------------
class TestConstants:
    def test_gost2177_table4_size_and_endpoints(self):
        """Табл.4 ГОСТ 2177-99: 21 строка; A[560]=0.231; A[760]=1.000."""
        assert len(GOST2177_TABLE4) == 21
        assert GOST2177_TABLE4[0][0] == 560
        assert GOST2177_TABLE4[0][1] == pytest.approx(0.231)
        assert GOST2177_TABLE4[-1][0] == 760
        assert GOST2177_TABLE4[-1][1] == pytest.approx(1.000)
        assert GOST2177_TABLE4[-1][2] == pytest.approx(0.0)
        assert GOST2177_PRESSURE_RANGE == (560.0, 760.0)

    def test_gost2177_table4_pressure_monotonic(self):
        """Давления в Табл.4 строго возрастают."""
        ps = [row[0] for row in GOST2177_TABLE4]
        assert ps == sorted(ps)

    def test_iso3405_g1_coefficients(self):
        """Таблица 6 гр.1: r₁=0.864S+1.214; R₁=1.736S+1.994."""
        a_r, b_r = ISO3405_G1_SLOPE_COEF["r"]
        a_R, b_R = ISO3405_G1_SLOPE_COEF["R"]
        assert (a_r, b_r) == pytest.approx((0.864, 1.214))
        assert (a_R, b_R) == pytest.approx((1.736, 1.994))
        # поправки к R: 5%:+1.11, 10-80:0, 90:−1.22, 95:−0.94
        assert ISO3405_G1_R_CORR["5"] == pytest.approx(1.11)
        assert ISO3405_G1_R_CORR["10"] == 0.0
        assert ISO3405_G1_R_CORR["80"] == 0.0
        assert ISO3405_G1_R_CORR["90"] == pytest.approx(-1.22)
        assert ISO3405_G1_R_CORR["95"] == pytest.approx(-0.94)

    def test_iso3405_g234_coeffs(self):
        """Таблица 7 (2022): IBP R=0.93S+2.8; 5-95 R=1.33S+1.8; FBP R=0.42S+3.1."""
        assert ISO3405_G234["IBP"]["R"] == pytest.approx((0.93, 2.8))
        assert ISO3405_G234["5_95"]["R"] == pytest.approx((1.33, 1.8))
        assert ISO3405_G234["FBP"]["R"] == pytest.approx((0.42, 3.1))

    def test_d86_table8_values(self):
        """Table 8 гр.1: 5% R=3.1+1.74Sc; 90 R=0.8+1.74Sc; 95 R=1.1+1.74Sc."""
        assert D86_TABLE8["5"]["R"] == pytest.approx((3.1, 1.74))
        assert D86_TABLE8["90"]["R"] == pytest.approx((0.8, 1.74))
        assert D86_TABLE8["95"]["R"] == pytest.approx((1.1, 1.74))
        assert D86_TABLE8["10"]["r"] == pytest.approx((1.2, 0.86))
        assert D86_TABLE8["IBP"]["R"] == pytest.approx((5.6, 0.0))
        assert D86_TABLE8["FBP"]["R"] == pytest.approx((7.2, 0.0))

    def test_d86_table9_coeffs(self):
        """Table 9 гр.2-4: IBP R=2.8+0.93Sc; 5-95 R=1.8+1.33Sc; FBP R=3.1+0.42Sc."""
        assert D86_TABLE9["IBP"]["R"] == pytest.approx((2.8, 0.93))
        assert D86_TABLE9["5_95"]["R"] == pytest.approx((1.8, 1.33))
        assert D86_TABLE9["FBP"]["R"] == pytest.approx((3.1, 0.42))

    def test_d1160_constants(self):
        """Константы 12.5.1/12.5.2: 1 мм r(5-50) a=0.439,b=0.241,M=2.9."""
        r1 = D1160_CONSTANTS[1]["r"]["5_50"]
        assert r1 == pytest.approx((0.439, 0.241, 2.9))
        r10 = D1160_CONSTANTS[10]["r"]["5_50"]
        assert r10 == pytest.approx((0.240, 0.350, 2.8))
        R1 = D1160_CONSTANTS[1]["R"]["5_50"]
        assert R1 == pytest.approx((1.338, 0.639, 3.3))
        R10 = D1160_CONSTANTS[10]["R"]["5_50"]
        assert R10 == pytest.approx((1.415, 0.409, 3.2))
        assert D1160_PRESSURE_RANGE == (1.0, 50.0)

    def test_k_criterion_coeff_0_84(self):
        """Коэффициент 0.84 для ОК — НЕ менять."""
        assert K_CRITERION_COEFF == pytest.approx(0.84)