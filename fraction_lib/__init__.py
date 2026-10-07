"""fraction_lib — Python-библиотека расчёта фракционного состава.

Источник истины для расчётов по:
- ГОСТ 2177-99 (методы А/Б),
- ГОСТ ISO 3405-2022 (ISO 3405:2019),
- ASTM D86-20a,
- ASTM D1160-18 (вакуумная разгонка, AET).

Единый интерфейс: FractionInput → calculate() → FractionResult.
Тройная валидация: аналитическая, нормативная, чистый лист (эталоны xlsx).
"""

from __future__ import annotations

from .models import (
    FractionInput,
    FractionResult,
    Metrics,
    k_criterion,
)
from .exceptions import (
    FractionLibError,
    PressureOutOfRangeError,
    MassBalanceViolation,
    NonMonotonicDistillationError,
    RepeatabilityViolation,
    ReproducibilityViolation,
    LossLimitExceeded,
    InvalidTestError,
    BoundaryRuleError,
    OverlapViolation,
    InvalidPressureError,
    CrackingIndicator,
)
from .base import BaseDistillationCalculator, round_half, interpolate_t
from .gost2177 import GOST2177Calculator, GOST2177ManualCalculator
from .iso3405 import ISO3405Calculator
from .d86 import D86Calculator
from .d1160 import D1160Calculator, maxwell_bonnell_aet, aet_k_correction, d1160_precision

__all__ = [
    "FractionInput",
    "FractionResult",
    "Metrics",
    "k_criterion",
    "FractionLibError",
    "PressureOutOfRangeError",
    "MassBalanceViolation",
    "NonMonotonicDistillationError",
    "RepeatabilityViolation",
    "ReproducibilityViolation",
    "LossLimitExceeded",
    "InvalidTestError",
    "BoundaryRuleError",
    "OverlapViolation",
    "InvalidPressureError",
    "CrackingIndicator",
    "BaseDistillationCalculator",
    "round_half",
    "interpolate_t",
    "GOST2177Calculator",
    "GOST2177ManualCalculator",
    "ISO3405Calculator",
    "D86Calculator",
    "D1160Calculator",
    "maxwell_bonnell_aet",
    "aet_k_correction",
    "d1160_precision",
]