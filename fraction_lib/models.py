"""Модели данных библиотеки fraction_lib.

Единый интерфейс входа/выхода по контракту ФР-23+:
- Вход:  FractionInput (P_atm, P_vac, object_type, V_ml, V_pct, T_obs, ...)
- Выход: FractionResult (T_corr, V_corr, Loss, Residue, Metrics, Flags)

Все dataclass строго типизированы; допускаются только стандартные типы
(dict/list/float/str) — совместимо с JSON-протоколами.
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Any


# ---------------------------------------------------------------------------
# Вход
# ---------------------------------------------------------------------------
@dataclass
class FractionInput:
    """Единая структура входа для всех методов (ГОСТ 2177, ISO 3405, D86, D1160).

    Атрибуты:
        P_atm: атмосферное давление (число); единица — параметр pressure_unit.
        pressure_unit: 'mmHg' (по умолчанию) или 'kPa'.
        P_vac: давление разгонки для D1160 в мм рт.ст. (1–50); иначе None.
        object_type: группа 1–4 (строка '1'..'4') для D86/ISO 3405;
                     либо тип продукта для ГОСТ 2177 метод А/Б ('A'/'B').
        method: 'gost2177' | 'iso3405' | 'd86' | 'd1160' (определяет калькулятор).
        V_ml: объём отгона в мл (сырые данные колбы) — опционально.
        V_pct: объём отгона в % (0, 5, 10, 20, …, 100) — опционально.
        T_obs: наблюдаемые температуры; ключ = % отгона (0 → НК, 100 → КК)
               или строковые 'IBP'/'FBP'.
        loss_ml: потери, мл (опционально, для материального баланса).
        residue_ml: остаток, мл (опционально).
        loss_pct: потери, % (если заданы напрямую).
        residue_pct: остаток, % (если заданы напрямую).
        auto_normalize: True — автонормализация баланса вместо блокировки.
        boundary_rule: 'inclusive_ge' (по умолчанию) — правило границ групп.
        overlap_tolerance_degC: допуск пересечения D86/D1160 (по умолчанию 5.0).
        decomposition_temp_degC: температура разложения продукта для D1160
                                 (параметр сценария А), °С.
    """

    P_atm: float
    pressure_unit: str = "mmHg"
    P_vac: float | None = None
    object_type: str = "1"
    method: str = "iso3405"
    V_ml: list[float] = field(default_factory=list)
    V_pct: list[float] = field(default_factory=list)
    T_obs: dict = field(default_factory=dict)
    loss_ml: float | None = None
    residue_ml: float | None = None
    loss_pct: float | None = None
    residue_pct: float | None = None
    auto_normalize: bool = False
    boundary_rule: str = "inclusive_ge"
    overlap_tolerance_degC: float = 5.0
    decomposition_temp_degC: float | None = None

    def pressure_mmhg(self) -> float:
        """Давление в мм рт.ст. (единый внутренний канон)."""
        if self.pressure_unit.lower() in ("kpa", "кпа"):
            return self.P_atm / 0.133322
        return self.P_atm

    def pressure_kpa(self) -> float:
        """Давление в кПа."""
        if self.pressure_unit.lower() in ("kpa", "кпа"):
            return self.P_atm
        return self.P_atm * 0.133322

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


# ---------------------------------------------------------------------------
# Выход
# ---------------------------------------------------------------------------
@dataclass
class Metrics:
    """Метрики результата: r, R, крутизна, критерий K.

    K — критерий по свидетельству ГСО, формулу НЕ менять:
        K = (1/√2) · √(R² − r²/2)
    Коэффициент 0.84 — из сборника «Метрология в химическом анализе», НЕ менять.
    """

    r: dict[str, float] = field(default_factory=dict)   # повторяемость по точкам
    R: dict[str, float] = field(default_factory=dict)   # воспроизводимость по точкам
    slope: dict[str, float] = field(default_factory=dict)  # крутизна °С/% по точкам
    K: dict[str, float] = field(default_factory=dict)   # критерий ГСО по точкам
    R_ok: dict[str, float] = field(default_factory=dict)  # R*0,84 (для ОК)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class FractionResult:
    """Единая структура выхода.

    Атрибуты:
        T_corr: скорректированные температуры, ключ = % отгона или 'IBP'/'FBP'.
        V_corr: скорректированные объёмы отгона, % (с коррекцией потерь).
        Loss: скорректированные потери, %.
        Residue: остаток, %.
        Metrics: метрики (r, R, slope, K, R_ok).
        Flags: предупреждения/флаги (CrackingIndicator, INVALID_TEST и т.п.).
        method: применённый метод.
        raw: сырые промежуточные значения (для отладки/протокола).
    """

    T_corr: dict = field(default_factory=dict)
    V_corr: dict = field(default_factory=dict)
    Loss: float = 0.0
    Residue: float = 0.0
    Metrics: Metrics = field(default_factory=Metrics)
    Flags: list[str] = field(default_factory=list)
    method: str = ""
    raw: dict = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["Metrics"] = self.Metrics.to_dict()
        return d


def k_criterion(r: float, R: float) -> float:
    """Критерий K по свидетельству ГСО. ФОРМУЛУ НЕ МЕНЯТЬ.

    K = (1/√2) · √(R² − r²/2)

    Формула сохраняется дословно; при R² − r²/2 < 0 (нефизичный случай,
    когда r > R·√2) критерий не определён — возвращается 0.0 (нет нарушения),
    а не math domain error.
    """
    from math import sqrt

    radicand = R * R - r * r / 2.0
    if radicand < 0.0:
        return 0.0
    return (1.0 / sqrt(2.0)) * sqrt(radicand)