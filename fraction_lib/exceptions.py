"""Исключения библиотеки fraction_lib.

Каждое исключение соответствует конкретному сценарию из контракта
(«Глубокое ветвление сценариев») и несёт машиночитаемый код + человекочитаемое
сообщение для протокола испытания.
"""

from __future__ import annotations


class FractionLibError(Exception):
    """Базовый класс всех ошибок fraction_lib."""

    code: str = "FRACTION_LIB_ERROR"

    def __init__(self, message: str, *, code: str | None = None, details: dict | None = None) -> None:
        self.message = message
        self.code = code or self.__class__.code
        self.details = details or {}
        super().__init__(self.message)

    def to_dict(self) -> dict:
        """Компактное представление для протокола (JSON-friendly)."""
        return {"code": self.code, "message": self.message, "details": self.details}


class PressureOutOfRangeError(FractionLibError):
    """Сценарий А ГОСТ 2177: давление вне [560, 760] мм рт.ст.

    Экстраполяция коэффициентов A, B таблицы 4 ГОСТ 2177-99 физически
    некорректна — библиотека отказывается считать, а не экстраполирует.
    """

    code = "PRESSURE_OUT_OF_RANGE"

    def __init__(self, pressure_mmhg: float, *, min_mmhg: float = 560.0, max_mmhg: float = 760.0) -> None:
        self.pressure_mmhg = pressure_mmhg
        self.min_mmhg = min_mmhg
        self.max_mmhg = max_mmhg
        super().__init__(
            f"Атмосферное давление {pressure_mmhg} мм рт.ст. вне допустимого "
            f"диапазона [{min_mmhg}, {max_mmhg}] для ГОСТ 2177-99 (таблица 4). "
            "Экстраполяция коэффициентов коррекции потерь не допускается.",
            details={
                "pressure_mmhg": pressure_mmhg,
                "min_mmhg": min_mmhg,
                "max_mmhg": max_mmhg,
            },
        )


class MassBalanceViolation(FractionLibError):
    """Сценарий Б ГОСТ 2177 / ISO 3405 / D86: материальный баланс нарушен.

    Отгон + Остаток + Потери != 100 % (допуск ±0,2 %).
    По умолчанию вывод блокируется; при auto_normalize=True выполняется
    автонормализация объёмов и в результат добавляется флаг.
    """

    code = "MASS_BALANCE_VIOLATION"

    def __init__(
        self,
        *,
        distillate: float,
        residue: float,
        loss: float,
        tolerance: float = 0.2,
        auto_normalize: bool = False,
    ) -> None:
        self.distillate = distillate
        self.residue = residue
        self.loss = loss
        self.tolerance = tolerance
        self.auto_normalize = auto_normalize
        total = distillate + residue + loss
        super().__init__(
            f"Материальный баланс нарушен: отгон {distillate}% + остаток {residue}% "
            f"+ потери {loss}% = {total}% (требуется 100 ± {tolerance}%). "
            + ("Выполнена автонормализация." if auto_normalize else "Требуется повторный эксперимент."),
            details={
                "distillate": distillate,
                "residue": residue,
                "loss": loss,
                "total": total,
                "tolerance": tolerance,
                "auto_normalize": auto_normalize,
            },
        )


class NonMonotonicDistillationError(FractionLibError):
    """Сценарий А ISO 3405 / D86: интерполяция дала немонотонный ряд температур.

    T(V_{i+1}) >= T(V_i) обязано выполняться для корректной кривой разгонки.
    """

    code = "NON_MONOTONIC_DISTILLATION"

    def __init__(self, *, point: str, t_prev: float, t_next: float, v_prev: float, v_next: float) -> None:
        self.point = point
        super().__init__(
            f"Немонотонная кривая разгонки в точке {point}: T({v_prev}%)={t_prev} °С, "
            f"T({v_next}%)={t_next} °С. После интерполяции температура не растёт с ростом отгона.",
            details={"point": point, "t_prev": t_prev, "t_next": t_next, "v_prev": v_prev, "v_next": v_next},
        )


class RepeatabilityViolation(FractionLibError):
    """Сценарий В ГОСТ 2177: |X1-X2| > r для точки.

    Расчёт средних блокируется — требуется повторный эксперимент.
    """

    code = "REPEATABILITY_VIOLATION"

    def __init__(self, *, point: str, x1: float, x2: float, r: float) -> None:
        self.point = point
        super().__init__(
            f"Нарушение повторяемости в точке {point}: |{x1} - {x2}| = {abs(x1 - x2):.2f} > r = {r:.2f}. "
            "Сходимость не выполняется — требуется повторный эксперимент.",
            details={"point": point, "x1": x1, "x2": x2, "diff": abs(x1 - x2), "r": r},
        )


class LossLimitExceeded(FractionLibError):
    """Сценарий В ISO 3405 / D86: потери превышают лимит группы.

    Протокол испытания помечается INVALID_TEST (флаг) — результат не
    используется для приёмки.
    """

    code = "LOSS_LIMIT_EXCEEDED"

    def __init__(self, *, loss: float, limit: float, group: int | None = None) -> None:
        self.loss = loss
        self.limit = limit
        super().__init__(
            f"Потери {loss:.2f}% превышают лимит {limit:.2f}%"
            + (f" для группы {group}" if group is not None else "")
            + ". Испытание признаётся недействительным (INVALID_TEST).",
            details={"loss": loss, "limit": limit, "group": group},
        )


class InvalidTestError(FractionLibError):
    """Испытание недействительно (общий случай протокола INVALID_TEST)."""

    code = "INVALID_TEST"

    def __init__(self, message: str, *, reason: str | None = None) -> None:
        super().__init__(message, details={"reason": reason or message})


class BoundaryRuleError(FractionLibError):
    """Сценарий Б ISO 3405 / D86: крутизна на границе интервалов r/R.

    Правило по умолчанию — «≥ включительно» (inclusive_ge). При нарушении
    параметра boundary_rule выбрасывается данное исключение.
    """

    code = "BOUNDARY_RULE_VIOLATION"

    def __init__(self, *, slope: float, boundary: float, rule: str) -> None:
        self.slope = slope
        self.boundary = boundary
        super().__init__(
            f"Крутизна {slope:.3f} °С/% на границе интервала {boundary:.3f} °С/% "
            f"не соответствует правилу '{rule}'.",
            details={"slope": slope, "boundary": boundary, "rule": rule},
        )


class OverlapViolation(FractionLibError):
    """Сценарий Б D1160: расхождение D86/D1160 в зоне перекрытия.

    Если продукт тестирован и по D86, и по D1160, разность температур в зоне
    перекрытия должна быть <= overlap_tolerance_degC (по умолчанию 5,0 °С).
    """

    code = "OVERLAP_VIOLATION"

    def __init__(self, *, pct: float, t_d86: float, t_d1160: float, tolerance: float = 5.0) -> None:
        self.pct = pct
        super().__init__(
            f"Расхождение D86/D1160 в зоне перекрытия при {pct}%: "
            f"|{t_d86:.2f} - {t_d1160:.2f}| = {abs(t_d86 - t_d1160):.2f} °С > допуск {tolerance} °С. "
            "Результаты методов несовместимы.",
            details={
                "pct": pct,
                "t_d86": t_d86,
                "t_d1160": t_d1160,
                "diff": abs(t_d86 - t_d1160),
                "tolerance": tolerance,
            },
        )


class InvalidPressureError(FractionLibError):
    """Некорректное давление разгонки D1160 (вне [1, 50] мм рт.ст.)."""

    code = "INVALID_PRESSURE"

    def __init__(self, *, pressure_mmhg: float, method: str, valid_range: tuple[float, float]) -> None:
        super().__init__(
            f"Давление {pressure_mmhg} мм рт.ст. вне допустимого диапазона {valid_range} "
            f"для метода {method}.",
            details={"pressure_mmhg": pressure_mmhg, "method": method, "valid_range": list(valid_range)},
        )


class CrackingIndicator(FractionLibError):
    """Сценарий А D1160: признак термического разложения.

    Выбрасывается как исключение (подкласс FractionLibError), но несёт
    severity='flag' — может быть перехвачен как предупреждение и добавлен
    в Flags результата.
    """

    code = "CRACKING_INDICATOR"

    def __init__(self, message: str, *, slope_jump: float | None = None, threshold: float | None = None) -> None:
        super().__init__(
            message,
            details={"slope_jump": slope_jump, "threshold": threshold, "severity": "flag"},
        )