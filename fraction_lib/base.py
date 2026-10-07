"""Базовый каркас калькуляторов фракционного состава.

Содержит общую для всех методов логику (не зависит от конкретного стандарта):
- нормализация входа (давление, объёмы, температуры),
- поправка давления Сиднея Янга,
- коррекция потерь (ISO/D86 и ГОСТ 2177 — отдельными методами),
- арифметическая интерполяция T(V),
- проверка монотонности,
- материальный баланс (сценарий Б),
- округление до 0,5 °С.

Конкретные калькуляторы (gost2177, iso3405, d86, d1160) наследуют этот класс
и реализуют: `_correct_loss`, `_precision_r`, `_precision_R`, `_slope_points`,
`_compute_extra` и `method_name`.
"""

from __future__ import annotations

import math
from abc import ABC, abstractmethod
from math import floor

from .constants import (
    MMHG_PER_KPA,
    YOUNG_C_MMHG,
    YOUNG_C_KPA,
    YOUNG_STD_MMHG,
    YOUNG_STD_KPA,
)
from .exceptions import (
    MassBalanceViolation,
    NonMonotonicDistillationError,
    PressureOutOfRangeError,
)
from .models import FractionInput, FractionResult, Metrics

# Ключи точек, для которых используется специальный набор формул r/R
IBP_KEY = "IBP"
FBP_KEY = "FBP"


def round_half(value: float) -> float:
    """Округление до 0,5 °С (half away from zero, как MROUND(x, 0.5))."""
    return math.copysign(math.floor(abs(value) * 2.0 + 0.5) / 2.0, value)


def interpolate_t(v: float, v_pct: list[float], t_obs: list[float]) -> float:
    """Арифметическая интерполяция T(V) (формула (6) ISO 3405 / (10) D86).

    v_pct и t_obs — параллельные отсортированные по возрастанию v списки.
    """
    if v <= v_pct[0]:
        return t_obs[0]
    if v >= v_pct[-1]:
        return t_obs[-1]
    for i in range(len(v_pct) - 1):
        v0, v1 = v_pct[i], v_pct[i + 1]
        if v0 <= v <= v1:
            t0, t1 = t_obs[i], t_obs[i + 1]
            if v1 == v0:
                return t0
            return t0 + (t1 - t0) * (v - v0) / (v1 - v0)
    return t_obs[-1]


class BaseDistillationCalculator(ABC):
    """Абстрактный калькулятор разгонки.

    Публичный API: `calculate(input: FractionInput) -> FractionResult`.
    """

    method_name: str = "base"

    # ------------------------------------------------------------------
    # Публичный API
    # ------------------------------------------------------------------
    def calculate(self, inp: FractionInput) -> FractionResult:
        # 1. Нормализация и проверка диапазонов давления
        self._validate_pressure(inp)

        # 2. Разбор температур по точкам (проценты), с ключами IBP/FBP
        points, t1, t2 = self._extract_temperatures(inp)
        # Сохраняем сырые пары (два определения) для проверки повторяемости
        self._last_pairs = {
            p: (t1[i], t2[i]) for i, p in enumerate(points)
            if t1[i] is not None and t2[i] is not None
        }

        # 3. Потери и остаток
        loss_pct = self._resolve_loss(inp)
        residue_pct = self._resolve_residue(inp)
        distillate_pct = self._resolve_distillate(inp)

        # 4. Материальный баланс (сценарий Б)
        auto_norm = inp.auto_normalize
        try:
            loss_pct, residue_pct, distillate_pct, flags_mb = self._mass_balance(
                distillate_pct, residue_pct, loss_pct, auto_normalize=auto_norm
            )
        except MassBalanceViolation:
            raise

        # 5. Коррекция потерь (метод-специфичная)
        loss_corr = self._correct_loss(loss_pct, inp)

        # 6. Поправка давления (Сидней Янг) к средним температурам
        p_mmhg = inp.pressure_mmhg()
        t_corr_raw = {}
        for pt, a, b in zip(points, t1, t2):
            if a is None:
                t_corr_raw[pt] = None
                continue
            t_avg = (a + b) / 2.0 if b is not None else a
            t_corr_raw[pt] = self._apply_pressure_correction(t_avg, p_mmhg)

        # 7. Интерполяция недостающих точек (сценарий А — разреженные данные)
        t_corr_full = self._fill_missing(points, t_corr_raw)

        # 8. Монотонность (аналитическое требование + сценарий А)
        self._check_monotonic(points, t_corr_full)

        # 9. Коррекция объёмов отгона на потери (V_corr)
        v_corr = self._correct_volumes(points, t_corr_full, loss_pct, loss_corr, inp)

        # 10. Крутизна (slope) по точкам
        slopes = self._compute_slopes(points, t_corr_full)

        # 11. r/R/K по точкам (метод-специфично)
        r_vals, R_vals = self._precision_all(points, t_corr_full, slopes, inp)

        # 12. Флаги: предупреждения коррекции потерь (TD-D86-001) + потери > лимита (сценарий В)
        flags = list(flags_mb) + list(getattr(inp, "_d86_loss_flags", []) or [])
        self._check_loss_limit(loss_corr, inp, flags)

        # 13. Метод-специфичные расширения (AET, cracking, overlap...)
        extra = self._compute_extra(points, t_corr_full, slopes, inp, flags)

        metrics = Metrics(
            r={self._to_pct_key(k): v for k, v in r_vals.items()},
            R={self._to_pct_key(k): v for k, v in R_vals.items()},
            slope={self._to_pct_key(k): v for k, v in slopes.items()},
            K={self._to_pct_key(k): self._k_criterion(r_vals.get(k, 0.0), R_vals.get(k, 0.0)) for k in r_vals},
            R_ok={self._to_pct_key(k): R_vals.get(k, 0.0) * 0.84 for k in R_vals},
        )

        return FractionResult(
            T_corr={self._to_pct_key(k): v for k, v in t_corr_full.items()},
            V_corr={self._to_pct_key(k): v for k, v in v_corr.items()},
            Loss=loss_corr,
            Residue=residue_pct,
            Metrics=metrics,
            Flags=list(dict.fromkeys(flags)),
            method=self.method_name,
            raw={"points": [str(p) for p in points], "loss_observed": loss_pct,
                 "loss_corrected": loss_corr, "extra": extra},
        )

    # ------------------------------------------------------------------
    # Общие шаги
    # ------------------------------------------------------------------
    def _validate_pressure(self, inp: FractionInput) -> None:
        """Проверка диапазона давления (переопределяется в подклассах)."""
        pass

    def _extract_temperatures(self, inp: FractionInput):
        """Разбор T_obs на (points, t1, t2).

        points — отсортированный список ключей в порядке % отгона.
        t1 — первое определение, t2 — второе (или None).
        Поддерживает числовые ключи (0..100) и 'IBP'/'FBP'.
        """
        t_obs = inp.T_obs
        keys: list = []
        for k, v in t_obs.items():
            keys.append(k)

        def sort_key(k):
            if isinstance(k, str):
                if k.upper() in (IBP_KEY, "НК", "IBP"):
                    return -1
                if k.upper() in (FBP_KEY, "КК", "FBP"):
                    return 1000
                try:
                    return float(k)
                except ValueError:
                    return 500
            return float(k)

        keys.sort(key=sort_key)
        points, t1, t2 = [], [], []
        for k in keys:
            v = t_obs[k]
            if isinstance(v, (list, tuple)):
                a, b = v[0], v[1] if len(v) > 1 else None
            elif isinstance(v, dict):
                a = v.get("x1", v.get("1"))
                b = v.get("x2", v.get("2"))
            else:
                a, b = v, None
            points.append(k)
            t1.append(float(a) if a is not None else None)
            t2.append(float(b) if b is not None else None)
        return points, t1, t2

    def _resolve_loss(self, inp: FractionInput) -> float:
        if inp.loss_pct is not None:
            return float(inp.loss_pct)
        if inp.loss_ml is not None:
            return float(inp.loss_ml)
        # Потери по умолчанию: 0, если не заданы
        return 0.0

    def _resolve_residue(self, inp: FractionInput) -> float:
        if inp.residue_pct is not None:
            return float(inp.residue_pct)
        if inp.residue_ml is not None:
            return float(inp.residue_ml)
        return 0.0

    def _resolve_distillate(self, inp: FractionInput) -> float:
        """Отгон, %: из V_pct (последняя точка) или V_ml (последняя), или 100-ост-пот."""
        if inp.V_pct:
            return float(inp.V_pct[-1])
        if inp.V_ml:
            return float(inp.V_ml[-1])
        # Если не задан явно — вывести из остатка/потерь
        loss = self._resolve_loss(inp)
        res = self._resolve_residue(inp)
        return 100.0 - loss - res

    def _mass_balance(self, distillate: float, residue: float, loss: float,
                      auto_normalize: bool = False, tolerance: float = 0.2):
        """Сценарий Б: отгон + остаток + потери == 100 ± 0,2 %.

        При auto_normalize=True объёмы нормализуются так, чтобы сумма стала 100,
        и добавляется флаг 'MASS_BALANCE_NORMALIZED'.
        """
        total = distillate + residue + loss
        if abs(total - 100.0) <= tolerance:
            return loss, residue, distillate, []
        if auto_normalize:
            scale = 100.0 / total
            return loss * scale, residue * scale, distillate * scale, ["MASS_BALANCE_NORMALIZED"]
        raise MassBalanceViolation(
            distillate=distillate, residue=residue, loss=loss,
            tolerance=tolerance, auto_normalize=False,
        )

    def _correct_loss(self, loss_pct: float, inp: FractionInput) -> float:
        """Коррекция потерь (переопределяется подклассами)."""
        return loss_pct

    def _apply_pressure_correction(self, t: float, p_mmhg: float) -> float:
        """Поправка Сиднея Янга: C = 0,00012·(760 − P_мм)·(273 + t).

        Эквивалент: C = 0,0009·(101,3 − P_кПа)·(273 + t).
        Значение p_mmhg передаётся уже в мм рт.ст. (внутренний канон).
        """
        return t + YOUNG_C_MMHG * (YOUNG_STD_MMHG - p_mmhg) * (273.0 + t)

    def _fill_missing(self, points, t_corr_raw: dict) -> dict:
        """Интерполяция пропущенных точек (сценарий А — разреженные данные).

        Пропущенными считаются точки с None в t_corr_raw. Для них выполняется
        линейная интерполяция по соседним известным. Если после интерполяции
        ряд не монотонный — исключение NonMonotonicDistillationError.
        """
        known = [(i, p) for i, p in enumerate(points) if t_corr_raw.get(p) is not None]
        result = dict(t_corr_raw)
        if len(known) < 2:
            return result
        for i, p in enumerate(points):
            if result.get(p) is None:
                # найти ближайшие известные слева и справа
                left = [k for k in known if k[0] < i]
                right = [k for k in known if k[0] > i]
                if left and right:
                    i0, p0 = left[-1]
                    i1, p1 = right[0]
                    v0 = float(p0) if not isinstance(p0, str) else -1.0
                    v1 = float(p1) if not isinstance(p1, str) else 1000.0
                    vp = float(p) if not isinstance(p, str) else 500.0
                    if v1 != v0:
                        t0 = result[p0]
                        t1 = result[p1]
                        result[p] = t0 + (t1 - t0) * (vp - v0) / (v1 - v0)
        return result

    def _check_monotonic(self, points, t_corr: dict) -> None:
        """Проверка монотонного роста T по мере роста % отгона.

        Строгое требование аналитического теста: T(V_{i+1}) >= T(V_i).
        Допуск на численную погрешность: −0,01 °С.
        """
        seq = [(p, t_corr[p]) for p in points if t_corr.get(p) is not None]
        for (p_prev, t_prev), (p_next, t_next) in zip(seq, seq[1:]):
            if t_next < t_prev - 0.01:
                raise NonMonotonicDistillationError(
                    point=str(p_next), t_prev=t_prev, t_next=t_next,
                    v_prev=float(p_prev) if not isinstance(p_prev, str) else -1.0,
                    v_next=float(p_next) if not isinstance(p_next, str) else 1000.0,
                )

    def _correct_volumes(self, points, t_corr, loss_pct, loss_corr, inp) -> dict:
        """Коррекция объёмов отгона: Rc = R + (L − Lc) (формула 8 D86 / (4) ISO)."""
        delta = loss_pct - loss_corr
        v_corr = {}
        for p in points:
            if isinstance(p, str):
                continue
            v_corr[p] = float(p) + delta
        return v_corr

    def _compute_slopes(self, points, t_corr) -> dict:
        """Крутизна ΔC/ΔV по точкам (формула (7) ISO 3405 / A4.1 D86).

        Для каждой точки берётся интервал (V_l, V_u) по таблице A4.3/5.
        Упрощённо: slope(p) = (T(p+) − T(p−))/(V(p+) − V(p−)) по соседям.
        Для IBP — интервал 0→5; для FBP — последний интервал.
        Ключи: числовые точки как float; IBP/FBP — строки.
        """
        slopes: dict = {}
        numeric = [(float(p), t_corr[p]) for p in points
                   if not isinstance(p, str) and t_corr.get(p) is not None]
        if not numeric:
            return slopes
        if len(numeric) == 1:
            return {numeric[0][0]: 0.0}

        # IBP: 0→5 (или первый интервал)
        p0, t0 = numeric[0]
        p1, t1 = numeric[1] if len(numeric) > 1 else numeric[0]
        slopes[IBP_KEY] = (t1 - t0) / (p1 - p0) if p1 != p0 else 0.0
        # для самой точки 0 тоже даём slope (первый интервал)
        slopes[float(p0)] = slopes[IBP_KEY]

        # промежуточные: по соседним точкам
        for i in range(1, len(numeric) - 1):
            pa, ta = numeric[i - 1]
            pb, tb = numeric[i + 1]
            slopes[numeric[i][0]] = (tb - ta) / (pb - pa) if pb != pa else 0.0

        # последняя числовая точка (например 95): по последнему интервалу
        pn, tn = numeric[-1]
        pm, tm = numeric[-2]
        slopes[float(pn)] = (tn - tm) / (pn - pm) if pn != pm else 0.0

        # FBP: последний интервал
        slopes[FBP_KEY] = (tn - tm) / (pn - pm) if pn != pm else 0.0
        return slopes

    def _precision_all(self, points, t_corr, slopes, inp) -> tuple[dict, dict]:
        """r/R по точкам (переопределяется подклассами)."""
        r_vals, R_vals = {}, {}
        for p in points:
            r_vals[p] = 0.0
            R_vals[p] = 0.0
        return r_vals, R_vals

    def _check_loss_limit(self, loss_corr: float, inp: FractionInput, flags: list) -> None:
        """Сценарий В: потери > лимита группы → флаг INVALID_TEST."""
        pass

    def _compute_extra(self, points, t_corr, slopes, inp, flags) -> dict:
        """Метод-специфичные расширения (AET, cracking, overlap...)."""
        return {}

    def _k_criterion(self, r: float, R: float) -> float:
        from .models import k_criterion

        return k_criterion(r, R)

    # ------------------------------------------------------------------
    # Хелперы
    # ------------------------------------------------------------------
    @staticmethod
    def _to_pct_key(p) -> str:
        """Нормализация ключа точки: 50.0 → '50'; 5.0 → '5'; 'IBP'/'FBP' как есть."""
        if isinstance(p, str):
            if p.upper() in (IBP_KEY, "НК", "НКК"):
                return IBP_KEY
            if p.upper() in (FBP_KEY, "КК", "ТКК"):
                return FBP_KEY
            return p
        f = float(p)
        if f.is_integer():
            return str(int(f))
        return f"{f:g}"


def slope_between(t_high: float, t_low: float, v_high: float, v_low: float) -> float:
    """Крутизна между двумя точками, °С/%."""
    if v_high == v_low:
        return 0.0
    return (t_high - t_low) / (v_high - v_low)