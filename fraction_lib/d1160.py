"""Калькулятор ASTM D1160-18 «Standard Test Method for Distillation of Petroleum
Products at Reduced Pressure».

Реализует:
- вакуумную разгонку при P_vac = 1–50 мм рт.ст. (0,13–6,7 кПа);
- пересчёт AET по Максвелл-Боннеллу (Annex A7):
    AET = 748,1·A/(1/(T + 273,1) + 0,3861·A − 0,00051606) − 273,1
    A (A7.3, P >= 2 мм): 5,994295 − 0,972546·lg(P)/(2663,129 − 95,76·lg(P))
    A (A7.5, P <  2 мм): 6,761559 − 0,987672·lg(P)/(3000,538 − 43,00·lg(P))
- поправку AET на K-фактор Watson: t = 1,4·(K − 12)·lg(Pa/Po), K = 12,0 ± 0,2;
- точность r/R: формулы (1)/(2) с константами 12.5.1/12.5.2,
  линейная интерполяция констант между 1 и 10 мм рт.ст.;
- IBP/FBP — фиксированные константы.

Сценарии:
  А: термическое разложение → CrackingIndicator (аномальный скачок крутизны
     или T_IBP_vac > decomposition_temp_degC);
  Б: пересечение с D86 → OverlapViolation (разность > overlap_tolerance_degC).
"""

from __future__ import annotations

from math import log10, exp, log

from .base import BaseDistillationCalculator, IBP_KEY, FBP_KEY
from .constants import (
    D1160_CONSTANTS,
    D1160_PRESSURE_RANGE,
    D1160_INTERP_PRESSURES,
    D1160_AET,
    D1160_K_CORR,
    D1160_CRACK_SLOPE_JUMP,
    K_WATSON_NOMINAL,
    EULER_E,
)
from .exceptions import (
    InvalidPressureError,
    CrackingIndicator,
    OverlapViolation,
)
from .models import FractionInput, FractionResult


def maxwell_bonnell_aet(t_obs: float, p_mmhg: float) -> float:
    """Пересчёт наблюдаемой температуры в AET (Annex A7, Maxwell-Bonnell).

    t_obs — наблюдаемая температура пара, °C (при давлении p_mmhg).

    A7.1: AET = 748,1·A / (1/(T+273,1) + 0,3861·A − 0,00051606) − 273,1
    A7.3 (P >= 2 мм): A = (5,994295 − 0,972546·lg(P)) / (2663,129 − 95,76·lg(P))
    A7.5 (P <  2 мм): A = (6,761559 − 0,987672·lg(P)) / (3000,538 − 43,00·lg(P))
    """
    cfg = D1160_AET
    if p_mmhg >= cfg["P_THRESHOLD"]:
        c = cfg["A7_3"]
    else:
        c = cfg["A7_5"]
    lg_p = log10(p_mmhg)
    # ВАЖНО: весь числитель делится на знаменатель (приоритет операций!)
    A = (c["n1"] - c["n2"] * lg_p) / (c["d1"] - c["d2"] * lg_p)
    denom = 1.0 / (t_obs + cfg["T_OFFSET"]) + cfg["K2"] * A - cfg["K3"]
    if abs(denom) < 1e-12:
        return t_obs
    return cfg["K1"] * A / denom - cfg["T_OFFSET"]


def aet_k_correction(aet: float, k_watson: float, p_atm_mmhg: float, p_vac_mmhg: float) -> float:
    """Поправка AET на K-фактор (A7.7): t = 1,4·(K − 12)·lg(Pa/Po)."""
    coef = D1160_K_CORR["coef"]
    nominal = D1160_K_CORR["K_nominal"]
    if p_vac_mmhg <= 0 or p_atm_mmhg <= 0:
        return aet
    ratio = p_atm_mmhg / p_vac_mmhg
    if ratio <= 0:
        return aet
    return aet + coef * (k_watson - nominal) * log10(ratio)


def d1160_precision(slope: float, p_mmhg: float, band: str, kind: str) -> float:
    """Точность D1160: r = M·[exp(a + b·ln(1,8·S))]/1,8.

    band: 'IBP' | '5_50' | '60_95' | 'FBP'.
    kind: 'r' | 'R'.
    Константы интерполируются линейно между 1 и 10 мм рт.ст.
    """
    p_lo, p_hi = D1160_INTERP_PRESSURES
    c_lo = D1160_CONSTANTS[int(p_lo)][kind][band]
    c_hi = D1160_CONSTANTS[int(p_hi)][kind][band]

    if p_mmhg <= p_lo:
        a, b, M = c_lo
    elif p_mmhg >= p_hi:
        a, b, M = c_hi
    else:
        t = (p_mmhg - p_lo) / (p_hi - p_lo)
        a = c_lo[0] + t * (c_hi[0] - c_lo[0])
        b = c_lo[1] + t * (c_hi[1] - c_lo[1])
        M = c_lo[2] + t * (c_hi[2] - c_lo[2])

    if b == 0:
        return M * exp(a) / 1.8
    s_eff = max(slope, 1e-9)
    return M * exp(a + b * log(1.8 * s_eff)) / 1.8


class D1160Calculator(BaseDistillationCalculator):
    """Расчёт вакуумной разгонки по ASTM D1160-18."""

    method_name = "d1160"

    # ------------------------------------------------------------------
    def _validate_pressure(self, inp: FractionInput) -> None:
        p_vac = inp.P_vac
        if p_vac is None:
            raise InvalidPressureError(pressure_mmhg=0.0, method="D1160",
                                       valid_range=D1160_PRESSURE_RANGE)
        lo, hi = D1160_PRESSURE_RANGE
        if p_vac < lo or p_vac > hi:
            raise InvalidPressureError(pressure_mmhg=p_vac, method="D1160",
                                       valid_range=D1160_PRESSURE_RANGE)

    # ------------------------------------------------------------------
    def _correct_loss(self, loss_pct: float, inp: FractionInput) -> float:
        """D1160 не корректирует потери на давление (вакуумная разгонка)."""
        return loss_pct

    # ------------------------------------------------------------------
    def _apply_pressure_correction(self, t: float, p_mmhg: float) -> float:
        """Для D1160 поправка Янга НЕ применяется — вместо неё AET.

        Этот метод не используется: AET считается в _compute_extra.
        Возвращаем t без изменений (сырые значения остаются для AET).
        """
        return t

    # ------------------------------------------------------------------
    def _precision_all(self, points, t_corr, slopes, inp) -> tuple[dict, dict]:
        """r/R по формулам (1)/(2) с константами 12.5.1/12.5.2.

        Точность дана в °C AET. slope — крутизна по AET.
        """
        p_vac = float(inp.P_vac)
        r_vals, R_vals = {}, {}
        for p in points:
            slope = self._slope_for_point(p, slopes, t_corr)
            band = self._band_for_point(p)
            r_vals[p] = d1160_precision(slope, p_vac, band, "r")
            R_vals[p] = d1160_precision(slope, p_vac, band, "R")
        return r_vals, R_vals

    # ------------------------------------------------------------------
    def _band_for_point(self, p) -> str:
        if isinstance(p, str):
            u = p.upper()
            if u in ("IBP", "НК"):
                return "IBP"
            if u in ("FBP", "КК"):
                return "FBP"
        f = float(p)
        if f <= 0.5:
            return "IBP"
        if f >= 99.5:
            return "FBP"
        if 5.0 <= f <= 50.0:
            return "5_50"
        return "60_95"

    def _slope_for_point(self, p, slopes, t_corr) -> float:
        if p in slopes:
            return slopes[p]
        if isinstance(p, str):
            try:
                return slopes.get(float(p), 0.0)
            except (TypeError, ValueError):
                return 0.0
        return slopes.get(float(p), 0.0)

    # ------------------------------------------------------------------
    def _compute_extra(self, points, t_corr, slopes, inp, flags) -> dict:
        """AET, K-поправка, CrackingIndicator, OverlapViolation.

        t_corr для D1160 содержит наблюдаемые (нескорректированные) значения;
        здесь превращаем их в AET и перезаписываем T_corr результата.
        """
        p_atm = inp.pressure_mmhg()
        p_vac = float(inp.P_vac)
        k_watson = K_WATSON_NOMINAL

        # 1. AET для каждой точки
        aet_vals = {}
        for p, t in t_corr.items():
            if t is None:
                continue
            aet = maxwell_bonnell_aet(float(t), p_vac)
            aet = aet_k_correction(aet, k_watson, p_atm, p_vac)
            aet_vals[p] = aet

        # 2. Сценарий А: термическое разложение
        self._check_cracking(t_corr, slopes, inp, flags)

        # 3. Сценарий Б: пересечение с D86 (сравнение AET с D86-данными)
        self._last_aet = aet_vals
        self._check_overlap(aet_vals, inp, flags)

        return {"aet": {self._to_pct_key(k): v for k, v in aet_vals.items()},
                "k_watson": k_watson}

    # ------------------------------------------------------------------
    def _check_cracking(self, t_corr, slopes, inp, flags) -> None:
        """Сценарий А: признаки термического разложения.

        1) T_IBP_vac > decomposition_temp_degC (если параметр задан);
        2) аномальный скачок крутизны > D1160_CRACK_SLOPE_JUMP.
        При срабатывании — CrackingIndicator (флаг; исключение не блокирует
        расчёт, т.к. это индикатор, а не фатальная ошибка).
        """
        reasons = []
        dec = getattr(inp, "decomposition_temp_degC", None)
        if dec is not None:
            ibp = t_corr.get(IBP_KEY)
            if ibp is None:
                for k, v in t_corr.items():
                    if k in (IBP_KEY, "IBP", 0, "0"):
                        ibp = v
                        break
            if ibp is not None and float(ibp) > float(dec):
                reasons.append(f"T_IBP_vac={ibp:.1f}°C > T_разложения={dec:.1f}°C")

        # скачок крутизны
        num_slopes = [(p, s) for p, s in slopes.items() if s is not None]
        for (p1, s1), (p2, s2) in zip(num_slopes, num_slopes[1:]):
            if s2 - s1 > D1160_CRACK_SLOPE_JUMP:
                reasons.append(f"скачок крутизны {s1:.1f}→{s2:.1f} °С/% (порог {D1160_CRACK_SLOPE_JUMP})")
                break

        if reasons:
            flags.append("CrackingIndicator: " + "; ".join(reasons))

    # ------------------------------------------------------------------
    def _check_overlap(self, aet_vals, inp, flags) -> None:
        """Сценарий Б: пересечение D86/D1160 в зоне перекрытия.

        Если в inp задан d86_overlap (dict {pct: t_d86}), сравниваем AET
        с температурами D86. Разность > допуска → OverlapViolation (флаг).
        """
        tol = getattr(inp, "overlap_tolerance_degC", 5.0)
        d86_data = getattr(inp, "d86_overlap", None)
        if not d86_data:
            return
        for pct, t_d86 in d86_data.items():
            key = self._to_pct_key(pct)
            # aet_vals может иметь int/str ключи — нормализуем оба
            aet_key = key
            if aet_key not in aet_vals:
                # ищем по числовому эквиваленту
                for k in aet_vals:
                    if self._to_pct_key(k) == key:
                        aet_key = k
                        break
            if aet_key in aet_vals:
                t_d1160 = aet_vals[aet_key]
                if abs(t_d1160 - float(t_d86)) > tol:
                    flags.append(f"OverlapViolation:{key}% Δ={abs(t_d1160 - float(t_d86)):.2f}°C>tol={tol}°C")

    # ------------------------------------------------------------------
    def calculate(self, inp: FractionInput) -> FractionResult:
        """Переопределение: AET перезаписывает T_corr после базового расчёта."""
        result = super().calculate(inp)
        # t_corr в базовом расчёте для D1160 — это наблюдаемые температуры
        # (поправка Янга отключена). Пересчитываем в AET и кладём в результат.
        p_vac = float(inp.P_vac)
        p_atm = inp.pressure_mmhg()
        aet_vals = {}
        for k, v in result.T_corr.items():
            if v is None:
                continue
            aet = maxwell_bonnell_aet(float(v), p_vac)
            aet = aet_k_correction(aet, K_WATSON_NOMINAL, p_atm, p_vac)
            aet_vals[k] = round(aet, 3)
        result.T_corr = aet_vals
        result.raw["aet_final"] = dict(aet_vals)
        self._last_aet = aet_vals
        return result


__all__ = ["D1160Calculator", "maxwell_bonnell_aet", "aet_k_correction", "d1160_precision"]