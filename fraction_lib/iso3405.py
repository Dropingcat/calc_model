"""Калькулятор ГОСТ ISO 3405-2022 (ISO 3405:2019) «Нефтепродукты. Определение
фракционного состава при атмосферном давлении».

Реализует ручной метод:
- группы 1–4 (object_type: '1'..'4');
- поправка давления Сиднея Янга (11.2, формула (2));
- коррекция потерь (11.3, формула (3)): Lc = 0,5 + (L − 0,5)/(1 + 0,125·(101,3 − p));
- скорректированный процент отгона (11.4, формула (4)): Rc = R + (L − Lc);
- наклон кривой (13.1, формула (7));
- повторяемость/воспроизводимость: таблица 6 (группа 1), таблица 7 (группы 2–4);
- лимиты потерь: гр. 1-2 → 1,5 %; гр. 3-4 → 2,0 % (сценарий В → INVALID_TEST).

Сценарии глубокого ветвления:
  А (разреженные данные): интерполяция T(V) + проверка монотонности;
  Б (границы групп): правило inclusive_ge при выборе r/R;
  В (потери > лимита): LossLimitExceeded → флаг INVALID_TEST.
"""

from __future__ import annotations

from .base import BaseDistillationCalculator, IBP_KEY, FBP_KEY, round_half
from .constants import (
    ISO3405_G1_SLOPE_COEF,
    ISO3405_G1_R_CORR,
    ISO3405_G1_FIXED,
    ISO3405_G234,
    ISO3405_LOSS_LIMIT,
    LOSS_CORR_KPA,
    LOSS_CORR_KPA_STD,
    MMHG_PER_KPA,
    YOUNG_C_KPA,
    YOUNG_STD_KPA,
)
from .exceptions import (
    LossLimitExceeded,
    BoundaryRuleError,
)
from .models import FractionInput, FractionResult


class ISO3405Calculator(BaseDistillationCalculator):
    """Расчёт фракционного состава по ГОСТ ISO 3405-2022 (ручной метод)."""

    method_name = "iso3405"

    # ------------------------------------------------------------------
    def _group(self, inp: FractionInput) -> int:
        try:
            return int(float(str(inp.object_type).strip()))
        except (TypeError, ValueError):
            raise ValueError(f"object_type должен быть группой 1-4, получено: {inp.object_type!r}")

    # ------------------------------------------------------------------
    def _correct_loss(self, loss_pct: float, inp: FractionInput) -> float:
        """Lc = 0,5 + (L − 0,5)/(1 + 0,125·(101,3 − p_кПа))  (11.3, формула (3)).

        Область применимости: формула выведена для давления НИЖЕ стандартного
        (101,3 кПа / 760 мм рт.ст.). При P > 760 мм:
          - знаменатель убывает и обращается в 0 при ~820 мм (полюс);
          - коррекция даёт ОТРИЦАТЕЛЬНЫЕ потери (физически абсурдно).
        Поэтому при P > 760 мм потери НЕ корректируются (возврат loss_pct).
        """
        p_kpa = inp.pressure_kpa()
        if p_kpa > 101.3:
            return loss_pct
        denom = 1.0 + LOSS_CORR_KPA * (LOSS_CORR_KPA_STD - p_kpa)
        if abs(denom) < 1e-9:
            return loss_pct
        return 0.5 + (loss_pct - 0.5) / denom

    # ------------------------------------------------------------------
    def _apply_pressure_correction(self, t: float, p_mmhg: float) -> float:
        """Поправка Сиднея Янга (11.2, формула (2)) — в кПа:
        Тс = 0,0009·(101,3 − p_кПа)·(273 + t).
        Вход p_mmhg конвертируется в кПа (0,133322 кПа/мм).
        """
        p_kpa = p_mmhg * MMHG_PER_KPA
        return t + YOUNG_C_KPA * (YOUNG_STD_KPA - p_kpa) * (273.0 + t)

    # ------------------------------------------------------------------
    def _precision_all(self, points, t_corr, slopes, inp) -> tuple[dict, dict]:
        """r/R по таблицам 6 (гр.1) и 7 (гр.2-4) ручного метода.

        slope — крутизна °С/% по точкам (ключи: числа, 'IBP', 'FBP').
        """
        group = self._group(inp)
        r_vals, R_vals = {}, {}
        boundary_rule = getattr(inp, "boundary_rule", "inclusive_ge")

        if group == 1:
            a_r, b_r = ISO3405_G1_SLOPE_COEF["r"]
            a_R, b_R = ISO3405_G1_SLOPE_COEF["R"]
            for p in points:
                slope = self._slope_for_point(p, slopes, t_corr)
                r = a_r * slope + b_r
                R = a_R * slope + b_R
                key = self._pct_key(p)
                corr = ISO3405_G1_R_CORR.get(key, 0.0)
                R += corr
                r_vals[p] = r
                R_vals[p] = R
            # IBP/FBP фиксированные
            if IBP_KEY in points or "IBP" in [str(x).upper() for x in points]:
                r_vals[IBP_KEY] = ISO3405_G1_FIXED["IBP"][0]
                R_vals[IBP_KEY] = ISO3405_G1_FIXED["IBP"][1]
            if FBP_KEY in points or "FBP" in [str(x).upper() for x in points]:
                r_vals[FBP_KEY] = ISO3405_G1_FIXED["FBP"][0]
                R_vals[FBP_KEY] = ISO3405_G1_FIXED["FBP"][1]
        else:
            # группы 2-4: таблица 7
            for p in points:
                slope = self._slope_for_point(p, slopes, t_corr)
                if isinstance(p, str) and p in (IBP_KEY, FBP_KEY):
                    coeffs = ISO3405_G234[p]
                elif self._is_ibp(p):
                    coeffs = ISO3405_G234["IBP"]
                elif self._is_fbp(p):
                    coeffs = ISO3405_G234["FBP"]
                else:
                    coeffs = ISO3405_G234["5_95"]
                a_r, b_r = coeffs["r"]
                a_R, b_R = coeffs["R"]
                r_vals[p] = a_r * slope + b_r
                R_vals[p] = a_R * slope + b_R

        # Сценарий Б: граница групп — проверка правила inclusive_ge
        self._check_boundary(points, slopes, boundary_rule)
        return r_vals, R_vals

    # ------------------------------------------------------------------
    def _slope_for_point(self, p, slopes, t_corr) -> float:
        if p in slopes:
            return slopes[p]
        if isinstance(p, str):
            if p in slopes:
                return slopes[p]
            # 'IBP'/'FBP' уже в slopes; для строк-чисел берём числовое значение
            try:
                return slopes.get(float(p), 0.0)
            except (TypeError, ValueError):
                return 0.0
        return slopes.get(float(p), 0.0)

    def _pct_key(self, p) -> str:
        if isinstance(p, str):
            return p
        return str(int(p)) if float(p).is_integer() else f"{p:g}"

    def _is_ibp(self, p) -> bool:
        if isinstance(p, str):
            return p.upper() in ("IBP", "НК", "НКК")
        return float(p) <= 0.5

    def _is_fbp(self, p) -> bool:
        if isinstance(p, str):
            return p.upper() in ("FBP", "КК", "ТКК")
        return float(p) >= 99.5

    # ------------------------------------------------------------------
    def _check_boundary(self, points, slopes, boundary_rule) -> None:
        """Сценарий Б: крутизна ровно на границе интервалов r/R.

        Правило по умолчанию inclusive_ge — граница включается в верхний
        интервал. Если параметр boundary_rule нарушен — BoundaryRuleError.
        """
        if boundary_rule != "inclusive_ge":
            # Для полноты: если задано другое правило, проверяем на границах
            # ключевых интервалов (5 %, 10 % и т.п.) — триггер BoundaryRuleError.
            for p, s in slopes.items():
                if s is None:
                    continue
                for boundary in (1.0, 2.0, 3.0, 5.0, 10.0):
                    if abs(s - boundary) < 1e-9:
                        raise BoundaryRuleError(slope=s, boundary=boundary, rule=boundary_rule)

    # ------------------------------------------------------------------
    def _check_loss_limit(self, loss_corr: float, inp: FractionInput, flags: list) -> None:
        """Сценарий В: потери > лимита группы → LossLimitExceeded + INVALID_TEST."""
        group = self._group(inp)
        limit = ISO3405_LOSS_LIMIT.get(group)
        if limit is None:
            return
        if loss_corr > limit:
            flags.append("INVALID_TEST")
            flags.append(f"LOSS_LIMIT_EXCEEDED:{loss_corr:.2f}>{limit:.2f}")

    # ------------------------------------------------------------------
    def _compute_extra(self, points, t_corr, slopes, inp, flags) -> dict:
        return {"group": self._group(inp)}


# ---------------------------------------------------------------------------
# Экспорт
# ---------------------------------------------------------------------------
__all__ = ["ISO3405Calculator"]