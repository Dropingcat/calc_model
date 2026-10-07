"""Калькулятор ASTM D86-20a «Standard Test Method for Distillation of Petroleum
Products and Liquid Fuels at Atmospheric Pressure».

Реализует ручной метод:
- группы 1–4 (object_type: '1'..'4');
- поправка давления Сиднея Янга (11.3, формулы (3)/(4));
- коррекция потерь (11.4, формулы (6)/(7));
- скорректированный процент отгона (11.4.1, формула (8));
- наклон кривой (Annex A4, A4.1);
- точность: Table 8 (гр.1), Table 9 (гр.2-4);
- лимиты потерь (Table 3): гр. 1-2 → 1,5 %; гр. 3-4 → 2,0 %.

Сценарии:
  А: интерполяция разреженных данных + проверка монотонности;
  Б: границы групп (inclusive_ge);
  В: потери > лимита → INVALID_TEST.
"""

from __future__ import annotations

from .base import BaseDistillationCalculator, IBP_KEY, FBP_KEY
from .constants import (
    D86_TABLE8,
    D86_TABLE9,
    D86_LOSS_LIMIT,
    LOSS_CORR_KPA,
    LOSS_CORR_KPA_STD,
    LOSS_CORR_MMHG_DIV,
    LOSS_CORR_MMHG_STD,
)
from .exceptions import LossLimitExceeded, BoundaryRuleError
from .models import FractionInput, FractionResult


class D86Calculator(BaseDistillationCalculator):
    """Расчёт фракционного состава по ASTM D86-20a (ручной метод)."""

    method_name = "d86"

    # ------------------------------------------------------------------
    def _group(self, inp: FractionInput) -> int:
        try:
            return int(float(str(inp.object_type).strip()))
        except (TypeError, ValueError):
            raise ValueError(f"object_type должен быть группой 1-4, получено: {inp.object_type!r}")

    # ------------------------------------------------------------------
    # Диапазон сингулярности/экстремума формулы 11.4 (мм рт.ст.):
    # denom = 1 + (760-P)/60 → 0 при P ≈ 820 мм (≈109,3 кПа).
    LOSS_SINGULAR_MMHG_LO = 810.0
    LOSS_SINGULAR_MMHG_HI = 830.0

    def _correct_loss(self, loss_pct: float, inp: FractionInput) -> float:
        """Коррекция потерь D86 (11.4, ASTM D86-23):

        кПа: Lc = 0,5 + (L − 0,5)/(1 + (101,3 − P)/8,00)
        мм:  Lc = 0,5 + (L − 0,5)/(1 + (760 − P)/60,0)

        TD-D86-001 (синхронизация с книгой v3.12): единая формула
        применяется для ВСЕГО диапазона давлений, включая P > 760 мм.
        Ранее при P > 760 коррекция молча отключалась — это расходилось
        с D86-23 и с Excel-книгой ASTM_D86_ручной_метод_v3.12.

        Защита от деления на ноль: вблизи P ≈ 820 мм (109,3 кПа)
        знаменатель → 0; там возвращаем наблюдаемые потери без
        изменения и добавляем предупреждение LOSS_CORR_UNSTABLE
        (флаг через inp._d86_loss_flags, подхватывается calculate()).
        """
        if inp.pressure_unit.lower() in ("kpa", "кпа"):
            p = inp.pressure_kpa()
            denom = 1.0 + (LOSS_CORR_KPA_STD - p) / 8.00
        else:
            p = inp.pressure_mmhg()
            denom = 1.0 + (LOSS_CORR_MMHG_STD - p) / LOSS_CORR_MMHG_DIV

        if abs(denom) < 1e-9 or self.LOSS_SINGULAR_MMHG_LO <= p <= self.LOSS_SINGULAR_MMHG_HI:
            self._flag_loss(inp, f"LOSS_CORR_UNSTABLE:P={p:.1f}mm(знаменатель~0),потери без коррекции")
            return loss_pct
        lc = 0.5 + (loss_pct - 0.5) / denom
        if lc < 0:
            # физически абсурдный результат (P сильно выше 820 мм)
            self._flag_loss(inp, f"LOSS_CORR_NEGATIVE:Lc={lc:.2f}<0,P={p:.1f}mm")
            return loss_pct
        return lc

    @staticmethod
    def _flag_loss(inp: FractionInput, msg: str) -> None:
        """Собрать предупреждение о коррекции потерь (подхватится base.calculate)."""
        flags = getattr(inp, "_d86_loss_flags", None)
        if flags is None:
            flags = []
            setattr(inp, "_d86_loss_flags", flags)
        if msg not in flags:
            flags.append(msg)

    # ------------------------------------------------------------------
    def _precision_all(self, points, t_corr, slopes, inp) -> tuple[dict, dict]:
        """r/R по Table 8 (гр.1) и Table 9 (гр.2-4)."""
        group = self._group(inp)
        r_vals, R_vals = {}, {}
        boundary_rule = getattr(inp, "boundary_rule", "inclusive_ge")

        if group == 1:
            for p in points:
                slope = self._slope_for_point(p, slopes, t_corr)
                if isinstance(p, str) and p in (IBP_KEY, FBP_KEY):
                    key = p
                elif self._is_ibp(p):
                    key = "IBP"
                elif self._is_fbp(p):
                    key = "FBP"
                else:
                    key = self._pct_key(p)
                entry = D86_TABLE8.get(key)
                if entry is None:
                    # промежуточные точки → как 10-80
                    entry = D86_TABLE8["10"]
                a_r, b_r = entry["r"]
                a_R, b_R = entry["R"]
                r_vals[p] = a_r + b_r * slope
                R_vals[p] = a_R + b_R * slope
        else:
            for p in points:
                slope = self._slope_for_point(p, slopes, t_corr)
                if isinstance(p, str) and p in (IBP_KEY, FBP_KEY):
                    coeffs = D86_TABLE9[p]
                elif self._is_ibp(p):
                    coeffs = D86_TABLE9["IBP"]
                elif self._is_fbp(p):
                    coeffs = D86_TABLE9["FBP"]
                else:
                    coeffs = D86_TABLE9["5_95"]
                a_r, b_r = coeffs["r"]
                a_R, b_R = coeffs["R"]
                r_vals[p] = a_r + b_r * slope
                R_vals[p] = a_R + b_R * slope

        self._check_boundary(points, slopes, boundary_rule)
        return r_vals, R_vals

    # ------------------------------------------------------------------
    def _slope_for_point(self, p, slopes, t_corr) -> float:
        if p in slopes:
            return slopes[p]
        if isinstance(p, str):
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
        """Сценарий Б: границы интервалов r/R (inclusive_ge)."""
        if boundary_rule != "inclusive_ge":
            for p, s in slopes.items():
                if s is None:
                    continue
                for boundary in (1.0, 2.0, 3.0, 5.0, 10.0):
                    if abs(s - boundary) < 1e-9:
                        raise BoundaryRuleError(slope=s, boundary=boundary, rule=boundary_rule)

    # ------------------------------------------------------------------
    def _check_loss_limit(self, loss_corr: float, inp: FractionInput, flags: list) -> None:
        """Сценарий В: потери > лимита группы → INVALID_TEST."""
        group = self._group(inp)
        limit = D86_LOSS_LIMIT.get(group)
        if limit is None:
            return
        if loss_corr > limit:
            flags.append("INVALID_TEST")
            flags.append(f"LOSS_LIMIT_EXCEEDED:{loss_corr:.2f}>{limit:.2f}")

    # ------------------------------------------------------------------
    def _compute_extra(self, points, t_corr, slopes, inp, flags) -> dict:
        return {"group": self._group(inp)}


__all__ = ["D86Calculator"]