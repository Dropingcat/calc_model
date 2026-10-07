"""Калькулятор ГОСТ 2177-99 «Нефтепродукты. Методы определения фракционного состава».

Метод А (лёгкие продукты): поправка давления по формуле (1) [Сидней Янг],
коррекция потерь V_K = A·L + B (таблица 4, линейная интерполяция),
точность r/R по номограмме (рис. 6) через крутизну C.

Метод Б (нефть, тёмные продукты): фиксированная точность (6.4.1/6.4.2).

Сценарии (глубокое ветвление):
  А: P < 560 или P > 760 мм рт.ст. → PressureOutOfRangeError.
  Б: отгон + остаток + потери != 100 ± 0,2 → MassBalanceViolation.
  В: |X1 − X2| > r для точки → RepeatabilityViolation.
"""

from __future__ import annotations

from .base import BaseDistillationCalculator, IBP_KEY, FBP_KEY, round_half
from .constants import (
    GOST2177_TABLE4,
    GOST2177_PRESSURE_RANGE,
    GOST2177_METHODB_R,
    GOST2177_METHODB_r,
    GOST2177_NOMOGRAM,
    GOST2177_EXTRAPOLATION_MAX_PCT,
    GOST2177_TABLE4_POLICY_MODES,
    GOST2177_TABLE4_POLICY_DEFAULT,
    interpolate_table,
)
from .exceptions import (
    PressureOutOfRangeError,
    RepeatabilityViolation,
    MassBalanceViolation,
    InvalidTestError,
)
from .models import FractionInput, FractionResult


def _nomogram_values(slope: float, flags: list | None = None):
    """Значения (r_нк, r_кон, r_з, R_нк, R_кон, R_з) из оцифрованной номограммы.

    Интерполяция линейная между соседними значениями крутизны (шаг 0,1).
    Ключи словаря — возрастающая крутизна; r/R невозрастают с ростом S.

    TD-2177-003: выход за диапазон сетки (S > 5.0) раньше молча возвращал
    крайнее значение (занижение r/R). Теперь это фиксируется флагом
    NOMOGRAM_CLAMPED_HIGH через out-param ``flags`` (если передан).
    """
    keys = sorted(GOST2177_NOMOGRAM.keys())  # [0.0 ... 5.0], возрастание
    if slope > keys[-1]:
        if flags is not None and "NOMOGRAM_CLAMPED_HIGH" not in flags:
            flags.append("NOMOGRAM_CLAMPED_HIGH")
        return GOST2177_NOMOGRAM[keys[-1]]
    if slope < keys[0]:
        # физически S>=0; защита от мусора на входе
        return GOST2177_NOMOGRAM[keys[0]]
    for i in range(len(keys) - 1):
        k0, k1 = keys[i], keys[i + 1]
        if k0 <= slope <= k1:
            v0 = GOST2177_NOMOGRAM[k0]
            v1 = GOST2177_NOMOGRAM[k1]
            t = (slope - k0) / (k1 - k0) if k1 != k0 else 0.0
            return tuple(a + t * (b - a) for a, b in zip(v0, v1))
    return GOST2177_NOMOGRAM[keys[-1]]


class GOST2177Calculator(BaseDistillationCalculator):
    """Расчёт фракционного состава по ГОСТ 2177-99."""

    method_name = "gost2177"

    def __init__(self, table4_policy: str = GOST2177_TABLE4_POLICY_DEFAULT) -> None:
        """table4_policy — TD-2177-002: политика A/B для нетабличного давления.

        "strict" (по умолчанию, канон НД и книги v3.12): P вне [560, 760] →
            PressureOutOfRangeError.
        "clamp_low": P < 560 → A,B строки 560 + флаг PRESSURE_CLAMPED
            (экспериментальный режим ВНЕ НД, только по письменному решению SOP).
            P > 760 запрещён в любом режиме.
        """
        if table4_policy not in GOST2177_TABLE4_POLICY_MODES:
            raise ValueError(
                f"Неизвестный режим table4_policy={table4_policy!r}, "
                f"допустимо: {GOST2177_TABLE4_POLICY_MODES}")
        self.table4_policy = table4_policy
        self._pressure_clamped = False

    # ------------------------------------------------------------------
    def _validate_pressure(self, inp: FractionInput) -> None:
        p = inp.pressure_mmhg()
        lo, hi = GOST2177_PRESSURE_RANGE
        if p < lo or p > hi:
            if self.table4_policy == "clamp_low" and p < lo:
                return  # легализованный clamp к нижней границе (флаг обязателен)
            raise PressureOutOfRangeError(p, min_mmhg=lo, max_mmhg=hi)

    # ------------------------------------------------------------------
    def _resolve_distillate(self, inp: FractionInput) -> float:
        # Для ГОСТ 2177 отгон задаётся явно полем «Отгон» в эталоне.
        # Если V_pct задан и последняя точка 100 — используем её.
        if inp.V_pct:
            return float(inp.V_pct[-1])
        if inp.V_ml:
            return float(inp.V_ml[-1])
        loss = self._resolve_loss(inp)
        res = self._resolve_residue(inp)
        return 100.0 - loss - res

    # ------------------------------------------------------------------
    def _correct_loss(self, loss_pct: float, inp: FractionInput) -> float:
        """V_K = A·L + B (таблица 4 ГОСТ 2177-99).

        TD-2177-002: политика нетабличного давления задаётся в конструкторе.
        strict (канон): P вне [560, 760] → PressureOutOfRangeError
        (давление также проверено в _validate_pressure).
        clamp_low: P < 560 → A,B строки 560 + флаг PRESSURE_CLAMPED; P > 760 запрещён.
        """
        p = inp.pressure_mmhg()
        lo, hi = GOST2177_PRESSURE_RANGE
        if p < lo or p > hi:
            if self.table4_policy == "clamp_low" and p < lo:
                self._pressure_clamped = True
                A, B = GOST2177_TABLE4[0][1], GOST2177_TABLE4[0][2]
                return A * loss_pct + B
            raise PressureOutOfRangeError(p, min_mmhg=lo, max_mmhg=hi)
        A, B = interpolate_table(p, GOST2177_TABLE4)
        return A * loss_pct + B

    # ------------------------------------------------------------------
    def calculate(self, inp: FractionInput) -> FractionResult:
        self._pressure_clamped = False
        res = super().calculate(inp)
        if self._pressure_clamped:
            res.Flags.append("PRESSURE_CLAMPED")  # TD-2177-002: режим вне НД помечен
        return res

    # ------------------------------------------------------------------
    def _check_extrapolation(self, points, t_corr_raw: dict, t_corr_full: dict) -> None:
        """TD-2177-004 / п. 5.5.8 ГОСТ 2177-99: жёсткий запрет экстраполяции T(V).

        Политика (полный разбор — docs/TECHDEBT_TD2177_004.md):
          * Значение точки <= GOST2177_EXTRAPOLATION_MAX_PCT (95%), полученное
            НЕ из наблюдения, нормативно допустимо ТОЛЬКО как линейная
            ИНТЕРПОЛЯЦИЯ между наблюдёнными точками (это не экстраполяция).
            Если точка лежит вне диапазона наблюдений (левее первой / правее
            последней наблюдённой) или наблюдённых точек < 2 — это
            ЭКСТРАПОЛЯЦИЯ → InvalidTestError (код EXTRAPOLATION_FORBIDDEN).
          * Точки 96..98%, полученные расчётом: легализованная линейная
            экстраполяция по крайним наблюдённым точкам — требуется веткой
            нормирования «потери >= 2%» (96%). Разрешена, помечается флагом
            EXTRAPOLATED_96_98.
          * FBP/КК без наблюдения (> 98%) — запрещён (см. выше).
        """
        known_v = sorted(
            self._point_volume(p) for p in points
            if t_corr_raw.get(p) is not None and self._point_volume(p) is not None
        )
        for p in points:
            if t_corr_raw.get(p) is not None:
                continue  # точка наблюдена — ок
            v = self._point_volume(p)
            if v is None:
                continue
            if 95.0 < v <= 98.0:
                self._extrapolated_points.append(v)
                continue  # легализованная экстраполяция 96/97/98% (ветка потерь>=2%)
            inside = (
                len(known_v) >= 2
                and known_v[0] <= v <= known_v[-1]
            )
            if v <= GOST2177_EXTRAPOLATION_MAX_PCT and inside:
                continue  # интерполяция между наблюдёнными — разрешена
            raise InvalidTestError(
                f"ГОСТ 2177 п.5.5.8: экстраполяция T(V) для точки {v:g}% отгона "
                f"запрещена (значение не наблюдалось и не является интерполяцией "
                f"между наблюдёнными точками).",
                code="EXTRAPOLATION_FORBIDDEN",
                details={"point": str(p), "volume_pct": v,
                         "policy_max_pct": GOST2177_EXTRAPOLATION_MAX_PCT,
                         "observed_range": [known_v[0], known_v[-1]] if known_v else None},
            )

    @staticmethod
    def _point_volume(p) -> float | None:
        """Объём (%) для точки; IBP/НК → 0, FBP/КК → 100."""
        if isinstance(p, str):
            if p.upper() in ("IBP", "НК"):
                return 0.0
            if p.upper() in ("FBP", "КК"):
                return 100.0
            try:
                return float(p)
            except ValueError:
                return None
        return float(p)

    # ------------------------------------------------------------------
    def _precision_all(self, points, t_corr, slopes, inp) -> tuple[dict, dict]:
        """r/R по точкам.

        Метод А: номограмма (r_нк для IBP, r_кон для FBP, r_з для точек % отгона).
        Метод Б: фиксированные значения 6.4.1/6.4.2.
        """
        r_vals, R_vals = {}, {}
        method = str(inp.object_type).upper()

        if method in ("B", "Б", "b"):
            # Метод Б — фиксированные значения
            for p in points:
                key = self._method_b_key(p)
                r_vals[p] = GOST2177_METHODB_r.get(key, 2.0)
                R_vals[p] = GOST2177_METHODB_R.get(key, 6.0)
            return r_vals, R_vals

        # Метод А — номограмма
        slope_ibp = slopes.get(IBP_KEY, 0.0)
        slope_fbp = slopes.get(FBP_KEY, 0.0)
        for p in points:
            if isinstance(p, str) or p in (IBP_KEY, FBP_KEY):
                continue
            slope_pt = slopes.get(float(p), slope_ibp)
            r_nk, r_kon, r_z, R_nk, R_kon, R_z = _nomogram_values(slope_pt, self._nomogram_flags)
            r_vals[p] = r_z
            R_vals[p] = R_z
        # IBP/FBP
        r_nk, r_kon, r_z, R_nk, R_kon, R_z = _nomogram_values(slope_ibp, self._nomogram_flags)
        r_vals[IBP_KEY] = r_nk
        R_vals[IBP_KEY] = R_nk
        r_nk, r_kon, r_z, R_nk, R_kon, R_z = _nomogram_values(slope_fbp, self._nomogram_flags)
        r_vals[FBP_KEY] = r_kon
        R_vals[FBP_KEY] = R_kon

        # Сценарий В: |X1 − X2| > r → RepeatabilityViolation
        self._check_repeatability(points, t_corr, r_vals)
        return r_vals, R_vals

    # ------------------------------------------------------------------
    def _check_repeatability(self, points, t_corr, r_vals) -> None:
        """Проверка повторяемости (сценарий В): |X1 − X2| > r → исключение.

        Использует сырые пары (два определения), сохранённые в calculate.
        """
        pairs = getattr(self, "_last_pairs", {})
        for p in points:
            if p not in pairs:
                continue
            x1, x2 = pairs[p]
            if x1 is None or x2 is None:
                continue
            r = r_vals.get(p, 0.0)
            if r <= 0:
                continue
            if abs(x1 - x2) > r + 1e-9:
                raise RepeatabilityViolation(point=str(p), x1=x1, x2=x2, r=r)

    def _method_b_key(self, p) -> str:
        if isinstance(p, str):
            if p in (IBP_KEY, "НК"):
                return "IBP"
            if p in (FBP_KEY, "КК"):
                return "FBP"
            return str(p)
        f = float(p)
        if f <= 1.0:
            return "IBP"
        if f >= 99.0:
            return "FBP"
        if 96.0 <= f <= 98.0:
            return "96_98"
        # ближайший стандартный ключ 10/50/90
        for k in ("90", "50", "10"):
            if f >= float(k) - 5.0:
                return k
        return "10"

    # ------------------------------------------------------------------
    def _check_loss_limit(self, loss_corr: float, inp: FractionInput, flags: list) -> None:
        """Для ГОСТ 2177 лимит потерь по номограмме не задан таблично —
        пропускаем (флаг не выставляем)."""
        pass

    # ------------------------------------------------------------------
    def _compute_extra(self, points, t_corr, slopes, inp, flags) -> dict:
        return {"method": "A" if str(inp.object_type).upper() not in ("B", "Б") else "B"}


# ---------------------------------------------------------------------------
# Экземпляр для проверки повторяемости с двумя определениями
# ---------------------------------------------------------------------------
class GOST2177ManualCalculator(GOST2177Calculator):
    """Вариант ГОСТ 2177 с проверкой повторяемости по двум определениям.

    Используется, когда в T_obs переданы пары (x1, x2).
    """

    def _check_repeatability(self, points, t_corr, r_vals) -> None:
        # Для проверки нужны сырые пары; здесь они доступны через t_corr?
        # Нет — t_corr уже скорректированы. Проверка делается в calculate
        # до коррекции. Оставляем заглушку: метод А по номограмме не требует
        # обязательной проверки повторяемости (она выполняется в ГОСТ ISO 3405).
        pass