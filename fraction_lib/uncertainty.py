"""GUM-модель неопределённости для ГОСТ 2177 (методы А и Б).

TD-2177-007: отдельная модель неопределённости по GUM (JCGM 100:2008),
согласованная с расчётным ядром fraction_lib.gost2177.

Модель величины Y = T_corr(V) (скорректированная температура отгона, °C):

    T_corr = MROUND(T_obs + ΔT_p + ΔT_m, 0.5)

где
    ΔT_p = 0.0009 * (101.3 - p_kPa) * (273 + t_bar)   — поправка на давление
                                                           (п. 5.6.2 / Таблица 4 A,B)
    ΔT_m = номограмма (рисунок 6, цифровая оцифровка; TD-2177-003)

Бюджет стандартных неопределённостей (компоненты, источник, тип, единицы):

    u1  наблюдение температуры T_obs:
        оценка из повторяемости опыта лаборатории (s_rep / sqrt(n)), тип A;
        по умолчанию half-width = 0.5 °C (цена деления термометра / 2,
        прямоугольное распределение, k = sqrt(3)).
    u2  барометрическое давление p:
        паспортная точность барометра ±0.5 кПа (ГОСТ 2177 п.4.3 допускает
        шкалу 0.5 кПа), прямоугольное → u = 0.5/sqrt(3)/... см. ниже:
        вводится как half-width в кПа, закон — прямоугольный.
    u3  ртутный термометр барометра t_bar:
        half-width 0.5 °C, прямоугольный.
    u4  поправка на выступающую нить (столбик ртути) barometric t_t:
        half-width 0.2 °C, прямоугольный (значений в НД нет, экспертная
        оценка; при отсутствии поправки компонент обнуляется).
    u5  цифровая номограмма ΔT_m:
        half-width 0.2 °C, прямоугольный — остаточная ошибка оцифровки
        рисунка 6 после независимой валидации KT-1..KT-4 (TD-2177-003);
        в узлах S=1.3–1.4 зафиксирован дефект до 0.3 °C, поэтому базовое
        значение консервативно взято 0.3 °C.
    u6  дискретизация округления MROUND(...,0.5):
        half-width 0.25 °C, прямоугольный (вклад ≈ 0.144 °C).

Чувствительности (аналитические производные c_i = dT_corr/dX_i):

    c_p  = -0.0009 * (273 + t_bar)              [°C/кПа]
    c_t  =  0.0009 * (101.3 - p_kPa)            [°C/°C]
    c_nom = 1.0                                  [°C/°C]
    c_round = 1.0 (полувклад через прямоугольный закон)

Комбинированная неопределённость:

    u_c = sqrt( u1^2 + (c_p*u_p)^2 + (c_t*u_t)^2 + u5^2 + u6^2 )

(u4 входит в c_t-группу давления как аддитивный член в законе распределения
барометрической цепи; здесь учтён отдельно как (c_t*u_tt)^2.)

Расширенная неопределённость U = k * u_c, по умолчанию k = 2 (~95 %).

Ограничения модели (задекларированы в docs/TECHDEBT_TD2177_007.md):
 - экстраполированные точки (EXTRAPOLATED_96_98) получают штрафной
   множитель u1 × EXTRAPOLATION_U_FACTOR (=2.0, экспертно);
 - при флагах PRESSURE_CLAMPED / NOMOGRAM_CLAMPED_HIGH модель помечает
   результат как внедиапазонный (valid_for_reporting=False).
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

SQRT3 = math.sqrt(3.0)

# --- Константы политики неопределённости (TD-2177-007) -----------------------
U_ROUND_HALFWIDTH = 0.25     # °C, полуинтервал MROUND(...,0.5)
U_NOMOGRAM_HALFWIDTH = 0.3   # °C, остаток оцифровки рис.6 (дефект S=1.3–1.4)
U_PRESSURE_HALFWIDTH_KPA = 0.5   # кПа, шкала барометра (п.4.3)
U_TBAR_HALFWIDTH = 0.5       # °C, цена деления барометра
U_TT_HALFWIDTH = 0.2         # °C, поправка выступающей нити (экспертно)
U_OBS_DEFAULT_HALFWIDTH = 0.5    # °C, если s_rep неизвестна (цена деления/… )
EXTRAPOLATION_U_FACTOR = 2.0     # штраф за легализованную экстраполяцию 96–98%
K_DEFAULT = 2.0                  # коэффициент расширенной неопределённости

COEFF_DT = 0.0009   # нормативный коэффициент kPa-формы поправки давления
P_STD_KPA = 101.3


@dataclass
class UncertaintyComponent:
    name: str          # код компонента (u1..u6)
    description: str
    kind: str          # 'A' | 'B'
    distribution: str  # 'normal' | 'rectangular'
    halfwidth: float   # полуширина интервала оценки (в своих единицах)
    sensitivity: float # c_i, приведённый к °C
    units_in: str      # единицы аргумента (°C, кПа)

    @property
    def u(self) -> float:
        """Стандартное неопределённость, приведённая к °C."""
        base = self.halfwidth / SQRT3 if self.distribution == "rectangular" else self.halfwidth
        return abs(self.sensitivity) * base


@dataclass
class GUMResult:
    components: list[UncertaintyComponent]
    uc: float                      # комбинированная, °C
    k: float
    U: float                       # расширенная, °C
    dof_eff: float                 # эффективные степени свободы (Welch–Satterthwaite)
    t95: float                     # Стьюдент 95% по nu_eff (приближённо)
    valid_for_reporting: bool      # False — есть флаги внедиапазонных режимов
    notes: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {
            "uc_C": round(self.uc, 3),
            "k": self.k,
            "U_C": round(self.U, 3),
            "nu_eff": round(self.dof_eff, 1),
            "t95": round(self.t95, 2),
            "valid_for_reporting": self.valid_for_reporting,
            "components": {c.name: round(c.u, 3) for c in self.components},
            "notes": self.notes,
        }


def _student_t95(nu: float) -> float:
    """Приближение двухсторонней 95% точки Стьюдента (Cornish–Fisher, nu>=1)."""
    if nu <= 0:
        return 2.0
    z = 1.959963985
    g1 = (z ** 3 + z) / 4.0
    g2 = (5 * z ** 5 + 16 * z ** 3 + 3 * z) / 96.0
    g3 = (3 * z ** 7 + 19 * z ** 5 + 17 * z ** 3) / 384.0
    return z + g1 / nu + g2 / nu ** 2 + g3 / nu ** 3


def evaluate_gum(
    p_kPa: float,
    t_bar_C: float,
    *,
    u_obs_halfwidth: float | None = None,
    n_repeats: int = 1,
    s_rep: float | None = None,
    flags: list[str] | None = None,
    k: float = K_DEFAULT,
    include_thread_correction: bool = True,
) -> GUMResult:
    """Собрать бюджет GUM для одной точки T_corr(V).

    Args:
        p_kPa: местное атмосферное давление, кПа.
        t_bar_C: температура барометра, °C.
        u_obs_halfwidth: полуширина интервала наблюдения (по умолчанию 0.5 °C).
        n_repeats / s_rep: если заданы — тип A: u1 = (s_rep/sqrt(n)) (нормальный,
            halfwidth трактуется как стандартное отклонение среднего).
        flags: Flags результата fraction_lib (PRESSURE_CLAMPED и т.п.).
        include_thread_correction: применять ли u4 (поправка выступающей нити).
    """
    flags = flags or []
    notes: list[str] = []

    # c_p:dT/dp, c_t: dT/dt_bar
    c_p = -COEFF_DT * (273.0 + t_bar_C)          # °C/кПа
    c_t = COEFF_DT * (P_STD_KPA - p_kPa)         # °C/°C

    # u1 — наблюдение
    if s_rep is not None and n_repeats >= 2:
        u1 = UncertaintyComponent("u1", "Наблюдение T_obs (тип A: s_rep/√n)",
                                  "A", "normal", s_rep / math.sqrt(n_repeats), 1.0, "°C")
        nu1 = n_repeats - 1
    else:
        hw = u_obs_halfwidth if u_obs_halfwidth is not None else U_OBS_DEFAULT_HALFWIDTH
        u1 = UncertaintyComponent("u1", "Наблюдение T_obs (цена деления, прямоугольный)",
                                  "B", "rectangular", hw, 1.0, "°C")
        nu1 = 0.0  # у прямоугольной без данных — бесконечные df, но conservative
    if "EXTRAPOLATED_96_98" in flags:
        u1 = UncertaintyComponent(u1.name + " ×extrapolation penalty",
                                  u1.description, u1.kind, u1.distribution,
                                  u1.halfwidth * EXTRAPOLATION_U_FACTOR,
                                  u1.sensitivity, u1.units_in)
        notes.append("EXTRAPOLATED_96_98: вклад наблюдения увеличен в "
                     f"{EXTRAPOLATION_U_FACTOR}× (экспертная политика).")

    comps = [
        u1,
        UncertaintyComponent("u2", "Барометрическое давление p", "B",
                             "rectangular", U_PRESSURE_HALFWIDTH_KPA, c_p, "кПа"),
        UncertaintyComponent("u3", "Температура барометра t_bar", "B",
                             "rectangular", U_TBAR_HALFWIDTH, c_t, "°C"),
        UncertaintyComponent("u5", "Цифровая номограмма ΔT_m (остаток оцифровки)",
                             "B", "rectangular", U_NOMOGRAM_HALFWIDTH, 1.0, "°C"),
        UncertaintyComponent("u6", "Дискретизация округления MROUND 0.5", "B",
                             "rectangular", U_ROUND_HALFWIDTH, 1.0, "°C"),
    ]
    if include_thread_correction:
        comps.insert(3, UncertaintyComponent(
            "u4", "Поправка выступающей нити barometer", "B",
            "rectangular", U_TT_HALFWIDTH, c_t, "°C"))

    uc2 = sum(c.u ** 2 for c in comps)
    uc = math.sqrt(uc2)

    # Welch–Satterthwaite: считаем все B-компоненты nu=inf, тип A — nu=n-1
    if nu1 > 0:
        num = uc2 ** 2
        den = uc2 ** 2
        # вычитаем вклад A из «inf»-части
        ua = u1.u ** 2
        den = (ua / nu1) + (uc2 - ua) / math.inf if ua else uc2 / math.inf
        nu_eff = num / den if den else math.inf
    else:
        nu_eff = math.inf

    t95 = _student_t95(nu_eff) if math.isfinite(nu_eff) else 1.96
    k_used = max(k, t95) if math.isfinite(t95) else k
    if k_used > k:
        notes.append(f"k повышен с {k} до {k_used:.2f} по ν_eff={nu_eff:.0f} "
                     "(Уэлч–Саттертуэйт).")

    valid = not any(f in flags for f in ("PRESSURE_CLAMPED", "NOMOGRAM_CLAMPED_HIGH"))
    if not valid:
        notes.append("Результат получен во внедиапазонном режиме политики "
                     f"(flags={[f for f in flags if f in ('PRESSURE_CLAMPED','NOMOGRAM_CLAMPED_HIGH')]});"
                     " не пригоден к отчёту без обоснования.")

    return GUMResult(components=comps, uc=uc, k=k_used, U=k_used * uc,
                     dof_eff=nu_eff, t95=t95, valid_for_reporting=valid, notes=notes)


# ===========================================================================
# TD-D86-006: GUM-модель неопределённости ASTM D86 (T_c(V), °C)
# ===========================================================================
"""Модель величины Y = T_c(V) — скорректированная наблюдаемая температура
отгона по ASTM D86-23 (Section 11):

    T_c = T_obs + C_p + ΔT_loss   (с округлением отчёта до 0.5 °C)

где C_p = 0.00012·(760 − P_mm)·(273 + t_bar) — поправка Янга на давление;
ΔT_loss — температурная поправка на потери/остаток, входящая через сдвиг
процента отгона и локальный наклон кривой slope (°C на %).

Бюджет (полуширины — константы D86_U_* в constants.py, прямоугольные,
кроме типа A при s_rep):

    u1 наблюдение T_obs            c=1
    u2 барометрическое давление P   c = -k_C·(273 + t_bar)   [°C/мм рт.ст.]
    u3 температура барометра t_bar  c =  k_C·(760 − P_mm)
    u4 мениск / выступающий столбик c=1 (экспертно)
    u5 отсчёт объёма V (мл)         c = slope·(100/V_flask)  → °C/% × %/мл
    u6 коррекция потерь ΔT_loss     c=1 (экспертно)
    u7 округление отчёта 0.5 °C     c=1 (halfwidth 0.25)

Флаги политики D86 (LOSS_CORR_UNSTABLE, TABLE5_INVALID:*, INVALID_TEST)
помечают результат как вне канона: valid_for_reporting=False.
"""

from .constants import (
    D86_YOUNG_K_MMHG, D86_U_OBS_HALFWIDTH, D86_U_BAROMETER_HALFWIDTH_MM,
    D86_U_TBAR_HALFWIDTH, D86_U_MENISCUS_HALFWIDTH, D86_U_VOLUME_HALFWIDTH,
    D86_U_LOSSCORR_HALFWIDTH, D86_U_ROUND_HALFWIDTH, LOSS_CORR_MMHG_STD,
)

D86_OUT_OF_CANON_FLAGS = ("LOSS_CORR_UNSTABLE", "TABLE5_INVALID", "INVALID_TEST")


def evaluate_gum_d86(
    p_mmHg: float,
    t_bar_C: float,
    slope_degC_per_pct: float,
    *,
    flask_volume_ml: float = 100.0,
    u_obs_halfwidth: float | None = None,
    n_repeats: int = 1,
    s_rep: float | None = None,
    flags: list[str] | None = None,
    k: float = K_DEFAULT,
) -> GUMResult:
    """Собрать бюджет GUM для одной точки T_c(V) ASTM D86.

    Args:
        p_mmHg: местное атмосферное давление, мм рт.ст.
        t_bar_C: температура барометра, °C.
        slope_degC_per_pct: локальный наклон кривой в точке (°C/%) —
            переводит ошибку отсчёта объёма в ошибку температуры.
        flask_volume_ml: залитый объём пробы (для пересчёта мл→%).
        s_rep/n_repeats: тип A по повторяемости (как в ГОСТ-модели).
        flags: флаги результата fraction_lib D86Calculator / d86_procedure.
    """
    flags = flags or []
    notes: list[str] = []

    c_p = -D86_YOUNG_K_MMHG * (273.0 + t_bar_C)          # °C / мм рт.ст.
    c_t = D86_YOUNG_K_MMHG * (LOSS_CORR_MMHG_STD - p_mmHg)  # °C / °C
    # мл → % отгона: u(V)=0.5 мл ≈ 0.5·(100/flask)% ; T(%) наклон slope ⇒
    # c_V = slope · (100/flask) [°C/мл]
    c_v = abs(slope_degC_per_pct) * (100.0 / flask_volume_ml)

    if s_rep is not None and n_repeats >= 2:
        u1 = UncertaintyComponent("u1", "Наблюдение T_obs (тип A: s_rep/√n)",
                                  "A", "normal", s_rep / math.sqrt(n_repeats), 1.0, "°C")
        nu1 = n_repeats - 1
    else:
        hw = u_obs_halfwidth if u_obs_halfwidth is not None else D86_U_OBS_HALFWIDTH
        u1 = UncertaintyComponent("u1", "Наблюдение T_obs (цена деления, прямоугольный)",
                                  "B", "rectangular", hw, 1.0, "°C")
        nu1 = 0.0

    comps = [
        u1,
        UncertaintyComponent("u2", "Барометрическое давление P (Янг)", "B",
                             "rectangular", D86_U_BAROMETER_HALFWIDTH_MM, c_p, "мм рт.ст."),
        UncertaintyComponent("u3", "Температура барометра t_bar", "B",
                             "rectangular", D86_U_TBAR_HALFWIDTH, c_t, "°C"),
        UncertaintyComponent("u4", "Мениск / выступающий столбик термометра", "B",
                             "rectangular", D86_U_MENISCUS_HALFWIDTH, 1.0, "°C"),
        UncertaintyComponent("u5", "Отсчёт объёма V (через локальный slope)", "B",
                             "rectangular", D86_U_VOLUME_HALFWIDTH, c_v, "мл"),
        UncertaintyComponent("u6", "Коррекция потерь/остатка ΔT_loss", "B",
                             "rectangular", D86_U_LOSSCORR_HALFWIDTH, 1.0, "°C"),
        UncertaintyComponent("u7", "Округление отчёта до 0.5 °C", "B",
                             "rectangular", D86_U_ROUND_HALFWIDTH, 1.0, "°C"),
    ]

    uc2 = sum(c.u ** 2 for c in comps)
    uc = math.sqrt(uc2)

    if nu1 > 0:
        ua = u1.u ** 2
        den = (ua / nu1) if ua < uc2 else uc2 / math.inf
        nu_eff = uc2 ** 2 / den if den else math.inf
    else:
        nu_eff = math.inf

    t95 = _student_t95(nu_eff) if math.isfinite(nu_eff) else 1.96
    k_used = max(k, t95) if math.isfinite(t95) else k
    if k_used > k:
        notes.append(f"k повышен с {k} до {k_used:.2f} по ν_eff={nu_eff:.0f} "
                     "(Уэлч–Саттертуэйт).")

    bad = [f for f in flags if any(f.startswith(pfx) for pfx in D86_OUT_OF_CANON_FLAGS)]
    valid = not bad
    if not valid:
        notes.append("Результат получен при флагах вне канона НД "
                     f"({', '.join(bad)}); не пригоден к отчёту без обоснования.")

    return GUMResult(components=comps, uc=uc, k=k_used, U=k_used * uc,
                     dof_eff=nu_eff, t95=t95, valid_for_reporting=valid, notes=notes)
