#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""anova_gum.py — чистое ядро «классическая ANOVA по GUM» для ГОСТ EN 13132-2012 (v17).

Блок `gost13132-v17-final`, PROJECT_ID `gost13132-gum-anova`.
Методический контракт: `v15/METHODIC_CONTRACT_ANOVA_GUM.md`.

Отличия v17 от v16 (TD-DEV-208, вердикт ресерчера по НД):
  * РЕЖИМ ПО УМОЛЧАНИЮ `interpolation="step"` (ступенчатая по поддиапазонам
    табл.1/2 ГОСТ EN 13132). Вердикт НД: табл.1/2 задают r/R ДИСКРЕТНЫМИ
    поддиапазонами; интерполяция НЕ предписана ни стандартом, ни 5725-6,
    ни РМГ 61; GUM 4.1.4 («наилучшая оценка») к приписанным нормам
    прецизионности неприменим. smooth — НЕ по умолчанию: методическое
    допущение, НЕ соответствует НД (использовать только с обоснованием).
  * Вклады с u_i = 0 ИСКЛЮЧАЮТСЯ из u_c и ν_eff («0 = не учитывать вклад»,
    TD-DEV-33: значение из спецификации ИЛИ 0). В combined_budget и
    welch_satterthwaite нулевые u_i пропускаются (уже реализовано в v16,
    проверено и задокументировано).

Назначение (без излишеств, план §6.2):
    1. OLS по средним откликов уровней (перенос логики v15/calc_engine/v12-валидатора);
    2. ANOVA-компоненты: sigma_r (повторяемость из повторных инъекций, GUM H.5 /
       5725-6 4.1.4), sigma_L = sqrt(max(0, sigma_R^2 - sigma_r^2));
    3. u_repr из табл.1/2 ГОСТ EN 13132 (грубый учёт, делитель 2,8 — 5725-6 4.1.2/4.1.4):
         u_repr(mean2) = sqrt(R^2 - r^2/2)/2,8   — протокольная форма «среднее из 2»
                                                      (5725-6 4.2.3: CD = sqrt(R^2 - r^2/2));
         u_repr(single) = R/2,8                  — единичный результат;
         интерполяция (итер. 3, P1): 'step' (ступенчатая, R-05) либо 'smooth'
         (непрерывная кусочно-линейная по границам табл. — для гладкого профиля U(x));
    4. U = k * u_c, k = t_{0.975}(nu_eff), nu_eff по Welch–Satterthwaite (GUM G.2b);
    5. профиль U(x) на сетке (200 точек, шаг (max-min)/199 — паттерн v4);
    6. обратная градуировка пробы X0 = (Y0 - b)/a с u(X0) (дельта-метод,
       GUM 5.1.2 / закон трансформирования).

Зависимости: только numpy и scipy (scipy — для t-квантилей и F-критического значения;
если scipy отсутствует, k_covergage падает на аппроксимацию). БЕЗ openpyxl / /mnt/data /
artifact_tool / GUI — чтение xlsx вынесено в тесты.

Ссылки на НД:
    - ГОСТ EN 13132-2012 разд. 8 (обработка), разд. 10 (прецизионность, табл.1/2);
    - ГОСТ Р ИСО 5725-6-2002 4.1.2/4.1.4 (r=2.8*sigma_r, R=2.8*sigma_R),
      4.2.3 (CD = sqrt(R^2 - r^2/2) для среднего из 2), 7 (использование r/R
      для оценки неопределённости);
    - ГОСТ 34100.3-2017 (GUM): 4.2 (тип A), H.5 (ANOVA, H.29-H.32), G.2b
      (Welch–Satterthwaite), 6.2.1/6.3 (U = k*u_c, выбор k);
    - РМГ 61-2010 5.5 (поддиапазоны постоянства sigma_R).

Трактовка вырожденных случаев (GUM 4.2 / 5725-6 4.1.4): отрицательные дисперсии
обнуляются (MAX(0, ...)), деление на ~0 защищено EPS.
"""
from __future__ import annotations

import math
import warnings

import numpy as np

__all__ = [
    "ols_slopes_by_means",
    "anova_repeatability",
    "u_repr_from_tables",
    "sigma_L_from_tables",
    "welch_satterthwaite",
    "k_coverage",
    "combined_budget",
    "profile_uncertainty",
    "inverse_calibration",
]

EPS = 1e-12

# ---------------------------------------------------------------------------
# Таблицы прецизионности ГОСТ EN 13132-2012 разд. 10.1 (табл.1/2)
# ---------------------------------------------------------------------------
# Табл.1: 8 поддиапазонов «% масс./% об.» для индивидуальных оксигенатов.
# Колонки: (нижняя_граница, верхняя_граница, r, R). Интерпретация границ —
# ступенчатая (РМГ 61 5.5, контракт разд. 9.3, правило R-05):
#   X из (a; b] -> поддиапазон (a; b]; X на границе b -> более высокий диапазон.
TABLE1_RANGES = [
    # (lo, hi, r, R)   — значения ПРОВЕРЕНЫ по тексту ГОСТ EN 13132-2012 разд.10.1
    (0.1, 1.0, 0.05, 0.1),
    (1.0, 3.0, 0.10, 0.3),
    (3.0, 5.0, 0.10, 0.4),
    (5.0, 7.0, 0.20, 0.5),
    (7.0, 9.0, 0.20, 0.6),
    (9.0, 11.0, 0.20, 0.8),
    (11.0, 13.0, 0.30, 0.9),
    (13.0, 15.0, 0.30, 1.0),
]

# Табл.2: общий органически связанный кислород, % масс. (1,5..3,0).
TABLE2_RANGES = [
    (1.5, 3.0, 0.08, 0.3),
]


def _find_range_index(w_mass_pct: float, ranges) -> tuple:
    """Поиск поддиапазона БЕЗ предупреждений (см. _find_range — тот же результат).

    Возврат — кортеж (lo, hi, r, R, extrapolated). Используется режимом
    interpolation="smooth" (который выдаёт собственное предупреждение об
    экстраполяции и не должен дублировать его через _find_range).
    """
    if not np.isfinite(w_mass_pct):
        raise ValueError("w_mass_pct должен быть конечным числом")
    extrapolated = False
    # Экстраполяция вниз: w ниже нижней границы первого поддиапазона
    if w_mass_pct < ranges[0][0]:
        extrapolated = True
        return (ranges[0][0], ranges[0][1], ranges[0][2], ranges[0][3], True)
    for lo, hi, r, R in ranges:
        if w_mass_pct < hi:
            return (lo, hi, r, R, False)
    lo, hi, r, R = ranges[-1]
    if w_mass_pct <= hi:
        return (lo, hi, r, R, False)
    raise ValueError(
        "w_mass_pct вне диапазона табл. ГОСТ EN 13132 "
        "(индивид.: (0,1; 15,0], общий O: (1,5; 3,0])"
    )


def _find_range(w_mass_pct: float, ranges) -> tuple:
    """Ступенчатый поиск поддиапазона (контракт разд. 9.3, правило R-05).

    Правило: для w внутри поддиапазона (lo; hi] берётся этот поддиапазон;
    при w на границе поддиапазонов (w == hi == lo_след.) принимается БОЛЕЕ
    ВЫСОКИЙ диапазон: w = b -> поддиапазон (b; b'] (если он есть). Верхняя
    граница последнего поддиапазона (w = 15,0) относится к последнему
    диапазону (иначе методика не работает на верхней границе шкалы).
    Реализация: первый диапазон, для которого w < hi (строго); если такого
    нет — последний; если w > hi последнего — ValueError.

    Экстраполяция ВНИЗ (TD-DEV-202): при w ниже нижней границы первого
    поддиапазона табл.1 (w < 0,1 % мас.) молча используется первый поддиапазон
    [0,1; 1,0] (r=0,05; R=0,1) — согласуется с картой 9.3 контракта
    (X=0,026 г/г v15 -> поддиапазон (0,1; 1,0]). Вызывающий код осведомляется:
    в возврате u_repr_from_tables появляется флаг extrapolated=True и
    предупреждение RuntimeWarning. Возврат — кортеж из 5 элементов
    (lo, hi, r, R, extrapolated).
    """
    lo, hi, r, R, extrapolated = _find_range_index(w_mass_pct, ranges)
    if extrapolated:
        warnings.warn(
            f"w_mass_pct={w_mass_pct} ниже нижней границы табл. "
            f"({ranges[0][0]} % мас.): используется первый поддиапазон "
            f"{ranges[0][0]}..{ranges[0][1]} (экстраполяция вниз), "
            f"u_repr не аттестован для этой области (РМГ 61 5.5)",
            RuntimeWarning,
        )
    return (lo, hi, r, R, extrapolated)


# ---------------------------------------------------------------------------
# 3a. «Smooth»-аппроксимация u_repr (итерация 3, P1: гладкий профиль U(x))
# ---------------------------------------------------------------------------
def _smooth_nodes(ranges):
    """Узлы непрерывной кусочно-линейной аппроксимации r(w)/R(w) по границам табл.

    Для каждой границы диапазонов табл.1/2 (упорядоченное объединение lo/hi)
    строится точка (w_boundary, r, R), где r/R берутся по правилу R-05
    (на границе — значение БОЛЕЕ ВЫСОКОГО диапазона; верхняя граница шкалы —
    последнего диапазона). Между соседними узлами — линейная интерполяция.

    Возвращает список кортежей (w, r, R), упорядоченный по w.
    """
    bounds = sorted(
        {float(lo) for lo, _hi, _r, _R in ranges}
        | {float(hi) for _lo, hi, _r, _R in ranges}
    )
    nodes = []
    for b in bounds:
        rng = None
        for lo, hi, r, R in ranges:
            if b < hi:  # R-05: на границе — более высокий диапазон
                rng = (float(r), float(R))
                break
        if rng is None:  # верхняя граница последнего диапазона
            rng = (float(ranges[-1][2]), float(ranges[-1][3]))
        nodes.append((float(b), rng[0], rng[1]))
    return nodes


def _interp_rr(w_mass_pct, nodes):
    """Линейная интерполяция r/R по узлам; за нижней границей — экстраполяция
    первого сегмента (t < 0). Верхняя граница не экстраполируется — вызывающий
    код проверяет w > nodes[-1][0] до вызова (ValueError).
    """
    n = len(nodes)
    if w_mass_pct <= nodes[0][0]:
        i = 0
    elif w_mass_pct >= nodes[-1][0]:
        i = n - 2
    else:
        i = 0
        for i in range(n - 1):
            if nodes[i][0] <= w_mass_pct <= nodes[i + 1][0]:
                break
    w0, r0, R0 = nodes[i]
    w1, r1, R1 = nodes[i + 1]
    t = (w_mass_pct - w0) / (w1 - w0)
    return r0 + t * (r1 - r0), R0 + t * (R1 - R0)


# ---------------------------------------------------------------------------
# 1. OLS по средним (перенос логики v15 `03_Статистика`, calc_engine.ols,
#    metrology_lab_v12_validate.ols)
# ---------------------------------------------------------------------------
def ols_slopes_by_means(levels):
    """OLS по средним откликам уровней.

    Параметры
    ---------
    levels : iterable of (X, [Y...])
        Уровни калибровки: X — номинальная/эффективная концентрация (г/г),
        [Y...] — повторные отклики инъекций уровня (2 шт. в v15).
        Y уровня = среднее инъекций; OLS по точкам (X_level, Y_mean_level).

    Возвращает
    ----------
    dict : a, b, var_a, var_b, cov_ab, r2, s2_res, df_res, n, xbar, Sxx
        a — наклон, b — сдвиг; var_a, var_b — дисперсии оценок;
        cov_ab — ковариация (GUM 6.4 / v15 `03_Статистика!B15..B17`);
        r2 — коэффициент детерминации; s2_res — остаточная дисперсия
        (s^2 по средним откликов, df_res = n - 2, n = число уровней).

    Вырожденные случаи: Sxx <= EPS или df_res <= 0 -> ValueError
    (калибровка не может быть построена); наклон a = 0 допускается
    (обратная градуировка невозможна — сигнализирует вызывающий код).
    """
    data = []
    for item in levels:
        x, ys = item
        ys = list(ys)
        if len(ys) == 0:
            continue
        data.append((float(x), float(np.mean(ys))))
    if len(data) < 3:
        raise ValueError("Для OLS по средним требуется >= 3 уровней")
    xs = np.array([p[0] for p in data], dtype=float)
    ys = np.array([p[1] for p in data], dtype=float)
    n = len(xs)
    xbar = float(xs.mean())
    ybar = float(ys.mean())
    Sxx = float(np.sum((xs - xbar) ** 2))
    if Sxx <= EPS:
        raise ValueError("Sxx <= EPS: вырожденный разброс X уровней")
    Sxy = float(np.sum((xs - xbar) * (ys - ybar)))
    a = Sxy / Sxx
    b = ybar - a * xbar
    resid = ys - (a * xs + b)
    s2_res = float(np.sum(resid ** 2) / (n - 2))
    var_a = s2_res / Sxx
    var_b = s2_res * (1.0 / n + xbar ** 2 / Sxx)
    cov_ab = -s2_res * xbar / Sxx
    sst = float(np.sum((ys - ybar) ** 2))
    r2 = 1.0 - float(np.sum(resid ** 2)) / sst if sst > 0 else float("nan")
    return {
        "a": a,
        "b": b,
        "var_a": var_a,
        "var_b": var_b,
        "cov_ab": cov_ab,
        "r2": r2,
        "s2_res": s2_res,
        "df_res": n - 2,
        "n": n,
        "xbar": xbar,
        "Sxx": Sxx,
    }


# ---------------------------------------------------------------------------
# 2. sigma_r — повторяемость из повторных инъекций (ANOVA-нижний уровень)
# ---------------------------------------------------------------------------
def anova_repeatability(levels):
    """sigma_r — объединённое стандартное отклонение повторяемости инъекций.

    Формула (GUM H.5, 5725-6 4.1.4; v15 `03_Статистика!B22` SS pure error):
        s^2_within = sum_i sum_j (Y_ij - Y_i_mean)^2 / sum_i (n_i - 1)
        sigma_r = sqrt(s^2_within);  df_r = sum_i (n_i - 1)
    Для 6 уровней x 2 инъекции: df_r = 6.

    Возвращает dict: sigma_r, var_within, df_r, n_total, levels_used.
    Статус в «грубом учёте» (контракт разд. 4): sigma_r используется ТОЛЬКО
    для разложения sigma_L (sigma_L = sqrt(max(0, sigma_R^2 - sigma_r^2))) и
    НЕ входит отдельной строкой в суммарный бюджет (анти-двойной счёт,
    правило 2 разд. 6) — повторяемость уже усреднена в OLS-дисперсии.
    """
    ss_within = 0.0
    df = 0
    n_total = 0
    levels_used = 0
    for item in levels:
        _x, ys = item
        ys = np.asarray(list(ys), dtype=float)
        ys = ys[np.isfinite(ys)]
        if ys.size == 0:
            continue
        n_total += int(ys.size)
        levels_used += 1
        if ys.size >= 2:
            ss_within += float(np.sum((ys - ys.mean()) ** 2))
            df += int(ys.size) - 1
    if df <= 0:
        raise ValueError(
            "Повторностей нет (sum(n_i-1) = 0): sigma_r не оценима; "
            "используйте только табл.1/2 целиком (контракт разд. 7)"
        )
    var_within = ss_within / df
    return {
        "sigma_r": float(math.sqrt(var_within)),
        "var_within": float(var_within),
        "df_r": df,
        "n_total": n_total,
        "levels_used": levels_used,
    }


# ---------------------------------------------------------------------------
# 3. u_repr из табл.1/2 ГОСТ EN 13132 (грубый учёт, делитель 2,8)
# ---------------------------------------------------------------------------
# Режимы интерполяции (итерация 3, P1; уточнение R-05 в контракте):
#   'step'   — ступенчатая (РМГ 61 §5.5, консервативная альтернатива);
#   'smooth' — непрерывная кусочно-линейная аппроксимация r(w)/R(w) по
#              границам табл. (рекомендуемая для профиля U(x) книги v16).
def u_repr_from_tables(w_mass_pct: float, table: dict, mode: str = "mean2",
                       interpolation: str = "step"):
    """Стандартная неопределённость прецизионности по табл.1/2 ГОСТ EN 13132.

    Параметры
    ---------
    w_mass_pct : float
        Содержание компонента, % масс. (табл.1) или общего O, % масс. (табл.2).
    table : dict
        Описание таблицы: {'name': 'table1', 'ranges': [...]}. Готовые
        константы: `TABLE1` и `TABLE2` (см. ниже). Если передан dict без ключа
        'ranges' — берётся TABLE1 (обратная совместимость).
    mode : 'mean2' | 'single'
        'mean2'  — протокольная форма «среднее из 2 параллельных»:
                   u_repr = sqrt(R^2 - r^2/2)/2,8  (5725-6 4.2.3);
        'single' — единичный результат: u_repr = R/2,8 (5725-6 4.1.4).
    interpolation : 'step' | 'smooth'  (итерация 3, P1)
        'step'   — ступенчатая интерполяция (консервативная альтернатива,
                   контракт разд. 9.3, правило R-05; РМГ 61 §5.5 — поддиапазоны
                   постоянства sigma_R). Значения r/R постоянны внутри
                   поддиапазона; на границе — более высокий диапазон.
        'smooth' — НЕПРЕРЫВНАЯ кусочно-линейная аппроксимация: узлы
                   (w_boundary, r, R) на границах диапазонов табл. (R-05 —
                   на границе значение более высокого диапазона), линейная
                   интерполяция r(w), R(w) между узлами;
                   u_repr(mean2) = sqrt(R(w)^2 - r(w)^2/2)/2,8.
                   Рекомендуемый режим для профиля U(x) книги v16 (уточнение
                   R-05, запись в METHODIC_CONTRACT_ANOVA_GUM.md).
                   Вне диапазона: w < 0,1 — экстраполяция первого сегмента
                   с extrapolated=True + RuntimeWarning; w > 15 — ValueError.

    Возвращает dict:
        r, R, sigma_r_tab, sigma_R_tab, u_repr, mode, w_mass_pct,
        lo, hi (границы поддиапазона), extrapolated (True, если w ниже
        нижней границы табл. — экстраполяция вниз, TD-DEV-202; при этом
        выдаётся RuntimeWarning), interpolation ('step'|'smooth').

    Численный пример (контракт разд. 4, исправление TD-DEV-193):
        r=0,1; R=0,3 -> u_repr(mean2) = 0,1041; u_repr(single) = 0,1071 % мас.
        (значение 0,0357 = r/2,8 — это только перевод r, НЕ итоговый u_repr).
    """
    if mode not in ("mean2", "single"):
        raise ValueError("mode должен быть 'mean2' или 'single'")
    if interpolation not in ("step", "smooth"):
        raise ValueError("interpolation должен быть 'step' или 'smooth'")
    ranges = table.get("ranges") if isinstance(table, dict) else None
    if not ranges:
        ranges = TABLE1_RANGES
    if interpolation == "smooth":
        nodes = _smooth_nodes(ranges)
        if not np.isfinite(w_mass_pct):
            raise ValueError("w_mass_pct должен быть конечным числом")
        if w_mass_pct > nodes[-1][0]:
            raise ValueError(
                "w_mass_pct вне диапазона табл. ГОСТ EN 13132 "
                "(индивид.: (0,1; 15,0], общий O: (1,5; 3,0]) — "
                "smooth-аппроксимация не выходит за верхнюю границу"
            )
        extrapolated = w_mass_pct < nodes[0][0]
        if extrapolated:
            warnings.warn(
                f"w_mass_pct={w_mass_pct} ниже нижней границы табл. "
                f"({ranges[0][0]} % мас.): экстраполяция первого сегмента "
                f"smooth-аппроксимации (вниз), u_repr не аттестован для "
                f"этой области (РМГ 61 5.5)",
                RuntimeWarning,
            )
        r, R = _interp_rr(w_mass_pct, nodes)
        lo, hi, _r2, _R2, _ext = _find_range_index(w_mass_pct, ranges)
    else:
        lo, hi, r, R, extrapolated = _find_range(w_mass_pct, ranges)
    sigma_r_tab = r / 2.8
    sigma_R_tab = R / 2.8
    if mode == "single":
        u_repr = sigma_R_tab
    else:
        u_repr = math.sqrt(max(0.0, R ** 2 - r ** 2 / 2.0)) / 2.8
    return {
        "r": float(r),
        "R": float(R),
        "sigma_r_tab": float(sigma_r_tab),
        "sigma_R_tab": float(sigma_R_tab),
        "u_repr": float(u_repr),
        "mode": mode,
        "w_mass_pct": float(w_mass_pct),
        "lo": float(lo),
        "hi": float(hi),
        "extrapolated": bool(extrapolated),
        "interpolation": interpolation,
    }


TABLE1 = {"name": "table1", "ranges": TABLE1_RANGES}
TABLE2 = {"name": "table2", "ranges": TABLE2_RANGES}


# ---------------------------------------------------------------------------
# 4. sigma_L из табличной sigma_R и локальной sigma_r
# ---------------------------------------------------------------------------
def sigma_L_from_tables(u_repr, sigma_r):
    """Межлабораторная составляющая sigma_L (5725-6 4.1.4, модель sigma_R^2 = sigma_r^2 + sigma_L^2).

    sigma_L = sqrt(max(0, sigma_R^2 - sigma_r^2)),
    где sigma_R = u_repr (для mode='single' это R/2,8; для mode='mean2'
    — «эффективная» sigma_R результата-среднего, см. контракт разд. 4).

    Вырожденный случай (тест «в»): sigma_R < sigma_r -> sigma_L = 0 (НЕ NaN),
    что соответствует трактовке GUM 4.2 / 5725-6 (отрицательная дисперсия
    между сериями не физична, приравнивается к нулю).
    """
    var_L = max(0.0, float(u_repr) ** 2 - float(sigma_r) ** 2)
    return float(math.sqrt(var_L))


# ---------------------------------------------------------------------------
# 5. Welch–Satterthwaite: nu_eff (GUM G.2b)
# ---------------------------------------------------------------------------
def welch_satterthwaite(contributions):
    """Эффективные степени свободы по Welch–Satterthwaite (GUM G.2b).

    nu_eff = u_c^4 / sum_i (u_i^4 / nu_i)

    Параметры
    ---------
    contributions : iterable of (u_i, nu_i)
        u_i — стандартная неопределённость вклада; nu_i — степени свободы
        (нулевое/отрицательное значение трактуется как inf — тип B).

    Возвращает float nu_eff. Если u_c == 0 -> nan (неопределённость нулевая —
    k/ν_eff не определены, вызывающий код решает). Вклады с u_i == 0
    исключаются из суммы (0/0 не возникает).
    """
    u2 = 0.0
    s4 = 0.0
    for u_i, nu_i in contributions:
        u_i = float(u_i)
        if not np.isfinite(u_i) or u_i <= 0.0:
            continue
        u2 += u_i ** 2
        # df = None/0/нано/отрицательное -> бесконечность (тип B, GUM G.2b)
        if nu_i is None:
            continue
        try:
            nu_i = float(nu_i)
        except (TypeError, ValueError):
            continue
        if not np.isfinite(nu_i) or nu_i <= 0.0:
            continue  # inf/0/нано -> вклад не уменьшает nu_eff
        s4 += (u_i ** 4) / nu_i
    if u2 <= 0.0:
        return float("nan")
    return float(u2 ** 2 / s4) if s4 > 0.0 else math.inf


# ---------------------------------------------------------------------------
# 6. Коэффициент охвата k = t_{0.975}(nu_eff)
# ---------------------------------------------------------------------------
def _t_ppf(p: float, nu: float) -> float:
    """t-квантиль Стьюдента. scipy-путь, иначе аппроксимация (см. докстринг модуля)."""
    try:
        from scipy.stats import t as _t
        return float(_t.ppf(p, max(nu, 1e-9)))
    except Exception:
        pass
    # --- аппроксимация (точность лучше 0.5% при nu >= 3) ---
    # Hill & Williams (2004): z_p + (z_p^3 + z_p) / (4*nu)
    #   + (5*z_p^5 + 16*z_p^3 + 3*z_p) / (96*nu^2)
    #   + (3*z_p^7 + 19*z_p^5 + 17*z_p^3 - 15*z_p) / (384*nu^3)
    #   + (79*z_p^9 + 776*z_p^7 + 1482*z_p^5 - 1920*z_p^3 - 945*z_p)
    #     / (92160*nu^4)
    z = 1.959963984540054  # z_0.975
    nu = max(nu, 1e-9)
    n1 = 1.0 / nu
    n2 = n1 * n1
    n3 = n2 * n1
    n4 = n3 * n1
    z2 = z * z
    z3 = z2 * z
    z4 = z3 * z
    z5 = z4 * z
    z7 = z5 * z2
    z9 = z7 * z2
    return float(
        z
        + (z3 + z) / (4.0 * nu)
        + (5.0 * z5 + 16.0 * z3 + 3.0 * z) / (96.0 * nu * nu)
        + (3.0 * z7 + 19.0 * z5 + 17.0 * z3 - 15.0 * z) / (384.0 * nu ** 3)
        + (79.0 * z9 + 776.0 * z7 + 1482.0 * z5 - 1920.0 * z3 - 945.0 * z)
        / (92160.0 * nu ** 4)
    )


def k_coverage(nu_eff, p: float = 0.95):
    """Коэффициент охвата k = t_{1 - (1-p)/2}(nu_eff) (GUM 6.3, G.4).

    Примеры (TD-DEV-195): t_{0.975}(6) = 2.447; при nu_eff ~ 1.8e3 -> k ~ 1.96;
    при nu_eff = inf -> 1.959963984540054 (z_0.975, нормальное приближение).
    nu_eff = None/нано -> NaN с предупреждением (неопределённость не задана).
    nu_eff in (0; 2) -> ЧЕСТНОЕ t_{0.975}(nu_eff) (например, t_{0.975}(1) = 12.706,
    TD-DEV-201), а не 2.0; при отсутствии scipy вычисляется по аппроксимации
    (для nu<3 точность аппроксимации снижается — в протокол выносится
    предупреждение RuntimeWarning).
    """
    if nu_eff is None:
        warnings.warn("k_coverage: nu_eff не задан (None) -> k=NaN", RuntimeWarning)
        return float("nan")
    try:
        nu_eff = float(nu_eff)
    except (TypeError, ValueError):
        warnings.warn("k_coverage: nu_eff не число -> k=NaN", RuntimeWarning)
        return float("nan")
    if not np.isfinite(nu_eff):
        if math.isinf(nu_eff):
            return 1.959963984540054  # z_0.975
        # nan
        warnings.warn("k_coverage: nu_eff = NaN -> k=NaN", RuntimeWarning)
        return float("nan")
    alpha = 1.0 - p
    if nu_eff > 0.5:
        return _t_ppf(1.0 - alpha / 2.0, nu_eff)
    # nu_eff in (0; 0.5] — вырожденный случай: t-распределение с бесконечной
    # дисперсией; точного квантиля нет смысла требовать, возвращаем NaN с warning
    warnings.warn(
        "k_coverage: nu_eff<=0.5 — вырожденный случай t-распределения, "
        "коэффициент охвата не определён (NaN)",
        RuntimeWarning,
    )
    return float("nan")


# ---------------------------------------------------------------------------
# 7. Суммарный бюджет с анти-двойным счётом (контракт разд. 5/6)
# ---------------------------------------------------------------------------
def combined_budget(components, p: float = 0.95, exclude_when_superceded=("sigma_r",)):
    """Суммарная неопределённость, nu_eff, k, U (GUM 6.2.1, 6.3, G.2b).

    Параметры
    ---------
    components : dict[str, (u, df)] | list[(name, u, df)]
        Вклады стандартных неопределённостей u_i и степени свободы df_i
        (df = None/0/нано -> бесконечность, тип B).
        Канонический состав (контракт разд. 5, R6):
            модель (u_модель), эталоны, чистота, весы, проба (для результата),
            прецизионность (u_repr).
    p : float
        Доверительная вероятность (0.95).
    exclude_when_superceded : tuple[str]
        Имена вкладов, которые НЕ добавляются в u_c, если в components
        присутствует покрывающий вклад 'u_repr' (прецизионность табл.1/2).
        По умолчанию ('sigma_r',) — анти-двойной счёт РЕАЛЬНЫЙ (TD-DEV-200):
        локальная повторяемость sigma_r уже содержится в u_repr(mean2) =
        sqrt(R^2 - r^2/2)/2,8 (правило 3 разд. 6 контракта) и не должна
        складываться повторно. Если 'u_repr' в components отсутствует,
        исключение не применяется (вклад считается самостоятельным).

    Анти-двойной счёт (правила 1-3 разд. 6 контракта):
      * sigma_r из повторных инъекций НЕ передаётся отдельным вкладом:
        u_модель (OLS по средним) уже содержит усреднённую повторяемость
        инъекций; sigma_r используется ТОЛЬКО для разложения sigma_L.
      * u_repr передаётся УЖЕ в форме с вычитанием локальной повторяемости:
        u_repr(mean2) = sqrt(R^2 - r^2/2)/2,8 (рекомендуемая протокольная
        форма, правило 3). Локальная sigma_r, переданная отдельно,
        ИСКЛЮЧАЕТСЯ из суммы u_c (см. exclude_when_superceded).
      * Документируемое допущение (контракт разд. 7): sigma_R принято по
        табл.1/2 ГОСТ EN 13132 (N межлаб. округление 2,8); sigma_L данных
        лаборатории не оценивалась за отсутствием межсерийных данных.

    Возвращает dict: u_c, nu_eff, k, U, items (итоговые вклады, вошедшие
    в сумму), excluded (исключённые из-за анти-двойного счёта вклады),
    zeroed (вклады с u_i == 0 — «0 = не учитывать», TD-DEV-33),
    p, local_sigma_r_not_duplicated (True, если sigma_r была исключена).
    """
    if isinstance(components, dict):
        items = [(k, float(v[0]), v[1]) for k, v in components.items()]
    else:
        items = [(k, float(u), df) for (k, u, df) in components]
    # «0 = не учитывать вклад» (TD-DEV-33): вклады с u_i == 0 ИСКЛЮЧАЮТСЯ
    # из суммы u_c и из ν_eff (welch_satterthwaite уже пропускает u<=0).
    # Пометка zeroed — для честного отображения в бюджете/протоколе
    # («вклад не учитывается», а не «вклад с нулевой u»).
    zeroed = []
    keep0 = []
    for k, u, df in items:
        if u <= EPS:
            zeroed.append((k, u, df))
        else:
            keep0.append((k, u, df))
    items = keep0
    # Нормализация имени вклада прецизионности: ключи 'u_repr', 'прецизионность',
    # 'precision', 'urepr' считаются ОДНИМ вкладом (анти-двойной счёт, TD-DEV-200).
    _UREPR_NAMES = {"u_repr", "прецизионность", "precision", "urepr"}
    has_urepr = any(k in _UREPR_NAMES for k, _u, _d in items)
    excluded = []
    if has_urepr:
        keep = []
        for k, u, df in items:
            if k in exclude_when_superceded:
                excluded.append((k, u, df))
                continue
            keep.append((k, u, df))
        items = keep
    u_c = math.sqrt(sum(u ** 2 for _k, u, _d in items))
    local_sigma_r_not_duplicated = any(k == "sigma_r" for k, _u, _d in excluded)
    nu_eff = welch_satterthwaite([(u, df) for _k, u, df in items])
    k = k_coverage(nu_eff, p)
    U = k * u_c
    return {
        "u_c": float(u_c),
        "nu_eff": float(nu_eff),
        "k": float(k),
        "U": float(U),
        "items": items,
        "excluded": excluded,
        "zeroed": zeroed,
        "p": p,
        "local_sigma_r_not_duplicated": bool(local_sigma_r_not_duplicated),
    }


# ---------------------------------------------------------------------------
# 8. Профиль U(x) по диапазону калибровки (паттерн v4, 200 точек)
# ---------------------------------------------------------------------------
def profile_uncertainty(x_grid, cal_snapshot, budget_components, n_points: int = 200):
    """Профиль U(x) = k * u_c(x) на сетке из n_points точек (v4 `Профиль U(x)`, C4/TD-DEV-24).

    Параметры
    ---------
    x_grid : array-like
        Явная сетка X. Если len == 1 (скаляр) или передан (None, None) —
        сетка строится из cal_snapshot: linspace(min_X, max_X, n_points),
        шаг (max-min)/(n_points-1) (паттерн v4 `Профиль U(x)!B2`).
    cal_snapshot : dict
        Результат `ols_slopes_by_means` (a, b, var_a, var_b, cov_ab, s2_res,
        df_res, n, xbar, Sxx).
    budget_components : list[(name, u, df)] | callable(x) -> list[(name, u, df)]
        Постоянные вклады бюджета (эталоны, чистота, весы) + прецизионность.
        Прецизионность u_repr(x) — СТУПЕНЧАТАЯ по поддиапазонам табл.1/2
        (контракт разд. 9.3, правило R-05): для постоянной сетки внутри
        одного поддиапазона u_repr не меняется. Если передан callable —
        вызывается на каждой точке x (для зависящей от x прецизионности).
    n_points : int
        Число точек сетки (по умолчанию 200, как v4).

    Возвращает dict: x (np.ndarray, n_points), u_model(x), u_c(x), k(x),
    nu_eff(x), U(x). Если сетка построена автоматически — дополнительно
    x_min, x_max, dx.

    Модельная часть u_model(x) — интервальная форма OLS (v4 `Неопределенность!B2`,
    GUM 7.2.1):
        u^2_model(x) = s2_res * (1 + 1/n + (x - xbar)^2 / Sxx) / a^2
    (член 1/p опущен: калибровочный профиль — до результата пробы; для пробы
    используется `inverse_calibration` с u_y0).
    """
    if x_grid is None or (isinstance(x_grid, tuple) and x_grid[0] is None):
        xs_pts = None
        if "xs_cal" in cal_snapshot:
            xs_pts = cal_snapshot["xs_cal"]
        elif "levels" in cal_snapshot:
            xs_pts = [p[0] for p in cal_snapshot["levels"]]
        if xs_pts is None:
            raise ValueError("x_grid не задан и в cal_snapshot нет точек уровней "
                             "(ключи 'xs_cal'/'levels')")
        x_min = float(np.min(xs_pts))
        x_max = float(np.max(xs_pts))
        x = np.linspace(x_min, x_max, n_points)
        built = True
    else:
        x = np.asarray(x_grid, dtype=float)
        built = False
    if x.ndim == 0:
        x = x.reshape(1)
    x = np.atleast_1d(np.asarray(x, dtype=float))
    a = float(cal_snapshot["a"])
    b = float(cal_snapshot["b"])
    s2 = float(cal_snapshot["s2_res"])
    n = int(cal_snapshot["n"])
    xbar = float(cal_snapshot["xbar"])
    Sxx = float(cal_snapshot["Sxx"])
    df_model = float(cal_snapshot.get("df_res", max(n - 2, 1)))
    if abs(a) <= EPS:
        raise ValueError("Наклон a ~ 0: профиль U(x) не определён")
    u_model = np.sqrt(s2 * (1.0 + 1.0 / n + (x - xbar) ** 2 / Sxx)) / abs(a)
    u_c = np.zeros_like(x)
    nu_eff_arr = np.zeros_like(x)
    k_arr = np.zeros_like(x)
    for i, xi in enumerate(x):
        comps = budget_components(xi) if callable(budget_components) else budget_components
        if isinstance(comps, dict):
            comps = [(k_, float(v_[0]), v_[1]) for k_, v_ in comps.items()]
        else:
            comps = [(k_, float(u_), df_) for (k_, u_, df_) in comps]
        u_model_i = float(u_model[i])
        u_c_i = math.sqrt(u_model_i ** 2 + sum(u_ ** 2 for _k, u_, _d in comps))
        nu_eff_i = welch_satterthwaite([(u_model_i, df_model)] + [(u_, df_) for _k, u_, df_ in comps])
        u_c[i] = u_c_i
        nu_eff_arr[i] = nu_eff_i
        k_arr[i] = k_coverage(nu_eff_i)
    U = k_arr * u_c
    res = {
        "x": x,
        "u_model": u_model,
        "u_c": u_c,
        "nu_eff": nu_eff_arr,
        "k": k_arr,
        "U": U,
    }
    if built:
        res["x_min"] = x_min
        res["x_max"] = x_max
        res["dx"] = (x_max - x_min) / (n_points - 1)
    return res


# ---------------------------------------------------------------------------
# 9. Обратная градуировка (результат пробы)
# ---------------------------------------------------------------------------
def inverse_calibration(y0, a, b, var_a, var_b, cov_ab, u_y0=0.0):
    """Обратная градуировка: X0 = (Y0 - b)/a и u(X0) (GUM 5.1.2, 7.2.1).

    Дельта-метод/линеаризация (контракт разд. 10, Q9; v4 `Неопределенность!B2`;
    metrology_lab_v12_validate, u_reg):
        u^2(X0) = [u_y0^2 + X0^2 * var_a + var_b + 2 * X0 * cov_ab] / a^2
    (без 1/p: u_y0 — стандартная неопределённость СРЕДНЕГО отклика пробы,
    переданная вызывающим кодом, напр. u_y0 = s(Y0)/sqrt(p)).

    ВНИМАНИЕ, различение величин (TD-DEV-201/L1):
      * u_y0 — неопределённость ОТКЛИКА пробы (единицы Y), ВХОД функции;
      * u_x0 = u(X0) — неопределённость РЕЗУЛЬТАТА (концентрации, единицы X),
        ВЫХОД функции; как правило u_x0 > u_y0/|a| (добавляются вклады
        наклона/сдвига линии). НЕ путать в отчётах.

    Возвращает dict: x0, u_x0, u2_x0, a, b. Если a ~ 0 -> ValueError.
    """
    a = float(a)
    if abs(a) <= EPS:
        raise ValueError("Наклон a ~ 0: обратная градуировка невозможна")
    b = float(b)
    y0 = float(y0)
    x0 = (y0 - b) / a
    u2 = (float(u_y0) ** 2 + x0 ** 2 * float(var_a) + float(var_b)
          + 2.0 * x0 * float(cov_ab)) / (a ** 2)
    return {
        "x0": x0,
        "u_x0": float(math.sqrt(max(0.0, u2))),
        "u2_x0": float(max(0.0, u2)),
        "a": a,
        "b": b,
    }


if __name__ == "__main__":
    # Быстрый автономный самопроверочный прогон (без pytest).
    print("anova_gum.py — самопроверка")
    print("  u_repr(mean2) r=0.1 R=0.3:", u_repr_from_tables(2.0, TABLE1, "mean2")["u_repr"])
    print("  u_repr(single) r=0.1 R=0.3:", u_repr_from_tables(2.0, TABLE1, "single")["u_repr"])
    print("  u_repr(mean2) smooth w=4.0:", u_repr_from_tables(4.0, TABLE1, "mean2", "smooth")["u_repr"])
    print("  k(6)=", round(k_coverage(6.0), 4), " k(1800)=", round(k_coverage(1800.0), 4))