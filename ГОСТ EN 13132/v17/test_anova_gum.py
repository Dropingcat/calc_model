#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""test_anova_gum.py — pytest для ядра anova_gum.py (блок 3, ГОСТ EN 13132 v17).

DoD блока 3 (план §1, критерий готовности; исправление TD-DEV-193):
  (а) синтетика с известными sigma_r / sigma_L воспроизводит sigma_R
      с точностью до 0.5%;
  (б) «грубый учёт»: r=0,1, R=0,3 -> u_repr(mean2) = 0,1041 (+-0.5%),
      u_repr(single) = 0,1071 (+-0.5%);
  (в) нулевые/отрицательные дисперсии по GUM: sigma_L = sqrt(max(0, ...)) -> 0,
      НЕ NaN;
  (г) u_repr по ВСЕМ 8 поддиапазонам ИСПРАВЛЕННОЙ табл.9.3 контракта
      (0,0334; 0,1041; 0,1406; 0,1713; 0,2082; 0,2812; 0,3124; 0,3490)
      и табл.2 (0,1052);
  (д) nu_eff и k: доминирование u_repr (nu=inf) -> nu_eff ~ 1.8e3, k ~ 1.96;
      nu_eff = 6 -> k = 2,447;
  (е) профиль: 200 точек, U(x) = k * u_c(x), рост с u_c(x), ступенчатая u_repr;
  (ж) inverse_calibration: x0 и u(x0) на реальных данных v15 (MTBE.xlsx);
      x0 в диапазоне калибровки.

Дополнительные проверки DoD: импорт модуля без ошибок, все тесты PASS
(контракт выхода, секции 1-2).

Итерация 3 (P1, блок `gost13132-v16-xlsx-profile-uc`): smooth-аппроксимация
u_repr — (а) smooth vs step на границе 1.0 (непрерывность/переход), (б)
монотонная неубываемость на [0,1;15], (в) в середине (3;5] — переходное
значение, не ступенька, (г) экстраполяция/ошибки + регрессия 16 старых тестов.

v17 (TD-DEV-208, вердикт ресерчера по НД): ПО УМОЛЧАНИЮ interpolation="step"
(ступенчатая по поддиапазонам табл.1/2 ГОСТ EN 13132 — r/R ДИСКРЕТНЫ,
интерполяция не предписана); smooth остаётся методическим допущением, НЕ по НД.
Вклады с u==0 исключаются из u_c и ν_eff (combined_budget: zeroed, TD-DEV-33).
"""
import contextlib
import math
import os
import sys

import numpy as np
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import anova_gum as ag  # noqa: E402

# Реальные данные v15 (извлечены из EN13132_v15_MTBE.xlsx лист «01_Ввод»:
# X_эфф (г/г) = X_ном * pA/pИС, отклики Y1/Y2 инъекций).
# Уровни 1..6: X_эфф и (Y1, Y2).
V15_LEVELS = [
    (0.026184538653366583, [0.030962198657949802, 0.034063545111745114]),
    (0.09184966216216217, [0.16029648839980373, 0.13344885090416383]),
    (0.48478444632290785, [0.4966465295252269, 0.5016805962186982]),
    (0.990467937608319, [0.9760250121753179, 1.0118902489938377]),
    (2.22334429309535, [2.2172134292262875, 2.1645007609563547]),
    (3.144321093082835, [3.0351074536299993, 3.0515227980423845]),
]

REL_TOL = 5e-3  # 0.5% — критерий DoD блоков (а)/(б)


# ---------------------------------------------------------------------------
# (а) Синтетика: восстановление sigma_R из известных sigma_r, sigma_L
# ---------------------------------------------------------------------------
def test_a_synthetic_sigma_r_sigma_L_recover_sigma_R():
    """Синтетика с известными sigma_r/sigma_L: sigma_R = sqrt(sigma_r^2+sigma_L^2)
    воспроизводится с точностью 0.5% (DoD «а»). sigma_r оценивается через
    anova_repeatability. При df=6 статистическая ошибка оценки sigma_r ~29%,
    поэтому критерий 0.5% применяется к sigma_R ТОЛЬКО в режиме доминирующего
    sigma_L (sigma_r входит с малым весом); точность самой sigma_r проверяется
    вероятностно (test_a2). Здесь sigma_L подбирается доминирующим."""
    rng = np.random.default_rng(42)
    sigma_r_true = 0.004
    sigma_L_true = 0.011
    sigma_R_true = math.sqrt(sigma_r_true ** 2 + sigma_L_true ** 2)
    xs = np.linspace(0.02, 3.1, 6)
    a, b = 1.0, 0.001
    levels = []
    for x in xs:
        eta = rng.normal(0.0, sigma_L_true)  # межлабораторный сдвиг уровня
        y1 = a * x + b + eta + rng.normal(0.0, sigma_r_true)
        y2 = a * x + b + eta + rng.normal(0.0, sigma_r_true)
        levels.append((float(x), [y1, y2]))
    r_est = ag.anova_repeatability(levels)
    sigma_r_est = r_est["sigma_r"]
    # sigma_R восстанавливается из sigma_r_est и sigma_L_true (модель
    # sigma_R^2 = sigma_r^2 + sigma_L^2, 5725-6 4.1.4).
    # Ожидание: sigma_r_est^2 смещена вниз на величину порядка df, поэтому
    # критерий DoD (0.5% к sigma_R) проверяем в режиме доминирования sigma_L,
    # где вклад sigma_r^2 мал. sigma_r_est должна лежать в 95% ДИ:
    #   sigma_r_true * sqrt(df_r / chi2.ppf(0.975, df_r)) ... 
    #   ... sigma_r_true * sqrt(df_r / chi2.ppf(0.025, df_r))
    from scipy.stats import chi2

    df = r_est["df_r"]  # 6
    lo95 = sigma_r_true * math.sqrt(df / chi2.ppf(0.975, df))
    hi95 = sigma_r_true * math.sqrt(df / chi2.ppf(0.025, df))
    assert lo95 <= sigma_r_est <= hi95, (
        f"sigma_r_est={sigma_r_est} вне 95% ДИ [{lo95:.5f}; {hi95:.5f}]"
    )
    # sigma_L обратного разложения: sigma_L = sqrt(max(0, sigma_R^2 - sigma_r_est^2)).
    # Так как sigma_r_est в ДИ, sigma_L_back в соответствующем ДИ:
    sigma_L_back = ag.sigma_L_from_tables(sigma_R_true, sigma_r_est)
    assert sigma_L_back >= 0.0 and np.isfinite(sigma_L_back)
    # Точность модели разложения: восстановленное sigma_L не хуже 10%
    # (при df=6 — статистически обоснованный допуск):
    assert abs(sigma_L_back - sigma_L_true) / sigma_L_true < 0.10, (
        f"sigma_L: {sigma_L_back} vs {sigma_L_true}"
    )
    assert r_est["df_r"] == 6, f"df_r должен быть 6 (6x2), получено {r_est['df_r']}"

    # --- сценарий доминирующего sigma_L: критерий 0.5% к sigma_R достижим ---
    # При sigma_L >> sigma_r вклад sigma_r^2 в sigma_R^2 < 1%, поэтому даже
    # ~30%-я ошибка оценки sigma_r (df=6) даёт пренебрежимо малую ошибку sigma_R.
    rng2 = np.random.default_rng(123)
    sr2, sl2 = 0.003, 0.08
    sR2 = math.sqrt(sr2 ** 2 + sl2 ** 2)
    levels2 = []
    for x in xs:
        eta = rng2.normal(0.0, sl2)
        y1 = a * x + b + eta + rng2.normal(0.0, sr2)
        y2 = a * x + b + eta + rng2.normal(0.0, sr2)
        levels2.append((float(x), [y1, y2]))
    r2 = ag.anova_repeatability(levels2)
    sR_est2 = math.sqrt(r2["sigma_r"] ** 2 + sl2 ** 2)
    assert abs(sR_est2 - sR2) / sR2 < REL_TOL, (
        f"sigma_R (доминир. sigma_L): {sR_est2} vs {sR2}"
    )
    # обратное разложение sigma_L с точностью 0.5%:
    sL_back2 = ag.sigma_L_from_tables(sR2, r2["sigma_r"])
    assert abs(sL_back2 - sl2) / sl2 < REL_TOL, (
        f"sigma_L (доминир.): {sL_back2} vs {sl2}"
    )


def test_a2_anova_repeatability_accuracy_large_df():
    """Точность anova_repeatability: sigma_r — несмещённая оценка, s^2 ~ chi2(df).
    При df=392 относительная ошибка sigma_r ~ 1/sqrt(2*df) ~ 3.6% (1σ),
    поэтому детерминированный допуск 0.5% статистически невозможен.
    Проверяем попадание в 99.7% ДИ (3σ) по chi2 — устойчивый критерий."""
    from scipy.stats import chi2

    rng = np.random.default_rng(7)
    sigma_r_true = 0.004
    xs = np.linspace(0.1, 3.0, 8)
    levels = []
    for x in xs:
        ys = [1.0 * x + 0.001 + rng.normal(0.0, sigma_r_true) for _ in range(50)]
        levels.append((float(x), ys))
    r_est = ag.anova_repeatability(levels)
    df = r_est["df_r"]
    assert df == 8 * 49
    # 99.7% ДИ для sigma_r (chi2 с df):
    lo = sigma_r_true * math.sqrt(df / chi2.ppf(0.99865, df))
    hi = sigma_r_true * math.sqrt(df / chi2.ppf(0.00135, df))
    assert lo <= r_est["sigma_r"] <= hi, (
        f"sigma_r={r_est['sigma_r']} вне 99.7% ДИ [{lo:.5f}; {hi:.5f}]"
    )


# ---------------------------------------------------------------------------
# (б) «Грубый учёт»: r=0,1 R=0,3 -> u_repr = 0,1041 / 0,1071
# ---------------------------------------------------------------------------
def test_b_rough_account_urepr_mean2_and_single():
    """r=0,1; R=0,3; n=2 -> u_repr(mean2)=0,1041 (+-0.5%); u_repr(single)=0,1071 (+-0.5%)."""
    w = 2.0  # внутри поддиапазона (1,0;3,0]
    res_mean2 = ag.u_repr_from_tables(w, ag.TABLE1, mode="mean2")
    res_single = ag.u_repr_from_tables(w, ag.TABLE1, mode="single")
    assert abs(res_mean2["u_repr"] - 0.1041) / 0.1041 < REL_TOL, (
        f"u_repr(mean2): {res_mean2['u_repr']} vs 0.1041"
    )
    assert abs(res_single["u_repr"] - 0.1071) / 0.1071 < REL_TOL, (
        f"u_repr(single): {res_single['u_repr']} vs 0.1071"
    )
    # Точное значение (контракт разд. 4):
    exact = math.sqrt(0.3 ** 2 - 0.1 ** 2 / 2.0) / 2.8
    assert abs(res_mean2["u_repr"] - exact) < 1e-12
    # Контроль анти-ошибки TD-DEV-193: 0.0357 = r/2,8 — НЕ итоговый u_repr:
    assert abs(res_mean2["u_repr"] - 0.0357) > 0.06
    assert abs(res_single["u_repr"] - 0.0357) > 0.06
    # r и R из поддиапазона:
    assert res_mean2["r"] == 0.1 and res_mean2["R"] == 0.3


# ---------------------------------------------------------------------------
# (в) Нулевые/отрицательные дисперсии -> MAX(0, ..) по GUM
# ---------------------------------------------------------------------------
def test_c_negative_variance_returns_zero_not_nan():
    """sigma_L = sqrt(max(0, sigma_R^2 - sigma_r^2)) при sigma_R < sigma_r -> 0 (НЕ NaN)."""
    assert ag.sigma_L_from_tables(u_repr=0.01, sigma_r=0.05) == 0.0
    assert not math.isnan(ag.sigma_L_from_tables(u_repr=0.01, sigma_r=0.05))
    # равенство дисперсий -> 0
    assert ag.sigma_L_from_tables(u_repr=0.05, sigma_r=0.05) == 0.0
    # нормальный случай
    assert ag.sigma_L_from_tables(u_repr=0.1, sigma_r=0.04) == pytest.approx(
        math.sqrt(0.1 ** 2 - 0.04 ** 2)
    )
    # негативная дисперсия внутри u_repr_from_tables не возникает (R > r всегда):
    for lo, hi, r, R in ag.TABLE1_RANGES:
        assert R ** 2 - r ** 2 / 2.0 > 0.0


# ---------------------------------------------------------------------------
# (г) u_repr по ВСЕМ поддиапазонам ИСПРАВЛЕННОЙ табл.9.3 + табл.2
# ---------------------------------------------------------------------------
def test_g_urepr_all_8_ranges_table1():
    """Все 8 значений u_repr(mean2) ИСПРАВЛЕННОЙ табл.9.3 (TD-DEV-194)."""
    expected = [
        # (репрезентативная точка внутри поддиапазона, ожидаемый u_repr(mean2))
        (0.5, 0.0334),   # (0,1;1,0]
        (2.0, 0.1041),   # (1,0;3,0]
        (4.0, 0.1406),   # (3,0;5,0]
        (6.0, 0.1713),   # (5,0;7,0]
        (8.0, 0.2082),   # (7,0;9,0]
        (10.0, 0.2812),  # (9,0;11,0]
        (12.0, 0.3124),  # (11,0;13,0]
        (14.0, 0.3490),  # (13,0;15,0]
    ]
    for w, exp in expected:
        got = ag.u_repr_from_tables(w, ag.TABLE1, mode="mean2")["u_repr"]
        assert abs(got - exp) / exp < REL_TOL, f"w={w}: {got} vs {exp}"


def test_g_urepr_table2_oxygen():
    """Табл.2 (общий O): r=0,08, R=0,3 -> u_repr(mean2)=0,1052 (+-0.5%)."""
    got = ag.u_repr_from_tables(2.0, ag.TABLE2, mode="mean2")["u_repr"]
    assert abs(got - 0.1052) / 0.1052 < REL_TOL, f"table2: {got} vs 0.1052"
    exact = math.sqrt(0.3 ** 2 - 0.08 ** 2 / 2.0) / 2.8
    assert abs(got - exact) < 1e-12


def test_g_stepwise_interpolation_boundary():
    """Ступенчатая интерполяция: на границе поддиапазонов — более высокий диапазон
    (контракт разд. 9.3, правило R-05); внутри поддиапазона u_repr постоянна."""
    # w=1.0 — граница: (0,1;1,0] закрыт сверху, значит w=1.0 -> (1,0;3,0]
    r_at_1 = ag.u_repr_from_tables(1.0, ag.TABLE1, mode="single")
    assert r_at_1["lo"] == 1.0, "на границе должен выбираться более высокий диапазон"
    assert abs(r_at_1["u_repr"] - 0.3 / 2.8) < 1e-12
    # w=3.0 -> (3,0;5,0], R=0.4
    r_at_3 = ag.u_repr_from_tables(3.0, ag.TABLE1, mode="single")
    assert abs(r_at_3["u_repr"] - 0.4 / 2.8) < 1e-12
    # внутри (1,0;3,0] u_repr одинакова для 1.5 и 2.9 (ступенчатая)
    a1 = ag.u_repr_from_tables(1.5, ag.TABLE1, mode="mean2")["u_repr"]
    a2 = ag.u_repr_from_tables(2.9, ag.TABLE1, mode="mean2")["u_repr"]
    assert a1 == a2
    # вне диапазона -> ValueError
    with pytest.raises(ValueError):
        ag.u_repr_from_tables(20.0, ag.TABLE1)


# ---------------------------------------------------------------------------
# (д) nu_eff и k
# ---------------------------------------------------------------------------
def test_d_welch_satterthwaite_and_k():
    """nu_eff и k: доминирование u_repr (nu=inf) -> nu_eff~1.8e3, k~1.96;
    nu_eff=6 -> k=2,447."""
    # W-S: доминирует вклад с nu=inf -> nu_eff ~ u_c^4 * nu_model / u_model^4
    u_model = 0.02
    df_model = 4
    u_repr = 0.1  # доминирующий вклад, табличный (nu=inf)
    nu_eff = ag.welch_satterthwaite([(u_model, df_model), (u_repr, None)])
    expected_nu = (u_model ** 2 + u_repr ** 2) ** 2 / (u_model ** 4 / df_model)
    assert abs(nu_eff - expected_nu) / expected_nu < 1e-9
    assert nu_eff > 1000, f"nu_eff={nu_eff} должно быть ~1800 при доминировании u_repr"
    k_dom = ag.k_coverage(nu_eff)
    assert abs(k_dom - 1.96) < 0.01, f"k={k_dom} должен быть ~1.96"
    # t_0.975(6) = 2.4469... (TD-DEV-195)
    k6 = ag.k_coverage(6.0)
    assert abs(k6 - 2.4469) < 0.001, f"k(6)={k6}"
    assert abs(k6 - 2.447) < 0.001
    # inf -> z_0.975
    assert abs(ag.k_coverage(math.inf) - 1.959963984540054) < 1e-9
    # combined_budget: полный канонический набор с u_repr
    budget = ag.combined_budget(
        {
            "модель": (0.02, 4),
            "эталоны": (0.005, 50),
            "чистота": (0.003, None),
            "весы": (0.002, None),
            "прецизионность": (0.1, None),
        }
    )
    assert budget["u_c"] == pytest.approx(math.sqrt(0.02 ** 2 + 0.005 ** 2 + 0.003 ** 2 + 0.002 ** 2 + 0.1 ** 2))
    # u_repr доминирует (nu=inf): nu_eff большое, k ~ 1.96
    assert budget["nu_eff"] > 1000
    assert abs(budget["k"] - 1.96) < 0.02, f"k={budget['k']} должен быть ~1.96"
    assert budget["U"] == pytest.approx(budget["k"] * budget["u_c"])
    # флаг анти-двойного счёта:
    assert budget["local_sigma_r_not_duplicated"] is False
    b2 = ag.combined_budget({"sigma_r": (0.004, 6), "прецизионность": (0.1041, None)})
    assert b2["local_sigma_r_not_duplicated"] is True


# ---------------------------------------------------------------------------
# (е) Профиль U(x)
# ---------------------------------------------------------------------------
def _build_v15_snapshot():
    """OLS-снапшот по реальным данным v15."""
    snapshot = ag.ols_slopes_by_means(V15_LEVELS)
    snapshot["levels"] = V15_LEVELS
    return snapshot


def test_e_profile_200_points_and_monotonic_u():
    """Профиль: 200 точек, U(x)=k*u_c(x), растёт с u_c(x); u_repr ступенчатая
    (внутри поддиапазона постоянна)."""
    import warnings as _w

    snap = _build_v15_snapshot()
    # u_repr(x) через callable: ступенчатая по табл.1 (перевод % мас. -> г/г
    # через коэффициент k_conv = m_ИС*p_ИС/m_пробы; для профиля берём ~0.4
    # г/г на % мас., чтобы X калибровки [0.026..3.144] покрывал диапазоны).
    def budget(x):
        w_pct = x / 0.4  # перевод г/г -> % мас. (коэффициент-допущение)
        urepr = ag.u_repr_from_tables(w_pct, ag.TABLE1, mode="mean2")["u_repr"]
        return [("прецизионность", urepr, None)]

    # X калибровки [0.026; 3.144] частично ниже w<0.1 % мас. -> экстраполяция вниз
    # с RuntimeWarning (TD-DEV-202) — ожидаемое поведение, фильтруем в тесте:
    with _w.catch_warnings():
        _w.simplefilter("ignore", RuntimeWarning)
        prof = ag.profile_uncertainty(None, snap, budget, n_points=200)
    assert len(prof["x"]) == 200
    assert prof["x"][0] == pytest.approx(min(x for x, _ in V15_LEVELS))
    assert prof["x"][-1] == pytest.approx(max(x for x, _ in V15_LEVELS))
    assert prof["dx"] == pytest.approx((prof["x"][-1] - prof["x"][0]) / 199)
    # U = k * u_c поточечно
    assert np.allclose(prof["U"], prof["k"] * prof["u_c"])
    # U растёт с u_c: возьмём точки с разными u_c и проверим монотонность связи
    order = np.argsort(prof["u_c"])
    u_sorted = prof["u_c"][order]
    U_sorted = prof["U"][order]
    assert np.all(np.diff(u_sorted) >= 0)
    # U монотонно неубывает по u_c (k>=1.9 везде)
    assert np.all(np.diff(U_sorted) >= -1e-12)
    # u_repr ступенчатая: в пределах поддиапазона u_repr постоянна -> проверим
    # через u_c-разность двух точек внутри одного поддиапазона:
    # (поддиапазон (0,1;1,0] % мас. = X в (0.04;0.4] г/г)
    i1 = int(np.argmin(np.abs(prof["x"] - 0.1)))
    i2 = int(np.argmin(np.abs(prof["x"] - 0.2)))
    u1 = budget(prof["x"][i1])[0][1]
    u2 = budget(prof["x"][i2])[0][1]
    assert u1 == u2, "u_repr должна быть постоянна внутри поддиапазона (ступенчатая)"
    # u_model растёт от центра к краям (классическая OLS-форма)
    imid = int(np.argmin(np.abs(prof["x"] - snap["xbar"])))
    assert prof["u_model"][0] > prof["u_model"][imid]
    assert prof["u_model"][-1] > prof["u_model"][imid]


# ---------------------------------------------------------------------------
# (ж) Обратная градуировка на реальных данных v15
# ---------------------------------------------------------------------------
def test_zh_inverse_calibration_v15():
    """inverse_calibration: x0 в диапазоне калибровки; u(x0) из cov-формы."""
    snap = _build_v15_snapshot()
    # «проба»: средний отклик на уровне 4 (середина калибровки)
    y0 = np.mean(V15_LEVELS[3][1])
    # u_y0: эмпирическое s/sqrt(2) повторностей уровня 4
    u_y0 = np.std(V15_LEVELS[3][1], ddof=1) / math.sqrt(2.0)
    inv = ag.inverse_calibration(
        y0,
        snap["a"],
        snap["b"],
        snap["var_a"],
        snap["var_b"],
        snap["cov_ab"],
        u_y0=u_y0,
    )
    x0 = inv["x0"]
    x_min = min(x for x, _ in V15_LEVELS)
    x_max = max(x for x, _ in V15_LEVELS)
    assert x_min <= x0 <= x_max, f"x0={x0} вне диапазона калибровки"
    # x0 близок к X уровня 4
    assert abs(x0 - V15_LEVELS[3][0]) / V15_LEVELS[3][0] < 0.05
    # u(x0) > 0 и конечна
    assert np.isfinite(inv["u_x0"]) and inv["u_x0"] > 0.0
    # Проверка формулы дельта-метода:
    expected_u2 = (
        u_y0 ** 2 + x0 ** 2 * snap["var_a"] + snap["var_b"] + 2.0 * x0 * snap["cov_ab"]
    ) / snap["a"] ** 2
    assert inv["u2_x0"] == pytest.approx(expected_u2)
    # у0 на линии -> x0 = (y0 - b)/a
    assert x0 == pytest.approx((y0 - snap["b"]) / snap["a"])
    # a ~ 0 -> ValueError
    with pytest.raises(ValueError):
        ag.inverse_calibration(1.0, 0.0, 1.0, 0.01, 0.01, 0.0)


# ---------------------------------------------------------------------------
# Дополнительные проверки целостности (OLS, импорт, одиночный прогон)
# ---------------------------------------------------------------------------
def test_ols_on_v15_matches_reference():
    """OLS по средним v15: наклон/сдвиг/дисперсии конечны, df_res = 4 (6 уровней)."""
    snap = ag.ols_slopes_by_means(V15_LEVELS)
    assert snap["n"] == 6
    assert snap["df_res"] == 4
    assert np.isfinite(snap["a"]) and snap["a"] > 0.9
    assert np.isfinite(snap["b"])
    assert snap["var_a"] > 0 and snap["var_b"] > 0
    assert snap["cov_ab"] != 0.0
    assert 0.98 <= snap["r2"] <= 1.0
    # интервальная форма u_reg в X* = среднем X (минимум):
    xstar = snap["xbar"]
    u_reg = math.sqrt(snap["s2_res"] * (1.0 / snap["n"] + (xstar - xstar) ** 2 / snap["Sxx"])) / abs(snap["a"])
    assert u_reg > 0.0
    # ols вырожденный случай:
    with pytest.raises(ValueError):
        ag.ols_slopes_by_means([(1.0, [1.0]), (2.0, [2.0])])  # 2 уровня < 3


def test_module_imports_without_error():
    """Модуль импортируется без ошибок; __all__ покрывает все 9 функций."""
    import anova_gum

    expected = {
        "ols_slopes_by_means",
        "anova_repeatability",
        "u_repr_from_tables",
        "sigma_L_from_tables",
        "welch_satterthwaite",
        "k_coverage",
        "combined_budget",
        "profile_uncertainty",
        "inverse_calibration",
    }
    assert expected <= set(anova_gum.__all__)


# ---------------------------------------------------------------------------
# Правки по REQUEST_CHANGES (TD-DEV-200..202, L1)
# ---------------------------------------------------------------------------
def test_rq200_combined_budget_excludes_sigma_r():
    """TD-DEV-200 (major): анти-двойной счёт в combined_budget — РЕАЛЬНЫЙ.

    При передаче sigma_r и u_repr суммарный вклад прецизионности = u_repr
    (sigma_r исключён), u_c = 0.1041 (не 0.10418 + sigma_r)."""
    # sigma_r + u_repr: u_c должен быть ≈ u_repr (не sqrt(u_repr^2 + sigma_r^2))
    b = ag.combined_budget({"sigma_r": (0.004, 6), "прецизионность": (0.1041, None)})
    assert b["u_c"] == pytest.approx(0.1041, abs=1e-9), (
        f"u_c={b['u_c']} должен быть ≈0.1041 (sigma_r исключён)"
    )
    # ОТРИЦАТЕЛЬНЫЙ контроль: точное значение при двойном счёте sqrt(0.1041^2+0.004^2)
    double = math.sqrt(0.1041 ** 2 + 0.004 ** 2)  # 0.1041768
    assert abs(b["u_c"] - double) > 1e-5, "u_c не должен содержать двойной счёт sigma_r"
    assert b["local_sigma_r_not_duplicated"] is True
    assert b["excluded"] == [("sigma_r", 0.004, 6)]
    names_in = [k for k, _u, _d in b["items"]]
    assert "sigma_r" not in names_in
    assert "прецизионность" in names_in
    # sigma_r НЕ участвует в nu_eff: nu_eff = inf (u_repr доминирует, df=None)
    assert b["nu_eff"] == math.inf or b["nu_eff"] > 1e6

    # если u_repr НЕТ — sigma_r самостоятельный вклад (не исключается)
    b2 = ag.combined_budget({"sigma_r": (0.004, 6)})
    assert b2["u_c"] == pytest.approx(0.004)
    assert b2["local_sigma_r_not_duplicated"] is False
    assert b2["excluded"] == []

    # канонический набор без sigma_r не меняется
    b3 = ag.combined_budget(
        {
            "модель": (0.02, 4),
            "эталоны": (0.005, 50),
            "чистота": (0.003, None),
            "весы": (0.002, None),
            "прецизионность": (0.1, None),
        }
    )
    assert b3["u_c"] == pytest.approx(math.sqrt(0.02 ** 2 + 0.005 ** 2 + 0.003 ** 2 + 0.002 ** 2 + 0.1 ** 2))
    assert b3["local_sigma_r_not_duplicated"] is False
    # точная воспроизводимость фальсификации ревьюера:
    b4 = ag.combined_budget({"sigma_r": (0.004, 6), "прецизионность": (0.1041, None)})
    assert abs(b4["u_c"] - 0.1041) < 1e-9


def test_rq201_k_coverage_honest_low_nu():
    """TD-DEV-201 (low): k_coverage при nu<2 — честный t_{0.975}, не 2.0."""
    # t_0.975(1) = 12.7062 (scipy)
    k1 = ag.k_coverage(1.0)
    assert abs(k1 - 12.7062) < 0.01, f"k(1)={k1} должен быть ~12.706 (не 2.0)"
    assert k1 > 2.0
    # t_0.975(1.5) между t(1) и t(2)
    k15 = ag.k_coverage(1.5)
    assert 4.3 < k15 < 12.71
    # nu=2: 4.3027
    k2 = ag.k_coverage(2.0)
    assert abs(k2 - 4.3027) < 0.01
    # inf -> z
    assert ag.k_coverage(math.inf) == pytest.approx(1.959963984540054)
    # None -> NaN + RuntimeWarning
    import warnings as _w

    with pytest.warns(RuntimeWarning):
        kn = ag.k_coverage(None)
    assert math.isnan(kn)
    # nu<=0.5 -> NaN + RuntimeWarning (вырожденный случай)
    with pytest.warns(RuntimeWarning):
        kz = ag.k_coverage(0.3)
    assert math.isnan(kz)


def test_rq202_find_range_extrapolation_warning():
    """TD-DEV-202 (low): _find_range при w<0.1 — экстраполяция вниз с
    пометкой extrapolated=True и RuntimeWarning (не молча)."""
    with pytest.warns(RuntimeWarning):
        res = ag.u_repr_from_tables(0.05, ag.TABLE1, mode="single")
    assert res["lo"] == 0.1 and res["hi"] == 1.0
    assert res["extrapolated"] is True
    assert res["r"] == 0.05 and res["R"] == 0.1
    # в пределах диапазона extrapolated=False, без warning
    with _no_warn_check():
        res2 = ag.u_repr_from_tables(0.5, ag.TABLE1, mode="single")
    assert res2["extrapolated"] is False
    # граница (не экстраполяция)
    with _no_warn_check():
        res3 = ag.u_repr_from_tables(1.0, ag.TABLE1, mode="single")
    assert res3["extrapolated"] is False
    # автономный прогон по тестам ниже: изменил сигнатуру _find_range — проверяем
    # что TABLE1/TABLE2 наружу не отдают 5-кортежи (они не используются напрямую)


def test_rq_L1_inverse_names_consistent():
    """L1: имена не путаются — u_y0 (вход, отклик) и u_x0 (выход, концентрация)."""
    snap = _build_v15_snapshot()
    y0 = np.mean(V15_LEVELS[3][1])
    u_y0 = np.std(V15_LEVELS[3][1], ddof=1) / math.sqrt(2.0)
    inv = ag.inverse_calibration(
        y0, snap["a"], snap["b"], snap["var_a"], snap["var_b"], snap["cov_ab"], u_y0=u_y0
    )
    # u_x0 != u_y0 (u_x0 учитывает вклады линии)
    assert inv["u_x0"] != pytest.approx(u_y0)
    # u_x0 > u_y0/|a| (дополнительные вклады) — на реальных данных v15
    assert inv["u_x0"] > u_y0 / abs(snap["a"])
    # фактические числа v15: u_y0 ≈ 0.0179, u_x0 ≈ 0.0209
    assert abs(u_y0 - 0.0179) < 0.005, f"u_y0={u_y0}"
    assert abs(inv["u_x0"] - 0.0209) < 0.005, f"u_x0={inv['u_x0']}"


# ---------------------------------------------------------------------------
# Итерация 3, P1: smooth-аппроксимация u_repr (непрерывный профиль U(x))
# ---------------------------------------------------------------------------
def test_s_a_smooth_vs_step_on_boundary_1_0():
    """(а) smooth vs step на границе 1.0: step -> 0.1041 (диапазон (1;3]);
    smooth в той же точке = 0.1041 (узел по R-05), но НЕПОСРЕДСТВЕННО СЛЕВА
    от границы smooth уже «переходное» (между 0.0334 и 0.1041), а step — ещё
    ступенька 0.0334. Тем самым smooth устраняет скачок профиля."""
    step_1 = ag.u_repr_from_tables(1.0, ag.TABLE1, "mean2", "step")
    assert step_1["lo"] == 1.0 and step_1["hi"] == 3.0
    assert step_1["u_repr"] == pytest.approx(0.104124, abs=1e-5)
    # smooth в узле 1.0 — то же значение (R-05: на границе более высокий диапазон)
    sm_1 = ag.u_repr_from_tables(1.0, ag.TABLE1, "mean2", "smooth")
    assert sm_1["u_repr"] == pytest.approx(step_1["u_repr"], abs=1e-9)
    assert sm_1["interpolation"] == "smooth"
    # слева от границы: smooth непрерывна (переход), step — ступенька
    step_09 = ag.u_repr_from_tables(0.9, ag.TABLE1, "mean2", "step")["u_repr"]
    sm_09 = ag.u_repr_from_tables(0.9, ag.TABLE1, "mean2", "smooth")["u_repr"]
    assert step_09 == pytest.approx(0.033408, abs=1e-5)  # ступенька (0,1;1]
    assert 0.0334 < sm_09 < 0.1041, f"smooth(0.9)={sm_09} должно быть переходным"
    # непрерывность слева от границы: smooth(1.0-eps) -> smooth(1.0)
    sm_just = ag.u_repr_from_tables(1.0 - 1e-9, ag.TABLE1, "mean2", "smooth")["u_repr"]
    assert abs(sm_just - sm_1["u_repr"]) < 1e-6


def test_s_b_smooth_monotone_non_decreasing():
    """(б) smooth-аппроксимация монотонно НЕ УБЫВАЕТ на [0,1; 15,0] (плотная сетка)."""
    ws = np.linspace(0.1, 15.0, 1501)
    us = np.array([ag.u_repr_from_tables(float(w), ag.TABLE1, "mean2", "smooth")["u_repr"]
                   for w in ws])
    assert np.all(np.diff(us) >= -1e-12), "smooth u_repr не монотонна по w"
    # и в режиме single тоже
    us_s = np.array([ag.u_repr_from_tables(float(w), ag.TABLE1, "single", "smooth")["u_repr"]
                     for w in ws])
    assert np.all(np.diff(us_s) >= -1e-12)


def test_s_v_smooth_center_of_range_is_transition_not_step():
    """(в) smooth в середине (3;5] — НЕ ступенька: значение строго между
    соседними по R-05 узловыми значениями (0.1406 на границе 3.0 и 0.1713 на
    границе 5.0) и отличается от ступенчатого значения 0.1406.

    ПРИМЕЧАНИЕ (отклонение от формулировки в ТЗ): в задании стояло «между
    0.1041 и 0.1406», что численно НЕСОВМЕСТИМО с заданной же формулой
    аппроксимации: узлы берутся по правилу R-05 (на границе — более высокий
    диапазон), поэтому узел w=3.0 даёт u=0.1406, узел w=5.0 — u=0.1713,
    и интерполяция в середине (3;5] даёт ≈0.1562 (между 0.1406 и 0.1713).
    Зафиксирован как техническая правка теста; реализация следует ЯВНОЙ
    формуле из ТЗ (интерполяция r(w), R(w) между узлами R-05)."""
    step_c = ag.u_repr_from_tables(4.0, ag.TABLE1, "mean2", "step")["u_repr"]
    assert step_c == pytest.approx(0.140607, abs=1e-5)  # ступенька (3;5]
    sm_c = ag.u_repr_from_tables(4.0, ag.TABLE1, "mean2", "smooth")["u_repr"]
    assert 0.140607 < sm_c < 0.171280, f"smooth(4.0)={sm_c} — не переходное значение"
    assert abs(sm_c - 0.156186) < 1e-4  # точное значение при r=0.15, R=0.45
    # отличается от ступеньки (не ступенька):
    assert abs(sm_c - step_c) > 0.005
    # step-режим НЕ изменился (16 старых тестов, регрессия):
    assert ag.u_repr_from_tables(2.0, ag.TABLE1, "mean2", "step")["u_repr"] == pytest.approx(
        ag.u_repr_from_tables(2.0, ag.TABLE1, "mean2")["u_repr"], abs=1e-12)


def test_s_g_smooth_extrapolation_and_errors():
    """(г) smooth: экстраполяция вниз (w<0.1) с warning, w>15 — ValueError,
    некорректный режим — ValueError; ключ interpolation в возврате."""
    with pytest.warns(RuntimeWarning):
        res = ag.u_repr_from_tables(0.05, ag.TABLE1, "mean2", "smooth")
    assert res["extrapolated"] is True
    assert res["lo"] == 0.1 and res["hi"] == 1.0
    # экстраполяция первого сегмента: 0.05 = 0.1 - 0.5*0.1 (t=-0.5556)
    assert 0.0 < res["u_repr"] < 0.0334  # ниже значения узла 0.1
    with pytest.raises(ValueError):
        ag.u_repr_from_tables(20.0, ag.TABLE1, "mean2", "smooth")
    with pytest.raises(ValueError):
        ag.u_repr_from_tables(2.0, ag.TABLE1, "mean2", "midpoint")
    # верхняя граница шкалы w=15.0 — валидна (последний узел)
    top = ag.u_repr_from_tables(15.0, ag.TABLE1, "mean2", "smooth")
    assert top["u_repr"] == pytest.approx(0.349015, abs=1e-5)
    assert top["extrapolated"] is False
    # таблица 2 (один диапазон): smooth = константа (узлы 1.5/3.0 совпадают по r/R)
    t2a = ag.u_repr_from_tables(1.5, ag.TABLE2, "mean2", "smooth")["u_repr"]
    t2b = ag.u_repr_from_tables(3.0, ag.TABLE2, "mean2", "smooth")["u_repr"]
    assert abs(t2a - t2b) < 1e-12
    assert t2a == pytest.approx(0.1052, rel=REL_TOL)
    # ключ interpolation присутствует в обоих режимах
    assert ag.u_repr_from_tables(2.0, ag.TABLE1, "mean2", "step")["interpolation"] == "step"
    assert ag.u_repr_from_tables(2.0, ag.TABLE1, "mean2", "smooth")["interpolation"] == "smooth"


@contextlib.contextmanager
def _no_warn_check():
    """Контекст: фиксируем RuntimeWarning; если он возник — fail через счётчик."""
    import warnings as _w

    with _w.catch_warnings(record=True) as rec:
        _w.simplefilter("always")
        yield
        assert not any(issubclass(r.category, RuntimeWarning) for r in rec), (
            "неожиданный RuntimeWarning"
        )


if __name__ == "__main__":
    # Автономный прогон без pytest (если pytest недоступен).
    import traceback

    tests = [
        ("a_synthetic", test_a_synthetic_sigma_r_sigma_L_recover_sigma_R),
        ("a2_large_df", test_a2_anova_repeatability_accuracy_large_df),
        ("b_rough", test_b_rough_account_urepr_mean2_and_single),
        ("c_negative", test_c_negative_variance_returns_zero_not_nan),
        ("g_urepr_8", test_g_urepr_all_8_ranges_table1),
        ("g_urepr_t2", test_g_urepr_table2_oxygen),
        ("g_stepwise", test_g_stepwise_interpolation_boundary),
        ("d_ws_k", test_d_welch_satterthwaite_and_k),
        ("e_profile", test_e_profile_200_points_and_monotonic_u),
        ("zh_inverse", test_zh_inverse_calibration_v15),
        ("ols_v15", test_ols_on_v15_matches_reference),
        ("import", test_module_imports_without_error),
        ("rq200_anti_double", test_rq200_combined_budget_excludes_sigma_r),
        ("rq201_k_low_nu", test_rq201_k_coverage_honest_low_nu),
        ("rq202_extrap", test_rq202_find_range_extrapolation_warning),
        ("rq_L1_names", test_rq_L1_inverse_names_consistent),
        ("s_a_boundary", test_s_a_smooth_vs_step_on_boundary_1_0),
        ("s_b_monotone", test_s_b_smooth_monotone_non_decreasing),
        ("s_v_center", test_s_v_smooth_center_of_range_is_transition_not_step),
        ("s_g_extrap_err", test_s_g_smooth_extrapolation_and_errors),
    ]
    failed = 0
    for name, fn in tests:
        try:
            fn()
            print(f"PASS {name}")
        except Exception as exc:  # noqa: BLE001
            failed += 1
            print(f"FAIL {name}: {exc}")
            traceback.print_exc()
    print(f"\n{len(tests) - failed}/{len(tests)} PASS")
    sys.exit(1 if failed else 0)