#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify_en13132_v17_charts.py — проверка EN13132_v17_CALC.xlsx (блок `gost13132-v17-final`).

Усиления итерации 3 (TD-DEV-207, P1, P3):
  * V2j: паттерн формулы `=G{r}*E{r}` проверяется через re (не только префикс "=");
  * V3q: сверка ВСЕХ 200 точек профиля U = k·u_c с ядром (не выборка 6);
  * V3r-новый: U(06_Итог) == U(x_pred) маркер («Графики»!D4) (±tol) — ловит TD-DEV-206;
  * V5: генерация при 4/5/8 уровнях (другой набор данных) не падает, ссылки диаграмм
        корректны (анти-хардкод P3);

Изменения v17 (TD-DEV-208, вердикт ресерчера по НД):
  * РЕЖИМ ПО УМОЛЧАНИЮ step: ядро вызывается с interpolation="step";
  * V6: профиль СТУПЕНЧАТЫЙ (step): u_repr(x) постоянна внутри поддиапазонов табл.1,
        столбец D совпадает с ядром во всех 200 точках, в шапке листа указан режим;
  * ожидаемые числа step (MTBE): u_repr(w0=3.967)=0.1406 %мас, U≈0.0798 г/г.

Эталонные числа — из модуля build_en13132_v17_calc.py (compute_calibration) по
уровням/входам, прочитанным ИЗ САМОЙ КНИГИ (01_Ввод — единый источник, P3).
Выход: 0 = все PASS, 1 = есть FAIL. Каждая строка "[PASS]"/"[FAIL]" + детали.
"""
from __future__ import annotations

import os
import re
import sys
import tempfile
import warnings

import numpy as np
import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import anova_gum as ag  # noqa: E402
import build_en13132_v17_calc as bld  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows-консоль: UTF-8
except Exception:
    pass

V17_DIR = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(V17_DIR, "EN13132_v17_CALC.xlsx")

PASS, FAIL = "PASS", "FAIL"
_results = []


def check(name, ok, detail=""):
    _results.append((name, bool(ok), detail))
    tag = PASS if ok else FAIL
    print(f"[{tag}] {name}" + (f" — {detail}" if detail else ""))


# ---------------------------------------------------------------------------
# Парсинг ссылок серий диаграмм (openpyxl)
# ---------------------------------------------------------------------------
_COL_RE = re.compile(r"^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$")
_RNG_RE = re.compile(r"^(?:'((?:[^']|'')+)'|([^!]+))!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?$")


def _col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def parse_range(f):
    if not f:
        return None
    m = _COL_RE.search(f.split("!")[-1])
    if not m:
        return None
    c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), m.group(4)
    if c2 is None:
        c2, r2 = c1, r1
    else:
        r2 = int(r2)
    sm = _RNG_RE.match(f)
    sheet = None
    if sm:
        sheet = (sm.group(1) or sm.group(2)).replace("''", "'")
    return {
        "sheet": sheet,
        "min_col": _col_to_num(c1),
        "max_col": _col_to_num(c2),
        "min_row": r1,
        "max_row": r2,
    }


def series_ref(series, which):
    obj = getattr(series, "yVal", None) if which == "y" else getattr(series, "xVal", None)
    if obj is None:
        return None
    numref = getattr(obj, "numRef", None)
    if numref is None:
        return None
    return parse_range(getattr(numref, "f", None))


def chart_title_text(chart):
    try:
        t = chart.title
        if t is None:
            return None
        return t.tx.rich.p[0].r[0].t
    except Exception:
        return None


def cells_numeric(wb_f, sheet, min_col, max_col, min_row, max_row):
    wb_d = openpyxl.load_workbook(wb_f, data_only=True)
    wb_v = openpyxl.load_workbook(wb_f, data_only=False)
    sh_d = wb_d[sheet]
    sh_v = wb_v[sheet]
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            v = sh_v.cell(r, c).value
            if v is None or v == "":
                return False, f"{sheet}!{openpyxl.utils.get_column_letter(c)}{r} пуст"
            if isinstance(v, str) and v.startswith("="):
                continue  # формула — допустима (столбец U)
            if not isinstance(v, (int, float)):
                return False, f"{sheet}!{openpyxl.utils.get_column_letter(c)}{r} не число: {v!r}"
            if sh_d.cell(r, c).value is None:
                return False, f"{sheet}!{openpyxl.utils.get_column_letter(c)}{r} нет кэша значения"
    return True, ""


# ---------------------------------------------------------------------------
# Проверка ссылок диаграмм для произвольной книги (V1/V2-блок)
# ---------------------------------------------------------------------------
def check_series_sheets(path, label=""):
    """V7: для КАЖДОЙ серии каждой диаграммы xVal и yVal на ОДНОМ листе.

    Excel-правило ScatterChart: если xVal/yVal серии лежат на разных листах,
    Excel рендерит категории 1..n вместо реальных X (баг пользователя
    «у всех графиков X разный (шаг сетки)»). Ловит регрессию.
    """
    prefix = label + " " if label else ""
    wb = openpyxl.load_workbook(path, data_only=False)
    gch = wb["Графики"]
    bad = []
    for ci, chart in enumerate(gch._charts, 1):
        for si, series in enumerate(chart.series, 1):
            xr = series_ref(series, "x")
            yr = series_ref(series, "y")
            if xr is None or yr is None:
                bad.append(f"chart{ci}/s{si}: нет xVal/yVal ссылки")
                continue
            if xr["sheet"] != yr["sheet"]:
                bad.append(f"chart{ci}/s{si}: xVal@{xr['sheet']} != yVal@{yr['sheet']}")
    check(f"{prefix}V7: xVal и yVal каждой серии на ОДНОМ листе",
          not bad, "; ".join(bad) if bad else "все серии консистентны")


def check_all_series_programmatic(path, label=""):
    """V8 (итерация 4): проверка ВСЕХ диаграмм/серий листа «Графики» программно.

    Требования контракта (v16-CALC, живые диаграммы):
      * число диаграмм на «Графики» >= 3 (калибровка / профиль U(x) / остатки);
      * КАЖДАЯ диаграмма — ScatterChart (иначе Excel рендерит категории 1..n);
      * у КАЖДОЙ серии xVal и yVal лежат на ОДНОМ листе (Excel-правило);
      * ссылки xVal/yVal ведут на ЧИСЛОВЫЕ диапазоны (не текст/пусто;
        формулы допустимы, напр. столбец H профиля =G*E).
    Читает книгу openpyxl и инспектирует chart.series -> series.xVal/series.yVal.
    """
    prefix = label + " " if label else ""
    wb = openpyxl.load_workbook(path, data_only=False)
    gch = wb["Графики"]
    n_charts = len(gch._charts)
    check(f"{prefix}V8a: на «Графики» не менее 3 диаграмм", n_charts >= 3,
          f"найдено {n_charts}")
    types = [type(c).__name__ for c in gch._charts]
    check(f"{prefix}V8b: все диаграммы — ScatterChart",
          all(t == "ScatterChart" for t in types), f"{types}")
    if n_charts == 0:
        return
    n_series = 0
    bad_sheet = []
    bad_num = []
    for ci, chart in enumerate(gch._charts, 1):
        for si, series in enumerate(chart.series, 1):
            n_series += 1
            xr = series_ref(series, "x")
            yr = series_ref(series, "y")
            if xr is None or yr is None:
                bad_num.append(f"chart{ci}/s{si}: нет xVal/yVal ссылки")
                continue
            if xr["sheet"] != yr["sheet"]:
                bad_sheet.append(f"chart{ci}/s{si}: "
                                 f"xVal@{xr['sheet']}!{xr['min_col']}:{xr['min_row']} "
                                 f"!= yVal@{yr['sheet']}!{yr['min_col']}:{yr['min_row']}")
            for which, rng in (("x", xr), ("y", yr)):
                ok, det = cells_numeric(path, rng["sheet"], rng["min_col"],
                                        rng["max_col"], rng["min_row"], rng["max_row"])
                if not ok:
                    bad_num.append(f"chart{ci}/s{si} {which}Val: {det}")
    check(f"{prefix}V8c: xVal и yVal каждой серии на ОДНОМ листе "
          f"(все {n_series} серий программно)",
          not bad_sheet, "; ".join(bad_sheet) if bad_sheet else f"проверено серий {n_series}")
    check(f"{prefix}V8d: xVal/yVal ссылаются на числовые диапазоны "
          f"(все {n_series} серий программно)",
          not bad_num, "; ".join(bad_num) if bad_num else f"проверено серий {n_series}")


def check_chart_refs(path, n_levels, label=""):
    """V1/V2: лист «Графики», 3 диаграммы, ссылки серий корректны (динамические строки)."""
    prefix = label + " " if label else ""
    wb = openpyxl.load_workbook(path, data_only=False)
    check(f"{prefix}V1a: лист «Графики» существует", "Графики" in wb.sheetnames)
    gch = wb["Графики"]
    n_charts = len(gch._charts)
    check(f"{prefix}V1b: на «Графики» ровно 3 диаграммы", n_charts == 3, f"найдено {n_charts}")
    if n_charts != 3:
        return
    check_series_sheets(path, label=label)  # V7: xVal/yVal серии — один лист
    ws_reg = wb["02_Регрессия"]
    data_hdr = bld.find_hdr_row(ws_reg, [(1, "X"), (2, "Y_ср")])
    check(f"{prefix}V2a: заголовок таблицы точек найден динамически",
          data_hdr is not None, f"DATA_HDR_ROW={data_hdr}")
    DATA_FIRST = data_hdr + 1
    DATA_LAST = DATA_FIRST + n_levels - 1
    GRID_PT = DATA_LAST + 2
    ws_prof = wb["04_Профиль U(X)"]
    prof_hdr = bld.find_hdr_row(ws_prof, [(1, "№"), (2, "X (г/г)")])
    check(f"{prefix}V2b: заголовок таблицы профиля найден динамически",
          prof_hdr is not None, f"PROF_HDR_ROW={prof_hdr}")
    PROF_FIRST = prof_hdr + 1
    PROF_LAST = PROF_FIRST + 199

    ch1, ch2, ch3 = gch._charts[0], gch._charts[1], gch._charts[2]
    # --- Диаграмма 1 ---
    ok1 = len(ch1.series) == 4
    check(f"{prefix}V2c: Диаграмма 1 — 4 серии", ok1, f"серий={len(ch1.series)}")
    if ok1:
        v0, x0_ = series_ref(ch1.series[0], "y"), series_ref(ch1.series[0], "x")
        v1, x1_ = series_ref(ch1.series[1], "y"), series_ref(ch1.series[1], "x")
        v2, x2_ = series_ref(ch1.series[2], "y"), series_ref(ch1.series[2], "x")
        v3, x3_ = series_ref(ch1.series[3], "y"), series_ref(ch1.series[3], "x")
        e = (
            all(r is not None for r in (v0, x0_, v1, x1_, v2, x2_, v3, x3_))
            and v0["sheet"] == "02_Регрессия" and v0["min_col"] == 2
            and v0["min_row"] == DATA_FIRST and v0["max_row"] == DATA_LAST
            and x0_["min_col"] == 1 and x0_["min_row"] == DATA_FIRST and x0_["max_row"] == DATA_LAST
            and v1["min_col"] == 3 and v1["min_row"] == GRID_PT and v1["max_row"] == GRID_PT + 99
            and x1_["min_col"] == 1 and x1_["min_row"] == GRID_PT and x1_["max_row"] == GRID_PT + 99
            and v2["min_col"] == 6 and v2["min_row"] == DATA_FIRST and v2["max_row"] == DATA_LAST
            and x2_["min_col"] == 1 and x2_["min_row"] == DATA_FIRST and x2_["max_row"] == DATA_LAST
            and v3["min_col"] == 5 and v3["min_row"] == DATA_FIRST and v3["max_row"] == DATA_LAST
            and x3_["min_col"] == 1 and x3_["min_row"] == DATA_FIRST and x3_["max_row"] == DATA_LAST
        )
        check(f"{prefix}V2d: Диаграмма 1 — ссылки на диапазоны калибровки", e)
        ok_n, det = cells_numeric(path, "02_Регрессия", 1, 7, DATA_FIRST, DATA_LAST)
        check(f"{prefix}V2e: ячейки данных калибровки числовые (A..G)", ok_n, det)
        zero_ok = all(ws_reg.cell(r, 7).value == 0 for r in range(DATA_FIRST, DATA_LAST + 1))
        check(f"{prefix}V2n: столбец G (линия нуля) = 0 для всех уровней", zero_ok,
              f"G{DATA_FIRST}:G{DATA_LAST}")
        ok_g1, det_g1 = cells_numeric(path, "02_Регрессия", 1, 1, GRID_PT, GRID_PT + 99)
        ok_g3, det_g3 = cells_numeric(path, "02_Регрессия", 3, 3, GRID_PT, GRID_PT + 99)
        check(f"{prefix}V2f: ячейки линии GRID_PT числовые (A и C)",
              ok_g1 and ok_g3, (det_g1 or "") + (det_g3 or ""))

    # --- Диаграмма 2 ---
    ok2 = len(ch2.series) == 3
    check(f"{prefix}V2g: Диаграмма 2 — 3 серии (u_c/U/маркер x_pred)", ok2,
          f"серий={len(ch2.series)}")
    if ok2:
        v0, x0_ = series_ref(ch2.series[0], "y"), series_ref(ch2.series[0], "x")
        v1, x1_ = series_ref(ch2.series[1], "y"), series_ref(ch2.series[1], "x")
        v2, x2_ = series_ref(ch2.series[2], "y"), series_ref(ch2.series[2], "x")
        e = (
            all(r is not None for r in (v0, x0_, v1, x1_, v2, x2_))
            and v0["sheet"] == "04_Профиль U(X)" and v0["min_col"] == 5
            and v0["min_row"] == PROF_FIRST and v0["max_row"] == PROF_LAST
            and x0_["min_col"] == 2 and x0_["min_row"] == PROF_FIRST and x0_["max_row"] == PROF_LAST
            and v1["min_col"] == 8 and v1["min_row"] == PROF_FIRST and v1["max_row"] == PROF_LAST
            and x1_["min_col"] == 2 and x1_["min_row"] == PROF_FIRST and x1_["max_row"] == PROF_LAST
            # маркер x_pred: xVal/yVal на листе ПРОФИЛЯ (K4/L4) — Excel-правило
            and v2["sheet"] == "04_Профиль U(X)" and v2["min_col"] == 12
            and v2["max_col"] == 12 and v2["min_row"] == 4 and v2["max_row"] == 4
            and x2_["sheet"] == "04_Профиль U(X)" and x2_["min_col"] == 11
            and x2_["max_col"] == 11 and x2_["min_row"] == 4 and x2_["max_row"] == 4
        )
        check(f"{prefix}V2h: Диаграмма 2 — ссылки на профиль + маркер x_pred", e)
        ok_n, det = cells_numeric(path, "04_Профиль U(X)", 2, 7, PROF_FIRST, PROF_LAST)
        check(f"{prefix}V2i: ячейки профиля B..G числовые", ok_n, det)
        # V2j (TD-DEV-207): паттерн формулы =G{r}*E{r} через re
        h_ok = True
        h_bad = ""
        for i in range(200):
            r = PROF_FIRST + i
            v = ws_prof.cell(r, 8).value
            if not (isinstance(v, str)
                    and re.fullmatch(rf"=G{r}\*E{r}", v) is not None):
                h_ok = False
                h_bad = f"{v!r} в H{r}"
                break
        check(f"{prefix}V2j: столбец H профиля — живые формулы =G{{r}}*E{{r}} (re)",
              h_ok, h_bad)

    # --- Диаграмма 3 ---
    ok3 = len(ch3.series) == 2
    check(f"{prefix}V2l: Диаграмма 3 — 2 серии (остатки/линия 0)", ok3,
          f"серий={len(ch3.series)}")
    if ok3:
        v0, x0_ = series_ref(ch3.series[0], "y"), series_ref(ch3.series[0], "x")
        v1, x1_ = series_ref(ch3.series[1], "y"), series_ref(ch3.series[1], "x")
        e = (
            all(r is not None for r in (v0, x0_, v1, x1_))
            and v0["sheet"] == "02_Регрессия" and v0["min_col"] == 4
            and v0["min_row"] == DATA_FIRST and v0["max_row"] == DATA_LAST
            and x0_["min_col"] == 1 and x0_["min_row"] == DATA_FIRST and x0_["max_row"] == DATA_LAST
            # линия нуля: yVal на 02_Регрессия (столбец G) ВМЕСТЕ с xVal (столбец A)
            and v1["sheet"] == "02_Регрессия" and v1["min_col"] == 7 and v1["max_col"] == 7
            and v1["min_row"] == DATA_FIRST and v1["max_row"] == DATA_LAST
            and x1_["min_col"] == 1 and x1_["min_row"] == DATA_FIRST and x1_["max_row"] == DATA_LAST
        )
        check(f"{prefix}V2m: Диаграмма 3 — ссылки на остатки и линию 0", e)
    return wb, ws_reg, ws_prof, gch, (DATA_FIRST, DATA_LAST, GRID_PT,
                                      PROF_FIRST, PROF_LAST)


# ---------------------------------------------------------------------------
# Эталонные числа: из САМОЙ КНИГИ (01_Ввод — единый источник, P3)
# ---------------------------------------------------------------------------
def kernel_from_book(path):
    wb = openpyxl.load_workbook(path, data_only=False)
    ws_in = wb["01_Ввод"]
    levels = bld.read_levels(ws_in)
    inputs = bld.read_inputs(ws_in)
    assert levels is not None and len(levels) >= 3, "01_Ввод: уровни не прочитаны"
    res = bld.compute_calibration(levels, inputs, "step")  # v17: step по умолчанию (TD-DEV-208)
    return levels, inputs, res


def synthetic_levels(n, seed=7):
    """Синтетический набор уровней (другой набор данных для V5-теста)."""
    rng = np.random.default_rng(seed)
    xs = np.linspace(0.06, 2.8, n)
    a, b, sr = 0.95, 0.03, 0.012
    levels = []
    for x in xs:
        y1 = a * x + b + rng.normal(0.0, sr)
        y2 = a * x + b + rng.normal(0.0, sr)
        levels.append((float(x), [float(y1), float(y2)]))
    return levels


# ---------------------------------------------------------------------------
print("=" * 70)
print("Проверка EN13132_v17_CALC.xlsx (v17-final: TD-DEV-206/207/208, step по умолчанию)")
print("=" * 70)

levels, inputs, res = kernel_from_book(OUT)
SNAP = res["SNAP"]
print(f"Эталон (из 01_Ввод книги): уровней={SNAP['n']}, инъекций={len(levels[0][1])}, "
      f"p пробы={res['p']}, интерполяция='{res['interpolation']}'")

# ---------------------------------------------------------------------------
# V1/V2: диаграммы (главная книга)
# ---------------------------------------------------------------------------
wb, ws_reg, ws_prof, gch, (DATA_FIRST, DATA_LAST, GRID_PT, PROF_FIRST, PROF_LAST) = \
    check_chart_refs(OUT, len(levels), label="[6уровней]")
check_all_series_programmatic(OUT, label="[6уровней]")  # V8 (итерация 4)

# ---------------------------------------------------------------------------
# V3: числа книги vs ядро
# ---------------------------------------------------------------------------
ws_anova = wb["03_ANOVA"]
r_sr = bld.find_label_row(ws_anova, "Повторяемость инъекций")
r_sl = bld.find_label_row(ws_anova, "Межлабораторная составляющая")
check("V3a: строки σr/σL в 03_ANOVA найдены", r_sr is not None and r_sl is not None)
if r_sr and r_sl:
    check("V3b: σr совпадает с ядром",
          np.isclose(ws_anova.cell(r_sr, 3).value, res["SIGMA_R"], rtol=1e-9, atol=1e-12),
          f"книга={ws_anova.cell(r_sr, 3).value}, ядро={res['SIGMA_R']}")
    check("V3c: σL совпадает с ядром",
          np.isclose(ws_anova.cell(r_sl, 3).value, res["sigma_L"], rtol=1e-9, atol=1e-12),
          f"книга={ws_anova.cell(r_sl, 3).value}, ядро={res['sigma_L']}")

v_a, v_b = ws_reg.cell(4, 2).value, ws_reg.cell(5, 2).value
v_r2, v_s2 = ws_reg.cell(9, 2).value, ws_reg.cell(10, 2).value
check("V3d: a совпадает с ядром", np.isclose(v_a, SNAP["a"], rtol=1e-9, atol=1e-12),
      f"книга={v_a}, ядро={SNAP['a']}")
check("V3e: b совпадает с ядром", np.isclose(v_b, SNAP["b"], rtol=1e-9, atol=1e-12),
      f"книга={v_b}, ядро={SNAP['b']}")
check("V3f: R² совпадает с ядром", np.isclose(v_r2, SNAP["r2"], rtol=1e-9, atol=1e-12),
      f"книга={v_r2}, ядро={SNAP['r2']}")
check("V3g: s²_res совпадает с ядром", np.isclose(v_s2, SNAP["s2_res"], rtol=1e-9, atol=1e-12),
      f"книга={v_s2}, ядро={SNAP['s2_res']}")

ws_bud = wb["05_Бюджет"]
r_uc = bld.find_label_row(ws_bud, "u_c (г/г)")
r_ne = bld.find_label_row(ws_bud, "ν_eff (Welch")
r_k = bld.find_label_row(ws_bud, "k (t_")
r_u = bld.find_label_row(ws_bud, "U (г/г) = k·u_c")
check("V3h: строки бюджета найдены",
      all(x is not None for x in (r_uc, r_ne, r_k, r_u)))
if all(x is not None for x in (r_uc, r_ne, r_k, r_u)):
    check("V3i: u_c совпадает с ядром",
          np.isclose(ws_bud.cell(r_uc, 2).value, res["U_C_RES"], rtol=1e-9, atol=1e-12),
          f"книга={ws_bud.cell(r_uc, 2).value}, ядро={res['U_C_RES']}")
    check("V3j: ν_eff совпадает с ядром",
          np.isclose(ws_bud.cell(r_ne, 2).value, res["NU_EFF_RES"], rtol=1e-9, atol=1e-12),
          f"книга={ws_bud.cell(r_ne, 2).value}, ядро={res['NU_EFF_RES']}")
    check("V3k: k совпадает с ядром",
          np.isclose(ws_bud.cell(r_k, 2).value, res["K_RES"], rtol=1e-9, atol=1e-12),
          f"книга={ws_bud.cell(r_k, 2).value}, ядро={res['K_RES']}")
    check("V3l: U (г/г) совпадает с ядром",
          np.isclose(ws_bud.cell(r_u, 2).value, res["U_RES_G"], rtol=1e-9, atol=1e-12),
          f"книга={ws_bud.cell(r_u, 2).value}, ядро={res['U_RES_G']}")
    check("V3m: U (% мас.) совпадает с ядром",
          np.isclose(ws_bud.cell(r_u + 1, 2).value, res["U_RES_PCT"], rtol=1e-9, atol=1e-12),
          f"книга={ws_bud.cell(r_u + 1, 2).value}, ядро={res['U_RES_PCT']}")
    # TD-DEV-206: строка «прецизионность» 05_Бюджет == u_repr(w0)/K_G2PCT (smooth)
    r_pr = bld.find_label_row(ws_bud, "прецизионность u_repr (mean2)")
    exp_pr = res["u_repr_w0"] / res["K_G2PCT"]
    check("V3m2: прецизионность 05_Бюджет = u_repr(w0)/K (TD-DEV-206)",
          r_pr is not None and np.isclose(ws_bud.cell(r_pr, 2).value, exp_pr,
                                          rtol=1e-9, atol=1e-12),
          f"книга={ws_bud.cell(r_pr, 2).value if r_pr else None}, ядро={exp_pr}")

ws_it = wb["06_Итог"]
r_x0 = bld.find_label_row(ws_it, "X0 (обратная градуировка)")
r_w = bld.find_label_row(ws_it, "w(MTBE)")
check("V3n: строки X0/w в 06_Итог найдены", r_x0 is not None and r_w is not None)
if r_x0 and r_w:
    check("V3o: x0 совпадает с ядром",
          np.isclose(ws_it.cell(r_x0, 2).value, res["X0"], rtol=1e-9, atol=1e-12),
          f"книга={ws_it.cell(r_x0, 2).value}, ядро={res['X0']}")
    check("V3p: w (% мас.) совпадает с ядром",
          np.isclose(ws_it.cell(r_w, 2).value, res["W0"], rtol=1e-9, atol=1e-12),
          f"книга={ws_it.cell(r_w, 2).value}, ядро={res['W0']}")

# V3q (TD-DEV-207): ВСЕ 200 точек профиля U = k·u_c vs ядро + k/u_c/прецизионность
u_mism = 0
u_checked = 0
k_mism = 0
pr_mism = 0
for i in range(200):
    r = PROF_FIRST + i
    k_cell = ws_prof.cell(r, 7).value
    uc_cell = ws_prof.cell(r, 5).value
    pr_cell = ws_prof.cell(r, 4).value
    if not (isinstance(k_cell, (int, float)) and isinstance(uc_cell, (int, float))):
        u_mism += 1
        continue
    u_book = k_cell * uc_cell
    if not np.isclose(u_book, res["U_X"][i], rtol=1e-6, atol=1e-12):
        u_mism += 1
    if not np.isclose(k_cell, res["K_ARR"][i], rtol=1e-6, atol=1e-12):
        k_mism += 1
    if not np.isclose(pr_cell, res["U_REPR_X"][i], rtol=1e-9, atol=1e-12):
        pr_mism += 1
    u_checked += 1
check("V3q: U(x) профиля = k·u_c совпадает с ядром во ВСЕХ 200 точках",
      u_mism == 0 and u_checked == 200, f"проверено {u_checked}, расхождений {u_mism}")
check("V3q2: k(x) совпадает с ядром (200 точек)", k_mism == 0, f"расхождений {k_mism}")
check("V3q3: u_repr(x) столбец D совпадает с ядром (200 точек, step)",
      pr_mism == 0, f"расхождений {pr_mism}")

# V3r (TD-DEV-206): U(06_Итог) == U(x_pred) маркер «Графики»!D4
r_u_it = bld.find_label_row(ws_it, "U, г/г")
u_itog = ws_it.cell(r_u_it, 2).value if r_u_it else None
u_marker = gch["D4"].value
ok206 = (
    u_itog is not None and isinstance(u_marker, (int, float))
    and np.isclose(u_itog, u_marker, rtol=1e-9, atol=1e-12)
    and np.isclose(u_itog, res["U_RES_G"], rtol=1e-9, atol=1e-12)
    and np.isclose(u_marker, res["U_RES_G"], rtol=1e-9, atol=1e-12)
)
check("V3r: U(06_Итог) == U(x_pred) маркер Графики!D4 == ядро (TD-DEV-206)", ok206,
      f"06_Итог U={u_itog}, D4={u_marker}, ядро={res['U_RES_G']}")
# согласованность маркера с КРИВОЙ профиля (в пределах шага сетки):
u_curve_xpred = res["U_X"][res["x_pred_idx"]]
rel = abs(u_curve_xpred - res["U_RES_G"]) / res["U_RES_G"]
check("V3r2: маркер D4 лежит на кривой профиля U(x_pred) в пределах шага сетки",
      rel < 1e-3, f"U_кривая(x_pred)={u_curve_xpred}, U_итог={res['U_RES_G']}, rel={rel:.2e}")

# V3s: маркер x_pred C4 == X0
check("V3s: маркер x_pred C4 = X0 (ядро)",
      np.isclose(gch["C4"].value, res["X0"], rtol=1e-9, atol=1e-12),
      f"C4={gch['C4'].value}, ядро X0={res['X0']}")

# ---------------------------------------------------------------------------
# V4: встроенных диаграмм в 02/04 нет
# ---------------------------------------------------------------------------
check("V4a: в 02_Регрессия нет встроенных диаграмм", len(ws_reg._charts) == 0,
      f"найдено {len(ws_reg._charts)}")
check("V4b: в 04_Профиль U(X) нет встроенных диаграмм", len(ws_prof._charts) == 0,
      f"найдено {len(ws_prof._charts)}")

# ---------------------------------------------------------------------------
# V6 (TD-DEV-208): профиль СТУПЕНЧАТЫЙ (step), режим указан в шапке
# ---------------------------------------------------------------------------
# u_repr(x) в step-режиме постоянна ВНУТРИ поддиапазона табл.1 и меняется
# только при переходе границы поддиапазона (в % мас.). Проверяем на сетке:
# внутри каждого сегмента сетки, целиком лежащего в одном поддиапазоне,
# |Δu_repr| == 0; число скачков ограничено числом границ поддиапазонов в
# диапазоне калибровки (для w от ~0.5 до ~3.5 %мас это 3..4 границы).
ranges = ag.TABLE1["ranges"]
w_grid = res["X_GRID"] * res["K_G2PCT"]  # % мас.
urepr_arr = np.asarray(res["U_REPR_X"], dtype=float)
# поддиапазон для каждой точки: нижняя граница lo (step-режим — константа внутри)
lo_arr = np.array([
    float(ag.u_repr_from_tables(float(w), ag.TABLE1, "mean2", "step")["lo"])
    for w in w_grid
], dtype=float)
# границы смены поддиапазона
d_lo = np.diff(lo_arr)
boundaries = np.where(np.abs(d_lo) > 1e-9)[0]
# внутри непрерывных сегментов u_repr должна быть постоянной
flat_ok = True
flat_bad = ""
for start, end in zip(np.concatenate(([0], boundaries + 1)),
                      np.concatenate((boundaries, [len(urepr_arr) - 1]))):
    vals = urepr_arr[start:end + 1]
    if len(vals) > 1 and not np.allclose(vals, vals[0], rtol=0, atol=1e-12):
        flat_ok = False
        flat_bad = f"сегмент {start}..{end}: {vals[0]:.7f}..{vals[-1]:.7f}"
        break
check("V6a: u_repr(x) СТУПЕНЧАТАЯ — постоянна внутри поддиапазонов табл.1 (step)",
      flat_ok, flat_bad if flat_bad else f"сегментов-поддиапазонов на сетке: {len(boundaries) + 1}")
prof_a2 = ws_prof["A2"].value or ""
check("V6b: в шапке 04_Профиль указан режим 'step' и формула mean2",
      "step" in prof_a2 and "√(R²−r²/2)/2,8" in prof_a2,
      f"шапка: {prof_a2[:80]!r}...")
# согласование с ядром: u_repr(w0) по шагу
urepr_w0_check = ag.u_repr_from_tables(res["W0"], ag.TABLE1, "mean2", "step")["u_repr"]
check("V6c: u_repr(w0) ядра step == ожидаемому 0.1406 %мас (TD-DEV-208)",
      np.isclose(urepr_w0_check, 0.140607, rtol=1e-3, atol=1e-5),
      f"u_repr(w0={res['W0']:.3f}) = {urepr_w0_check:.6f} %мас")

# ---------------------------------------------------------------------------
# V5 (P3, TD-DEV-207): генерация при 4/5/8 уровнях (другой набор данных)
# ---------------------------------------------------------------------------
for n_lev in (4, 5, 8):
    try:
        lv = synthetic_levels(n_lev, seed=10 + n_lev)
        inp = bld.default_inputs(lv)
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
            tmp = tf.name
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", RuntimeWarning)
            bld.build_workbook(lv, inp, tmp, interpolation="step")
        ok_build = os.path.exists(tmp)
        check(f"V5a[{n_lev}уровней]: генерация не падает", ok_build, f"файл {tmp}")
        if ok_build:
            # ссылки диаграмм корректны при динамическом числе уровней
            check_chart_refs(tmp, n_lev, label=f"[{n_lev}уровней]")
            check_all_series_programmatic(tmp, label=f"[{n_lev}уровней]")  # V8
            # согласованность U(Итог)==D4 и для пересобранной книги
            _lv2, _inp2, res2 = kernel_from_book(tmp)
            _wb2 = openpyxl.load_workbook(tmp, data_only=False)
            _it2 = _wb2["06_Итог"]
            _r_u2 = bld.find_label_row(_it2, "U, г/г")
            _u2 = _it2.cell(_r_u2, 2).value
            _d4 = _wb2["Графики"]["D4"].value
            check(f"V5b[{n_lev}уровней]: U(06_Итог)==D4 в пересобранной книге",
                  np.isclose(_u2, _d4, rtol=1e-9, atol=1e-12)
                  and np.isclose(_u2, res2["U_RES_G"], rtol=1e-9, atol=1e-12),
                  f"06_Итог={_u2}, D4={_d4}")
        try:
            os.remove(tmp)
        except OSError:
            pass
    except Exception as exc:  # noqa: BLE001
        check(f"V5a[{n_lev}уровней]: генерация не падает", False, f"исключение: {exc}")

# ---------------------------------------------------------------------------
print("=" * 70)
n_pass = sum(1 for _n, ok, _d in _results if ok)
n_fail = sum(1 for _n, ok, _d in _results if not ok)
print(f"ИТОГ: PASS={n_pass}, FAIL={n_fail}")
sys.exit(0 if n_fail == 0 else 1)