#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_en13132_v17_word.py — Word-отчёт v17 (блок `gost13132-v17-final`).

Генерирует EN13132_v17_Отчёт.docx по структуре эталона v15
(EN13132_v15_Отчёт_МТБЭ.docx: заголовок, идентификация, уравнение/LoF,
таблица уровней, бюджет, подписи) + расширения v16/v17:
  * параметры калибровки (n, a, b, R², s²_res, LoF);
  * таблица ANOVA (источник/SS/df/MS/F) + компоненты σr/σR/σL/u_repr;
  * полный бюджет неопределённости (вклад/u/ν/доля%) + u_c, ν_eff, k, U;
  * протокольная запись результата w±U (P=0,95);
  * профиль U(x): текст + таблица ключевых точек (маркер x_pred);
  * примечания (табл.1 ГОСТ EN 13132, 5725-6, GUM H.5.3.3, TD-DEV-33/208);
  * подписи Выполнил/Проверил/Утвердил.

Отличия v17 от v16 (TD-DEV-208, вердикт ресерчера по НД):
  * u_repr — СТУПЕНЧАТАЯ по табл.1 ГОСТ EN 13132 (step, ПО УМОЛЧАНИЮ):
    r/R ДИСКРЕТНЫ, интерполяция не предписана ни стандартом, ни 5725-6,
    ни РМГ 61; GUM 4.1.4 к приписанным нормам прецизионности неприменим.
    smooth НЕ используется (методическое допущение, не соответствует НД).
  * Вклады-поля ввода (эталоны/чистота/весы): значение из спецификации ИЛИ
    0 = «не учитывать» (исключается из u_c и ν_eff, TD-DEV-33).
  * Статус «ПРЕДВАРИТЕЛЬНО» снят: значения вкладов — из спецификации
    пользователя; результат НЕ блокируется.

ВСЕ числа — из ядра anova_gum.py (через compute_calibration модуля
build_en13132_v17_calc, interpolation="step") и входов 01_Ввод книги
EN13132_v17_CALC.xlsx. Никаких хардкодов чисел в текст — только константы-
эталоны REFS для самопроверки при генерации (rtol=1e-3).

Правила: кириллица UTF-8; python-docx; стили v15 (Arial 8.5, заголовок 13pt
bold 214868, шапки таблиц 315A7C белым, зебра F0F4F8, Table Grid).
"""
from __future__ import annotations

import math
import os
import sys

import numpy as np
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import build_en13132_v17_calc as bld  # noqa: E402  (ядро + чтение входов)

V17_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DOCX = os.path.join(V17_DIR, "EN13132_v17_Отчёт.docx")

# Эталоны для самопроверки (rtol=1e-3). Числа НЕ подставляются в текст отчёта.
# step-режим (TD-DEV-208): u_repr по табл.1 для w0=3.967 %мас -> (3;5]: R=0.4, r=0.1.
REFS = {
    "a": 0.960998, "b": 0.036203, "r2": 0.999740, "s2_res": 4.788e-4, "df_res": 4,
    "F_lof": 2.2575, "F_crit": 4.5337,
    "sigma_r": 0.020596, "sigma_R_tab": 0.142857, "sigma_L": 0.139091,
    "u_repr": 0.140607,           # % мас., mean2, step (w0=3.967 -> (3;5]: R=0.4,r=0.1)
    "u_c": 0.040164, "nu_eff": 86.27, "k": 1.9878, "U": 0.079841,
    "U_pct": 0.3178, "U_rel": 8.01, "w0": 3.967,
}

# ---------------------------------------------------------------------------
# Стили (паттерн v15 make_en13132_v15_reports.py)
# ---------------------------------------------------------------------------
C_TITLE = RGBColor(0x21, 0x48, 0x68)
C_TEXT = RGBColor(0x2C, 0x39, 0x44)
C_WARN = RGBColor(0x96, 0x3F, 0x2F)
C_HDR_FILL = "315A7C"
C_ZEBRA = "F0F4F8"


def fnum(v, n=5):
    """Формат числа как в v15: запятая-разделитель, 'н/д' для None/NaN."""
    if v is None:
        return "н/д"
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "н/д"
    if not np.isfinite(v):
        return "н/д"
    return f"{v:.{n}f}".replace(".", ",")


def fnum_sci(v, n=3):
    """Научный формат с запятой в мантиссе (для малых u/s2)."""
    if v is None:
        return "н/д"
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "н/д"
    if not np.isfinite(v):
        return "н/д"
    s = f"{v:.{n}e}"
    mant, exp = s.split("e")
    return mant.replace(".", ",") + "e" + str(int(exp))


def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def style_cell_text(cell, size=8, bold=False, color=None):
    for para in cell.paragraphs:
        para.paragraph_format.space_after = Pt(0)
        for run in para.runs:
            run.font.size = Pt(size)
            run.font.bold = bold
            if color is not None:
                run.font.color.rgb = color


def make_table(doc, headers, rows, widths=None, zebra=True, hdr_size=8, body_size=8):
    """Таблица Table Grid с шапкой 315A7C (белый текст) и зеброй."""
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for j, h in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.text = str(h)
        shade(cell, C_HDR_FILL)
        style_cell_text(cell, size=hdr_size, bold=True, color=RGBColor(255, 255, 255))
    for i, row in enumerate(rows):
        cells = table.add_row().cells
        for j, v in enumerate(row):
            cells[j].text = "" if v is None else str(v)
            if zebra and i % 2 == 1:
                shade(cells[j], C_ZEBRA)
            style_cell_text(cells[j], size=body_size)
    if widths:
        for j, w in enumerate(widths):
            for row in table.rows:
                row.cells[j].width = Cm(w)
    return table


def para(doc, text="", bold=False, italic=False, size=None, color=None,
         align=None, space_after=3, space_before=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    if size is not None:
        run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    return p


def add_runs(p, parts):
    """parts: список (текст, bold, italic, size, color)."""
    for text, bold, italic, size, color in parts:
        r = p.add_run(text)
        r.bold = bold
        r.italic = italic
        if size is not None:
            r.font.size = Pt(size)
        if color is not None:
            r.font.color.rgb = color
    return p


# ---------------------------------------------------------------------------
# Данные: один источник — ядро + 01_Ввод книги CALC
# ---------------------------------------------------------------------------
def oneway_anova(levels):
    """Однофакторная ANOVA по повторностям уровней (источник/SS/df/MS/F).

    SS_between = Σ n_i·(Ȳ_i−Ȳ)²; SS_within = ΣΣ(Y_ij−Ȳ_i)².
    F — информационная (между уровнями = систематика калибровки, НЕ случайный
    эффект; не входит в бюджет неопределённости).
    """
    means = [float(np.mean(ys)) for _x, ys in levels]
    ns = [len(list(ys)) for _x, ys in levels]
    N = sum(ns)
    Ybar = float(np.average(means, weights=ns))
    SS_b = sum(ni * (mi - Ybar) ** 2 for ni, mi in zip(ns, means))
    SS_w = sum(float(np.sum((np.asarray(list(ys), dtype=float) - mi) ** 2))
               for (_x, ys), mi in zip(levels, means))
    df_b = len(levels) - 1
    df_w = N - len(levels)
    df_t = N - 1
    SS_t = SS_b + SS_w
    MS_b = SS_b / df_b if df_b > 0 else float("nan")
    MS_w = SS_w / df_w if df_w > 0 else float("nan")
    F = MS_b / MS_w if (MS_w > 0 and MS_b == MS_b) else float("nan")
    return {
        "SS_b": SS_b, "df_b": df_b, "MS_b": MS_b,
        "SS_w": SS_w, "df_w": df_w, "MS_w": MS_w,
        "SS_t": SS_t, "df_t": df_t, "F": F, "N": N,
    }


def collect_data():
    """Чтение входов из CALC-книги + расчёт ядром (step, TD-DEV-208). Возвращает dict чисел."""
    levels, inputs = bld.load_existing_inputs(bld.OUT_XLSX)
    if levels is None:
        raise SystemExit(
            "build_en13132_v17_word: не найдены входы в "
            f"{bld.OUT_XLSX} — сначала соберите CALC-книгу (build_en13132_v17_calc.py)"
        )
    res = bld.compute_calibration([(x, ys) for x, ys, *_ in levels], inputs, "step")
    SNAP = res["SNAP"]
    anova = oneway_anova([(x, ys) for x, ys, *_ in levels])

    # таблица уровней: №, X, Y_ср, Ŷ, e
    cal_rows = []
    for i, (x, ys) in enumerate([(x, ys) for x, ys, *_ in levels], 1):
        y_mean = float(np.mean(ys))
        y_calc = SNAP["a"] * float(x) + SNAP["b"]
        cal_rows.append((i, float(x), y_mean, y_calc, y_mean - y_calc))

    # бюджет: вклад/u/ν/доля%
    items = res["BUDG"]["items"]
    u2_total = sum(u ** 2 for _k, u, _d in items)
    budget_rows = []
    label_map = {
        "модель": "Модель (OLS: 1/p+1/n+(x−x̄)²/Sxx)",
        "эталоны": "Эталоны (u(x_cal))",
        "чистота": "Чистота аналита/ИС (pA, pИС)",
        "весы": "Масса навесок (весы)",
        "прецизионность": "Прецизионность u_repr (mean2)",
    }
    for k, u, df in items:
        name = label_map.get(k, k)
        share = (u ** 2 / u2_total * 100.0) if u2_total > 0 else 0.0
        budget_rows.append((name, u, df, share))

    # профиль: 14 ключевых точек по сетке 200 (равномерно) + маркер x_pred
    grid = res["X_GRID"]
    n_pts = len(grid)
    idx = sorted(set(np.linspace(0, n_pts - 1, 14).astype(int).tolist()))
    prof_rows = []
    for i in idx:
        prof_rows.append({
            "x": float(grid[i]), "uc": float(res["U_C"][i]),
            "nu": float(res["NU_EFF"][i]), "k": float(res["K_ARR"][i]),
            "U": float(res["U_X"][i]), "is_pred": False,
        })
    pred_i = res["x_pred_idx"]
    prof_rows.append({
        "x": float(grid[pred_i]), "uc": float(res["U_C"][pred_i]),
        "nu": float(res["NU_EFF"][pred_i]), "k": float(res["K_ARR"][pred_i]),
        "U": float(res["U_X"][pred_i]), "is_pred": True,
    })
    prof_rows.sort(key=lambda r: r["x"])

    urepr = bld.ag.u_repr_from_tables(res["W0"], bld.ag.TABLE1, "mean2", "step")
    return {
        "res": res, "SNAP": SNAP, "anova": anova, "cal_rows": cal_rows,
        "budget_rows": budget_rows, "prof_rows": prof_rows,
        "u2_total": u2_total, "levels": levels, "inputs": inputs,
        "u_repr_full": urepr, "X_min": float(np.min(grid)), "X_max": float(np.max(grid)),
    }


# ---------------------------------------------------------------------------
# Сборка документа
# ---------------------------------------------------------------------------
def build(d, ref_check=None):
    doc = Document()
    sec = doc.sections[0]
    sec.page_height = Cm(29.7)
    sec.page_width = Cm(21)
    sec.top_margin = Cm(1.28)
    sec.bottom_margin = Cm(1.2)
    sec.left_margin = Cm(1.65)
    sec.right_margin = Cm(1.65)

    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(8.5)
    normal.font.color.rgb = C_TEXT
    normal.paragraph_format.space_after = Pt(3)

    S = d["SNAP"]
    res = d["res"]
    anova = d["anova"]
    w0, X0 = res["W0"], res["X0"]

    # ---- 1. ЗАГОЛОВОК (стиль v15) ----
    p = para(doc, "ОТЧЁТ ОПРЕДЕЛЕНИЯ MTBE ПО ГОСТ EN 13132-2012",
             bold=True, size=13, color=C_TITLE, align=WD_ALIGN_PARAGRAPH.CENTER,
             space_after=3)
    para(doc, "ГОСТ EN 13132-2012  ·  Весовое приготовление в н-гептане  ·  "
              "Диагностическая модель OLS  ·  Классическая ANOVA по GUM (v17)",
         italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=5)

    # ---- 2. ИДЕНТИФИКАЦИЯ (поля-заглушки) ----
    para(doc, "1. Идентификация", bold=True, size=10, space_before=4)
    ident = [
        ("Дата анализа", "____.__.20____"),
        ("Оператор", "____________________"),
        ("Проба (ID)", "____________________"),
        ("Методика", "ГОСТ EN 13132-2012 (MTBE; весовое приготовление в н-гептане)"),
        ("Прибор", "____________________"),
        ("Партия", "____________________"),
    ]
    make_table(doc, ["Поле", "Значение"], ident, widths=(5.0, 11.7), zebra=False)

    # ---- 3. ПАРАМЕТРЫ КАЛИБРОВКИ ----
    para(doc, "2. Параметры калибровки", bold=True, size=10, space_before=6)
    lof_status = res["lof"]["status"] if res["lof"] else "н/д"
    lof_text = (f"F = {fnum(res['lof']['F'], 4)}; Fкр = {fnum(res['lof']['F_crit'], 4)}; "
                f"{lof_status}") if res["lof"] else "LoF не оценим"
    cal_params = [
        ("Число уровней n", f"{S['n']} (6 уровней × 2 инъекции)"),
        ("Диапазон X", f"{fnum(d['X_min'], 6)} … {fnum(d['X_max'], 6)} г/г"),
        ("Наклон a", f"{fnum(S['a'], 6)} Y/X"),
        ("Свободный член b", f"{fnum(S['b'], 7)} Y"),
        ("R²", f"{fnum(S['r2'], 6)}"),
        ("Остаточная дисперсия s²_res", f"{fnum_sci(S['s2_res'], 3)} Y²  (df_res = {S['df_res']})"),
        ("Lack-of-fit", lof_text),
    ]
    make_table(doc, ["Параметр", "Значение"], cal_params, widths=(6.2, 10.5), zebra=False)

    # ---- 4. ТАБЛИЦА КАЛИБРОВКИ ----
    para(doc, "3. Таблица калибровки (Y_ср — среднее инъекций уровня; "
              "Ŷ = a·X + b; e — остаток)", bold=True, size=10, space_before=6)
    cal_tbl = [(i, x, ym, yc, e) for i, x, ym, yc, e in d["cal_rows"]]
    make_table(
        doc,
        ["№", "X (г/г)", "Y_ср", "Ŷ = a·X + b", "остаток e"],
        [[i, fnum(x, 6), fnum(ym, 6), fnum(yc, 6), fnum(e, 6)] for i, x, ym, yc, e in cal_tbl],
        widths=(1.2, 3.4, 3.4, 3.9, 3.8),
    )

    # ---- 5. ТАБЛИЦА ANOVA + КОМПОНЕНТЫ ----
    para(doc, "4. Таблица ANOVA и компоненты прецизионности (GUM H.5, 5725-6 §4.1.4)",
         bold=True, size=10, space_before=6)
    anova_rows = [
        ("Между уровнями (систематика калибровки)", fnum(anova["SS_b"], 6),
         fnum(anova["df_b"], 0), fnum(anova["MS_b"], 6), fnum(anova["F"], 3)),
        ("Внутри уровней (pure error, повторяемость σr)", fnum(anova["SS_w"], 6),
         fnum(anova["df_w"], 0), fnum(anova["MS_w"], 6), "—"),
        ("Итого", fnum(anova["SS_t"], 6), fnum(anova["df_t"], 0), "—", "—"),
    ]
    make_table(
        doc,
        ["Источник", "SS", "df", "MS", "F"],
        anova_rows,
        widths=(6.4, 2.6, 1.5, 2.7, 2.5),
    )
    p = para(doc, "", space_before=3, space_after=3)
    add_runs(p, [
        ("σr = ", True, False, None, None),
        (f"{fnum(res['SIGMA_R'], 6)} (Y, df = {res['ANOVA']['df_r']})", False, False, None, None),
        ("    σR (табл.1) = ", True, False, None, None),
        (f"{fnum(res['sigma_R_tab'], 6)} % мас. (R(w0) = {fnum(res['R_w0'], 4)}; "
         f"σR = R/2,8)", False, False, None, None),
        ("    σL = ", True, False, None, None),
        (f"{fnum(res['sigma_L'], 6)} % мас. (√(σR²−σr²))", False, False, None, None),
        ("    u_repr(mean2) = ", True, False, None, None),
        (f"{fnum(d['u_repr_full']['u_repr'], 6)} % мас. (√(R²−r²/2)/2,8)", False, False, None, None),
    ])

    # ---- 6. БЮДЖЕТ НЕОПРЕДЕЛЁННОСТИ ----
    para(doc, "5. Бюджет неопределённости результата пробы (X0 = "
              f"{fnum(X0, 6)} г/г, w0 = {fnum(w0, 3)} % мас.)",
         bold=True, size=10, space_before=6)
    bud_rows = []
    for name, u, df, share in d["budget_rows"]:
        bud_rows.append((name, fnum(u, 7),
                         "∞" if df is None else fnum(df, 0),
                         fnum(share, 2) + " %"))
    make_table(
        doc,
        ["Вклад", "u (г/г)", "ν", "Доля дисперсии"],
        bud_rows,
        widths=(6.4, 2.8, 1.5, 4.0),
    )
    totals = [
        ("u_c (г/г) = √Σu²", fnum(res["U_C_RES"], 7), "—", "—"),
        ("ν_eff (Welch–Satterthwaite, GUM G.2b)", fnum(res["NU_EFF_RES"], 2), "—", "—"),
        ("k = t_{0.975}(ν_eff), GUM 6.3", fnum(res["K_RES"], 4), "—", "—"),
        ("U = k·u_c, г/г", fnum(res["U_RES_G"], 7), "—", "—"),
        ("U, % мас. (= U(г/г)·K_G2PCT)", fnum(res["U_RES_PCT"], 4), "—", "—"),
        ("U_rel, % (отн.)", fnum(res["U_REL"] * 100.0, 2), "—", "—"),
    ]
    make_table(doc, ["Итог", "Значение", "", ""], totals, widths=(6.4, 2.8, 1.5, 4.0),
               zebra=False)

    # ---- 7. РЕЗУЛЬТАТ ----
    para(doc, "6. Результат (протокольная запись)", bold=True, size=10, space_before=6)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    add_runs(p, [
        ("w(MTBE) = ", True, False, 12, C_TITLE),
        (f"{fnum(w0, 3)} ± {fnum(res['U_RES_PCT'], 2)} % мас. ", True, False, 12, C_TITLE),
        (f"(P = 0,95, k = {fnum(res['K_RES'], 2)}, ν_eff = {fnum(res['NU_EFF_RES'], 0)})",
         False, False, 12, C_TITLE),
    ])
    para(doc, "Протокольное округление до 0,1 % мас. (§9 ГОСТ EN 13132): "
              f"w = {fnum(w0, 1)} ± {fnum(res['U_RES_PCT'], 1)} % мас. "
              "(округление U — GUM 7.2.6, вверх).",
         italic=True, size=7.5)

    # ---- 8. ПРОФИЛЬ НЕОПРЕДЕЛЁННОСТИ ----
    para(doc, "7. Профиль неопределённости U(x) = k(x)·u_c(x)", bold=True, size=10,
         space_before=6)
    para(doc, f"Профиль рассчитан на сетке из {len(res['X_GRID'])} точек "
              f"(шаг {fnum((d['X_max'] - d['X_min']) / (len(res['X_GRID']) - 1), 6)} г/г) "
              f"по диапазону калибровки {fnum(d['X_min'], 6)} … {fnum(d['X_max'], 6)} г/г. "
              f"Модельная часть — интервальная форма OLS u²(x) = s²_res·(1/p + 1/n + "
              f"(x−x̄)²/Sxx)/a², p = {res['p']}; прецизионность u_repr(x) — СТУПЕНЧАТАЯ "
              f"по поддиапазонам табл.1 ГОСТ EN 13132 (step, mean2, √(R²−r²/2)/2,8; "
              f"r/R дискретны — TD-DEV-208); "
              f"U(x) = k(ν_eff(x))·u_c(x). Маркер x_pred = {fnum(X0, 6)} г/г. "
              f"Нижняя граница профиля U(Xmin) = {fnum(res['U_X'][0], 6)} г/г; "
              f"результат U(X0) = {fnum(res['U_X'][res['x_pred_idx']], 6)} г/г; "
              f"верхняя U(Xmax) = {fnum(res['U_X'][-1], 6)} г/г.",
         size=8)
    prof_rows = []
    for r_ in d["prof_rows"]:
        mark = "  ← x_pred" if r_["is_pred"] else ""
        prof_rows.append((fnum(r_["x"], 6), fnum(r_["uc"], 6),
                          fnum(r_["nu"], 1), fnum(r_["k"], 4), fnum(r_["U"], 6) + mark))
    make_table(
        doc,
        ["X (г/г)", "u_c(x) (г/г)", "ν_eff(x)", "k(x)", "U(x) (г/г)"],
        prof_rows,
        widths=(3.2, 3.2, 3.2, 3.0, 4.1),
    )

    # ---- 9. ПРИМЕЧАНИЯ ----
    para(doc, "8. Примечания", bold=True, size=10, space_before=6)
    notes = [
        "· σR принято по табл.1 ГОСТ EN 13132 §10.1 (5725-6 §4.1.2/4.1.4): "
        f"R(w0 = {fnum(w0, 3)} % мас.) = {fnum(res['R_w0'], 4)} % мас., "
        f"σR = R/2,8 = {fnum(res['sigma_R_tab'], 6)} % мас.; u_repr(mean2) = "
        f"√(R²−r²/2)/2,8 = {fnum(d['u_repr_full']['u_repr'], 6)} % мас. (5725-6 §4.2.3). "
        "u_repr — СТУПЕНЧАТАЯ по поддиапазонам табл.1 (step, ПО УМОЛЧАНИЮ): r/R "
        "ДИСКРЕТНЫ, интерполяция не предписана ни стандартом, ни 5725-6, ни РМГ 61; "
        "GUM 4.1.4 к приписанным нормам прецизионности неприменим (TD-DEV-208, вердикт "
        "ресерчера по НД). smooth не используется — методическое допущение, не по НД.",
        f"· σL = √(max(0, σR² − σr²)) = {fnum(res['sigma_L'], 6)} % мас. "
        "(GUM 4.2: отрицательные дисперсии обнуляются; грубый учёт — σr в Y-единицах, "
        "σR в % мас., аппроксимация).",
        "· GUM H.5.3.3: на однодневных данных (1 день, 1 лаборатория, 6 уровней × "
        "2 инъекции) полный иерархический ANOVA невозможен — воспроизводим только "
        "нижний уровень σr; межсерийная σL лаборатории не оцениваема (QC BUD-17, вне "
        "плана); применён «грубый учёт» σR из табл.1/2 ГОСТ EN 13132 (5725-6 §7, "
        "РМГ 61 §5.5).",
        "· Анти-двойной счёт (контракт разд.6, правила 1–3): u_модель включает "
        "усреднённую повторяемость инъекций; σr отдельной строкой бюджета НЕ "
        "добавляется; u_repr входит в форме √(R²−r²/2)/2,8 без задвоения локальной "
        "повторяемости.",
        f"· Вклады-поля ввода (TD-DEV-33): весы, эталоны, чистота — значение из "
        f"спецификации ИЛИ 0 = «не учитывать» (вклад исключается из u_c и ν_eff). "
        f"В текущем прогоне: u_весы = "
        f"{fnum(next((u for k, u, _ in res['BUDG']['items'] if k == 'весы'), 0.0), 7)} г/г.",
        "· Профиль U(x) — СТУПЕНЧАТЫЙ по табл.1 ГОСТ EN 13132 (step): внутри "
        "поддиапазона u_repr постоянна; на границе — более высокий диапазон (R-05, "
        "РМГ 61 §5.5). Smooth-аппроксимация НЕ используется (TD-DEV-208).",
        "· Документируемое допущение (17025 §7.6): «σR принято по табл.1 ГОСТ EN 13132 "
        "(N межлаб. округление 2,8); σL данных лаборатории не оценивалась за "
        "отсутствием межсерийных данных».",
        "· Значения вкладов — из спецификации пользователя; заполнить перед протоколом. "
        "Результат НЕ блокируется (пользователь сам решает, включать вклад или 0).",
    ]
    for n in notes:
        para(doc, n, size=7.5, space_after=2)

    # ---- 10. ПОДПИСИ ----
    para(doc, "9. Подписи", bold=True, size=10, space_before=6)
    sig_rows = [
        ("Выполнил", "____________________", "____.__.20____"),
        ("Проверил", "____________________", "____.__.20____"),
        ("Утвердил", "____________________", "____.__.20____"),
    ]
    make_table(doc, ["Роль", "ФИО", "Дата"], sig_rows, widths=(3.5, 7.0, 5.2), zebra=False)

    doc.save(OUT_DOCX)
    return OUT_DOCX


# ---------------------------------------------------------------------------
# Самопроверка: числа отчёта == ядро (rtol=1e-3)
# ---------------------------------------------------------------------------
def selfcheck(d):
    res = d["res"]
    S = d["SNAP"]
    checks = [
        ("a", S["a"], REFS["a"]),
        ("b", S["b"], REFS["b"]),
        ("r2", S["r2"], REFS["r2"]),
        ("s2_res", S["s2_res"], REFS["s2_res"]),
        ("df_res", S["df_res"], REFS["df_res"]),
        ("F_lof", res["lof"]["F"], REFS["F_lof"]),
        ("F_crit", res["lof"]["F_crit"], REFS["F_crit"]),
        ("sigma_r", res["SIGMA_R"], REFS["sigma_r"]),
        ("sigma_R_tab", res["sigma_R_tab"], REFS["sigma_R_tab"]),
        ("sigma_L", res["sigma_L"], REFS["sigma_L"]),
        ("u_repr", d["u_repr_full"]["u_repr"], REFS["u_repr"]),
        ("u_c", res["U_C_RES"], REFS["u_c"]),
        ("nu_eff", res["NU_EFF_RES"], REFS["nu_eff"]),
        ("k", res["K_RES"], REFS["k"]),
        ("U", res["U_RES_G"], REFS["U"]),
        ("U_pct", res["U_RES_PCT"], REFS["U_pct"]),
        ("U_rel", res["U_REL"] * 100.0, REFS["U_rel"]),
        ("w0", res["W0"], REFS["w0"]),
    ]
    print("=== Самопроверка: числа отчёта vs эталоны (rtol=1e-3) ===")
    allok = True
    for name, val, ref in checks:
        ok = abs(val - ref) <= 1e-3 * abs(ref)
        allok = allok and ok
        print(f"[{'PASS' if ok else 'FAIL'}] {name:<12} = {val:.8g}  (эталон {ref:.8g})")
    print("ИТОГ:", "PASS" if allok else "FAIL")
    return allok


def dump_structure(path):
    """Открыть docx через python-docx и вывести структуру (параграфы/таблицы)."""
    from docx import Document as D2
    doc = D2(path)
    print("=" * 70)
    print(f"СТРУКТУРА {os.path.basename(path)} (параграфы/таблицы по порядку)")
    print("=" * 70)
    for block in doc.element.body:
        tag = block.tag.split("}")[1]
        if tag == "p":
            # текст параграфа
            texts = [t.text or "" for t in block.iter(qn("w:t"))]
            txt = "".join(texts)
            if txt.strip():
                print(f"  P: {txt[:110]}")
        elif tag == "tbl":
            rows = block.findall(qn("w:tr"))
            print(f"  --- ТАБЛИЦА ({len(rows)} строк) ---")
            for tr in rows[:8]:
                cells = []
                for tc in tr.findall(qn("w:tc")):
                    ct = "".join("".join(t.text or "" for t in p.iter(qn("w:t")))
                                 for p in tc.findall(qn("w:p")))
                    cells.append(ct[:26])
                print("     ", " | ".join(cells))
            if len(rows) > 8:
                print(f"      ... ещё {len(rows) - 8} строк")


def main():
    print("build_en13132_v17_word.py — генерация Word-отчёта v17 (step, TD-DEV-208)")
    print("Источник входов:", bld.OUT_XLSX)
    d = collect_data()
    out = build(d)
    print("Отчёт сохранён:", out)
    ok = selfcheck(d)
    dump_structure(out)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())