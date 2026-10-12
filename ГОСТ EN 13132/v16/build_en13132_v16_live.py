#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_en13132_v16_live.py — генерация ЖИВОЙ расчётной книги v16 (блок 5).

EN13132_v16_LIVE.xlsx: каждая расчётная ячейка — ПРОСЛЕЖИВАЕМАЯ ФОРМУЛА Excel
со ссылками на лист «01_Ввод» (вход калибровки/пробы). При изменении входных
данных вся книга пересчитывается Excel-ом БЕЗ перегенерации (требование
пользователя; аудитор блок 4 «block_5_input_live_formulas»).

Листы:
  01_Ввод         данные градуировки (до 12 уровней; Y_ср/n/SS_PE/остатки — формулы)
  02_Регрессия    OLS по средним: n, x̄, Sxx, s²_res, b1, b0, var/cov, R², LoF — живые
  03_ANOVA        σr (из повторностей), σR/σL по табл.1 (step VLOOKUP / smooth INDEX+MATCH),
                  таблица ANOVA (SS/df/MS/F)
  04_Профиль U(X) сетка 200 точек; u_model, u_repr(x), u_c, ν_eff, k, U — ВСЁ живое
  05_Бюджет       5 вкладов + u_c + ν_eff + k + U — живые (модель через обратную градуировку)
  06_Итог         X0, w[%мас], U, U_rel, протокольная запись (конкатенация), секции Word v15
  Графики         3 диаграммы на живые диапазоны
  Инструкция      как пользоваться живой книгой
  Ядро_сверка     [DEV, скрытый] значения ядра anova_gum.py параллельно (для сверки)

Анти-массивные формулы (верификатор formulas 1.3.4 не считает SUMPRODUCT-массивы):
  * суммы-по-строкам вынесены во вспомогательные столбцы 01_Ввод J..N
    (SS_PE_i = DEVSQ(C:E); e²_i = (Y_ср − ŷ)²; e1²/e2²/e3² по инъекциям);
  * s²_res = SUM(K5:K16)/(n−2); SS_res(инъекции) = SUM(L5:N16) — без SUMPRODUCT.

u_repr (табл.1 ГОСТ EN 13132, mean2 = √(R²−r²/2)/2,8):
  * режим 1 — step: VLOOKUP(w; lo-hi-r-R; TRUE) по столбцу lo (R-05: на границе —
    более высокий диапазон; экстраполяция вниз → первый поддиапазон, TD-DEV-202);
  * режим 2 — smooth: INDEX/MATCH + пропорция по узлам R-05 (совпадает с ядром
    u_repr_from_tables(..., interpolation='smooth'), проверено: u_repr(w0)=0.1556745).
  Режим выбирается ячейкой 01_Ввод!B29 (по умолчанию 2=smooth — совпадает с ядром
  и книгой CALC: U=0.0862 г/г; step даст U≈0.0795 г/г — задокументировано).

Протокольная запись (06_Итог) — конкатенация TEXT/ROUND; символы «±»/«ν» в формуле
корректно отображаются в Excel (formulas-пересчёт может их упростить — не влияет
на числа; верификатор проверяет числовые ячейки, не текст протокола).
"""
from __future__ import annotations

import os
import sys
import warnings

import numpy as np
import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import anova_gum as ag  # noqa: E402
import build_en13132_v16_calc as bld  # noqa: E402  (чтение входов, ядро)

V16_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_CALC = os.path.join(V16_DIR, "EN13132_v16_CALC.xlsx")
XLSX_V15 = os.path.join(V16_DIR, "..", "v15", "EN13132_v15_MTBE.xlsx")
OUT_XLSX = os.path.join(V16_DIR, "EN13132_v16_LIVE.xlsx")

N_PROFILE = 200
N_LEVELS_MAX = 12
FIRST_LEVEL = 5
LAST_LEVEL = FIRST_LEVEL + N_LEVELS_MAX - 1  # 16

HDR_FILL = PatternFill("solid", fgColor="DDEBF7")
HDR_FONT = Font(bold=True, size=11)
TITLE_FONT = Font(bold=True, size=13)
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
SUP = "9E480E"

STATUS_TEXT = "ПРЕДВАРИТЕЛЬНО: ВЕСЫ/ПАРТИИ НЕ ПОДТВЕРЖДЕНЫ (TD-DEV-33)"

# ---------------------------------------------------------------------------
# ФИКСИРОВАННЫЕ АДРЕСА (детерминированные ссылки для формул)
# ---------------------------------------------------------------------------
# 01_Ввод
P_ROW = LAST_LEVEL + 2          # 18 — заголовок параметров
MODE_CELL = "'01_Ввод'!$B$29"   # режим u_repr (1=step, 2=smooth)
P_INJ = "'01_Ввод'!$B$24"       # число инъекций пробы p
Y0_CELL = "'01_Ввод'!$B$23"     # Y0 пробы
M_PROBY = "'01_Ввод'!$B$19"
M_IS = "'01_Ввод'!$B$20"
P_IS = "'01_Ввод'!$B$21"
P_A = "'01_Ввод'!$B$22"
U_MNET = "'01_Ввод'!$B$25"
U_PA = "'01_Ввод'!$B$26"
U_REL_ETAL = "'01_Ввод'!$B$27"

# 02_Регрессия (таблица статистики: заголовок в строке 4, данные 5..19)
R_N = 5
R_XBAR = 6
R_YBAR = 7
R_SXX = 8
R_S2 = 9
R_B1 = 10
R_B0 = 11
R_VARB1 = 12
R_VARB0 = 13
R_COV = 14
R_R2 = 15
R_DFRES = 16
R_SRES = 17
R_XMIN = 18
R_XMAX = 19
# 02_Регрессия (LoF: заголовок 22, данные 23..34; k_band 36)
R_SSPE = 23
R_DFPE = 24
R_SSRES = 25
R_NINJ = 26
R_DFRESINJ = 27
R_SSLOF = 28
R_DFLOF = 29
R_MSPE = 30
R_MSLOF = 31
R_FLOF = 32
R_FCRIT = 33
R_STAT = 34
R_KBAND = 36

# 03_ANOVA
R_SSW = 5
R_DFR = 6
R_SIGR = 7
R_W0 = 9
TAB_FIRST = 13
TAB_LAST = 20          # табл.1: A13:D20 (lo|hi|r|R)
NODE_FIRST = 24
NODE_LAST = 32         # узлы smooth: F24:H32 (w|r|R)
R_RSTEP = 35
R_RSTEP2 = 36
R_UREPRSTEP = 37
R_RSMOOTH = 38
R_RSMOOTH2 = 39
R_UREPRSMOOTH = 40
R_SIGMA_R = 43
R_SIGMA_L = 44

# 04_Профиль
PROF_HDR = 4
PROF_FIRST = PROF_HDR + 1              # 5
PROF_LAST = PROF_FIRST + N_PROFILE - 1  # 204

# 05_Бюджет
B_UREL_MASS = 18
B_UREL_PUR = 19
B_UY0 = 20
B_ROW_MODEL = 4
B_ROW_ETAL = 5
B_ROW_PUR = 6
B_ROW_W = 7
B_ROW_REPR = 8
B_UC = 11
B_NUEFF = 12
B_K = 13
B_U = 14
B_UPCT = 15
B_UREL = 16

# 06_Итог
I_KG2PCT = 3
I_X0 = 19
I_W0 = 20
I_UC = 22
I_NUEFF = 23
I_K = 24
I_U = 25
I_UPCT = 26
I_UREL = 27


def comment(text, author="worker"):
    return Comment(text, author)


def set_cell(ws, coord, value, num_fmt=None, bold=False, fill=None, cmt=None):
    cell = ws[coord]
    cell.value = value
    if num_fmt:
        cell.number_format = num_fmt
    if bold:
        cell.font = Font(bold=True)
    if fill:
        cell.fill = fill
    if cmt:
        cell.comment = comment(cmt)
    return cell


def smooth_formula_r(w_expr, rcol="H", wcol="F", first=NODE_FIRST, last=NODE_LAST,
                     sheet="'03_ANOVA'"):
    """Кусочно-линейная интерполяция R(w) по узлам R-05 (колонка rcol по wcol).
    Узлы на листе 03_ANOVA (F/H или G/H колонки, строки first..last).
    Ниже нижнего узла — экстраполяция первого сегмента (ядро _interp_rr),
    выше верхнего — последний узел."""
    W = f"{sheet}!${wcol}${first}:${wcol}${last}"
    V = f"{sheet}!${rcol}${first}:${rcol}${last}"
    return (
        f"IF({w_expr}<{sheet}!${wcol}${first},{sheet}!${rcol}${first}"
        f"+({w_expr}-{sheet}!${wcol}${first})"
        f"*({sheet}!${rcol}${first + 1}-{sheet}!${rcol}${first})"
        f"/({sheet}!${wcol}${first + 1}-{sheet}!${wcol}${first}),"
        f"IF({w_expr}>={sheet}!${wcol}${last},{sheet}!${rcol}${last},"
        f"INDEX({V},MATCH({w_expr},{W},1))"
        f"+(INDEX({V},MATCH({w_expr},{W},1)+1)-INDEX({V},MATCH({w_expr},{W},1)))"
        f"*({w_expr}-INDEX({W},MATCH({w_expr},{W},1)))"
        f"/(INDEX({W},MATCH({w_expr},{W},1)+1)-INDEX({W},MATCH({w_expr},{W},1)))))"
    )


def urepr_formula(w_expr, mode_cell=MODE_CELL):
    """Полная формула u_repr(mean2) = √(max(0,R²−r²/2))/2,8 по активному режиму.
    Табл.1 и узлы smooth берутся ЯВНО с листа 03_ANOVA (именно поэтому формулы
    корректны и на 04_Профиль, и на 05_Бюджет). Возвращает строку формулы БЕЗ '='."""
    tab = f"'03_ANOVA'!$A${TAB_FIRST}:$D${TAB_LAST}"
    r_step = f"IFERROR(VLOOKUP({w_expr},{tab},4,TRUE),0.1)"
    r_smooth = smooth_formula_r(w_expr, rcol="H", wcol="F")
    R_sel = f"IF({mode_cell}=1,{r_step},{r_smooth})"
    r_step_r = f"IFERROR(VLOOKUP({w_expr},{tab},3,TRUE),0.05)"
    r_smooth_r = smooth_formula_r(w_expr, rcol="G", wcol="F")
    r_sel = f"IF({mode_cell}=1,{r_step_r},{r_smooth_r})"
    return f"SQRT(MAX(0,{R_sel}^2-{r_sel}^2/2))/2.8"


# ---------------------------------------------------------------------------
# ЛИСТ 01_Ввод
# ---------------------------------------------------------------------------
def write_input_sheet_live(wb, levels4, inputs, mode):
    ws = wb.create_sheet("01_Ввод")
    ws["A1"] = "ГРАДУИРОВКА МТБЭ | ГОСТ EN 13132-2012 | ЖИВАЯ КНИГА (LIVE)"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:N1")
    ws.append([])
    ws["A3"] = ("Данные градуировки (до 12 уровней, инъекции 1..3). Y_ср, n, SS_PE и остатки — "
                "ЖИВЫЕ формулы. При изменении X/Y всё пересчитается автоматически.")
    ws["A3"].font = Font(italic=True, color=SUP)
    ws.merge_cells("A3:N3")

    hdr = 4
    headers = ["Уровень", "X (г/г)", "Y1", "Y2", "Y3", "Y_ср", "n", "mA (г)", "mИС (г)",
               "SS_PE_i", "e²_i", "e1²", "e2²", "e3²"]
    for c, h in enumerate(headers, 1):
        ws.cell(hdr, c, h)
        ws.cell(hdr, c).fill = HDR_FILL
        ws.cell(hdr, c).font = HDR_FONT
        ws.cell(hdr, c).border = BORDER

    for i in range(N_LEVELS_MAX):
        r = FIRST_LEVEL + i
        ws.cell(r, 1, i + 1)
        for c in range(1, 15):
            ws.cell(r, c).border = BORDER
        ws.cell(r, 6, f'=IF(COUNT(C{r}:E{r})>0,AVERAGE(C{r}:E{r}),"")')
        ws.cell(r, 6).number_format = "0.000000"
        ws.cell(r, 7, f'=IF(COUNT(C{r}:E{r})>0,COUNT(C{r}:E{r}),"")')
        ws.cell(r, 10, f'=IF(COUNT(C{r}:E{r})>=2,DEVSQ(C{r}:E{r}),0)')
        ws.cell(r, 11, f'=IF(ISNUMBER(F{r}),(F{r}-(\'02_Регрессия\'!$B${R_B1}*B{r}+\'02_Регрессия\'!$B${R_B0}))^2,"")')
        ws.cell(r, 12, f'=IF(ISNUMBER(C{r}),(C{r}-(\'02_Регрессия\'!$B${R_B1}*B{r}+\'02_Регрессия\'!$B${R_B0}))^2,"")')
        ws.cell(r, 13, f'=IF(ISNUMBER(D{r}),(D{r}-(\'02_Регрессия\'!$B${R_B1}*B{r}+\'02_Регрессия\'!$B${R_B0}))^2,"")')
        ws.cell(r, 14, f'=IF(ISNUMBER(E{r}),(E{r}-(\'02_Регрессия\'!$B${R_B1}*B{r}+\'02_Регрессия\'!$B${R_B0}))^2,"")')

    for i, (x, ys, mA, mIS) in enumerate(levels4):
        r = FIRST_LEVEL + i
        ws.cell(r, 2, x).number_format = "0.000000"
        for j in range(3):
            if j < len(ys):
                ws.cell(r, 3 + j, ys[j]).number_format = "0.000000"
        if mA is not None:
            ws.cell(r, 8, mA).number_format = "0.000"
        if mIS is not None:
            ws.cell(r, 9, mIS).number_format = "0.000"

    p_row = P_ROW
    ws.cell(p_row, 1, "ПАРАМЕТРЫ РЕЗУЛЬТАТА ПРОБЫ (ввод методиста)")
    ws.merge_cells(f"A{p_row}:I{p_row}")
    ws.cell(p_row, 1).font = TITLE_FONT
    for i, (key, label, _def, fmt, units, cmt) in enumerate(bld.INPUT_FIELDS):
        r = p_row + 1 + i
        ws.cell(r, 1, label)
        ws.cell(r, 1).alignment = Alignment(wrap_text=True)
        set_cell(ws, f"B{r}", inputs[key], num_fmt=fmt, cmt=cmt)
        ws.cell(r, 1).border = BORDER
        ws.cell(r, 2).border = BORDER
        ws.cell(r, 3, units)
        ws.cell(r, 3).border = BORDER

    r_mode = p_row + 1 + len(bld.INPUT_FIELDS) + 1  # 29
    ws.cell(r_mode, 1, "Режим u_repr (1=step, 2=smooth)")
    ws.cell(r_mode, 1).border = BORDER
    set_cell(ws, f"B{r_mode}", mode, num_fmt="0",
             cmt="1 — ступенчатая по табл.1 (R-05, РМГ61 §5.5, консервативная); "
                 "2 — кусочно-линейная по узлам R-05 (как ядро anova_gum.py, рекомендована).")
    ws.cell(r_mode, 3, "1/2")
    ws.cell(r_mode, 3).border = BORDER

    st_r = r_mode + 2
    ws.cell(st_r, 1, "Статус").font = Font(bold=True)
    set_cell(ws, f"B{st_r}", STATUS_TEXT,
             cmt="Блокировка вывода протокола до паспортных данных (контракт §10, TD-DEV-33)")

    for col in "JKLMN":
        ws.column_dimensions[col].hidden = True
        ws.column_dimensions[col].width = 12
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 14
    for col in "CDEFGHI":
        ws.column_dimensions[col].width = 12
    return ws


# ---------------------------------------------------------------------------
# ЛИСТ 02_Регрессия
# ---------------------------------------------------------------------------
def fill_regression_live(wb, levels, inputs):
    ws = wb.create_sheet("02_Регрессия")
    ws["A1"] = "ГРАДУИРОВОЧНАЯ КРИВАЯ (OLS по средним) — ЖИВЫЕ ФОРМУЛЫ"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:F1")
    ws["A2"] = ("n=COUNT, x̄=AVERAGE, Sxx=DEVSQ, b1=SLOPE, b0=INTERCEPT, R²=RSQ; "
                "все формулы ссылаются на 01_Ввод. Изменение X/Y пересчитывает всё.")
    ws["A2"].font = Font(italic=True, color=SUP)
    ws.merge_cells("A2:F2")

    ws.append([])
    ws.append(["Параметр", "Значение", "Ед.", "Примечание"])  # строка 4
    hr = 4
    for c in range(1, 5):
        ws.cell(hr, c).fill = HDR_FILL
        ws.cell(hr, c).font = HDR_FONT
        ws.cell(hr, c).border = BORDER

    rows = [
        ("n (число уровней)", "=COUNT('01_Ввод'!B5:B16)", "1", "по заполненным X"),
        ("x̄ — центр калибровки", "=AVERAGE('01_Ввод'!B5:B16)", "г/г", "среднее X"),
        ("ȳ — средний отклик", "=AVERAGE('01_Ввод'!F5:F16)", "Y", "по Y_ср уровней"),
        ("Sxx", "=DEVSQ('01_Ввод'!B5:B16)", "(г/г)²", "Σ(x−x̄)²"),
        ("s²_res (остаточная дисперсия)",
         f"=IF(B{R_N}>2,SUM('01_Ввод'!K5:K16)/(B{R_N}-2),NA())",
         "Y²", "Σe²_i/(n−2); e²_i — вспом. столбец K на 01_Ввод"),
        ("Наклон b1", "=SLOPE('01_Ввод'!F5:F16,'01_Ввод'!B5:B16)", "Y/X", "OLS по средним"),
        ("Свободный член b0", "=INTERCEPT('01_Ввод'!F5:F16,'01_Ввод'!B5:B16)", "Y", "OLS"),
        ("var(b1)", f"=IF(ISNUMBER(B{R_S2}),B{R_S2}/B{R_SXX},NA())", "(Y/X)²", "s²/Sxx"),
        ("var(b0)",
         f"=IF(ISNUMBER(B{R_S2}),B{R_S2}*(1/B{R_N}+B{R_XBAR}^2/B{R_SXX}),NA())",
         "Y²", "s²·(1/n+x̄²/Sxx)"),
        ("cov(b0,b1)", f"=IF(ISNUMBER(B{R_S2}),-B{R_S2}*B{R_XBAR}/B{R_SXX},NA())", "Y²/X",
         "−s²·x̄/Sxx"),
        ("R²", "=RSQ('01_Ввод'!F5:F16,'01_Ввод'!B5:B16)", "1", "коэффициент детерминации"),
        ("df_res", f"=IF(B{R_N}>2,B{R_N}-2,NA())", "1", "n−2"),
        ("s_res", f"=IF(ISNUMBER(B{R_S2}),SQRT(B{R_S2}),NA())", "Y", "√s²_res"),
        ("X min", "=MIN('01_Ввод'!B5:B16)", "г/г", "нижняя граница калибровки"),
        ("X max", "=MAX('01_Ввод'!B5:B16)", "г/г", "верхняя граница калибровки"),
    ]
    for i, (label, formula, unit, note) in enumerate(rows):
        r = hr + 1 + i  # 5..19
        ws.cell(r, 1, label)
        set_cell(ws, f"B{r}", formula, num_fmt="0.00000000")
        ws.cell(r, 3, unit)
        ws.cell(r, 4, note)
        for c in range(1, 5):
            ws.cell(r, c).border = BORDER

    # --- LoF (заголовок 22, данные 23..34) ---
    lof_hdr = hr + 1 + len(rows) + 2  # 22
    ws.cell(lof_hdr, 1, "Lack-of-fit (диагностика; по повторностям инъекций, 5725-2)")
    ws.cell(lof_hdr, 1).font = HDR_FONT
    ws.merge_cells(f"A{lof_hdr}:F{lof_hdr}")
    lof_rows = [
        ("SS_PE (внутри уровней)", "=SUM('01_Ввод'!J5:J16)", "Y²",
         "Σ_i DEVSQ(C_i:E_i); вспом. столбец J на 01_Ввод"),
        ("df_PE", "=SUMIF('01_Ввод'!G5:G16,\">0\",'01_Ввод'!G5:G16)-COUNTIF('01_Ввод'!G5:G16,\">0\")",
         "1", "Σ(n_i−1)"),
        ("SS_res (инъекции)", "=SUM('01_Ввод'!L5:N16)", "Y²",
         "Σ_ij(Y_ij−ŷ_i)²; вспом. столбцы L..N на 01_Ввод"),
        ("N (всего инъекций)", "=SUM('01_Ввод'!G5:G16)", "шт", "Σn_i"),
        ("df_res(инъекции)", f"=IF(B{R_NINJ}>2,B{R_NINJ}-2,NA())", "1", "N−2"),
        ("SS_LoF", f"=MAX(0,B{R_SSRES}-B{R_SSPE})", "Y²", "SS_res−SS_PE"),
        ("df_LoF", f"=IF(AND(ISNUMBER(B{R_DFRESINJ}),ISNUMBER(B{R_DFPE})),B{R_DFRESINJ}-B{R_DFPE},NA())",
         "1", "df_res−df_PE"),
        ("MS_PE", f"=IF(B{R_DFPE}>0,B{R_SSPE}/B{R_DFPE},NA())", "Y²", "SS_PE/df_PE"),
        ("MS_LoF", f"=IF(B{R_DFLOF}>0,B{R_SSLOF}/B{R_DFLOF},NA())", "Y²", "SS_LoF/df_LoF"),
        ("F LoF", f"=IF(AND(ISNUMBER(B{R_MSPE}),ISNUMBER(B{R_MSLOF}),B{R_MSPE}>0),B{R_MSLOF}/B{R_MSPE},NA())",
         "1", "MS_LoF/MS_PE"),
        ("F крит. 95%",
         f"=IF(AND(ISNUMBER(B{R_FLOF}),B{R_DFLOF}>0,B{R_DFPE}>0),FINV(0.05,B{R_DFLOF},B{R_DFPE}),NA())",
         "1", "FINV(0.05; df_LoF; df_PE)"),
        ("Статус LoF",
         f"=IF(NOT(ISNUMBER(B{R_FLOF})),\"н/д\",IF(B{R_FLOF}<=B{R_FCRIT},\"НЕ ВЫЯВЛЕНА\","
         f"\"ВЫЯВЛЕНА (модель не утверждается)\"))", "", "F ≤ Fкр → линейность адекватна"),
    ]
    for i, (label, formula, unit, note) in enumerate(lof_rows):
        r = lof_hdr + 1 + i  # 23..34
        ws.cell(r, 1, label)
        set_cell(ws, f"B{r}", formula, num_fmt="0.000000")
        ws.cell(r, 3, unit)
        ws.cell(r, 4, note)
        for c in range(1, 5):
            ws.cell(r, c).border = BORDER

    k_band_r = lof_hdr + 1 + len(lof_rows) + 1  # 36
    ws.cell(k_band_r, 1, "k_band = t_{0.975}(df_res)")
    set_cell(ws, f"B{k_band_r}", f"=IF(B{R_DFRES}>0,TINV(0.05,B{R_DFRES}),2)", num_fmt="0.0000")
    ws.cell(k_band_r, 3, "1")
    ws.cell(k_band_r, 4, "для полосы ŷ±k·u(ŷ)")
    for c in range(1, 5):
        ws.cell(k_band_r, c).border = BORDER

    # --- Таблица точек для графика (живые формулы) ---
    # Столбец G — константа 0 для ЛИНИИ НУЛЯ графика остатков: xVal и yVal
    # серии '0 (линия нуля)' должны лежать на ОДНОМ листе (Excel-правило),
    # поэтому нули здесь, на 02_Регрессия, а не на «Графики».
    tbl_hdr = k_band_r + 2  # 38
    for c, h in enumerate(["X", "Y_ср", "Y_расч (линия)", "e (остаток)", "ŷ−k·u(ŷ)",
                           "ŷ+k·u(ŷ)", "0 (линия нуля)"], 1):
        ws.cell(tbl_hdr, c, h)
        ws.cell(tbl_hdr, c).fill = HDR_FILL
        ws.cell(tbl_hdr, c).font = HDR_FONT
        ws.cell(tbl_hdr, c).border = BORDER
    ws.cell(tbl_hdr, 7).comment = comment(
        "Константа 0 для линии нуля графика остатков. xVal и yVal серии '0' "
        "на ОДНОМ листе (02_Регрессия) — Excel-правило ScatterChart."
    )
    DATA_FIRST = tbl_hdr + 1  # 39
    DATA_LAST = DATA_FIRST + N_LEVELS_MAX - 1  # 50
    for i in range(N_LEVELS_MAX):
        r = DATA_FIRST + i
        src = FIRST_LEVEL + i
        ws.cell(r, 1, f'=IF(ISNUMBER(\'01_Ввод\'!B{src}),\'01_Ввод\'!B{src},"")')
        ws.cell(r, 1).number_format = "0.000000"
        ws.cell(r, 2, f'=IF(ISNUMBER(\'01_Ввод\'!F{src}),\'01_Ввод\'!F{src},"")')
        ws.cell(r, 2).number_format = "0.000000"
        ws.cell(r, 3, f"=IF(ISNUMBER(A{r}),$B${R_B1}*A{r}+$B${R_B0},\"\")")
        ws.cell(r, 3).number_format = "0.000000"
        ws.cell(r, 4, f"=IF(ISNUMBER(B{r}),B{r}-C{r},\"\")")
        ws.cell(r, 4).number_format = "0.000000"
        uy = f"SQRT($B${R_S2}*(1/$B${R_N}+(A{r}-$B${R_XBAR})^2/$B${R_SXX}))"
        ws.cell(r, 5, f'=IF(ISNUMBER(C{r}),C{r}-$B${k_band_r}*{uy},"")')
        ws.cell(r, 5).number_format = "0.000000"
        ws.cell(r, 6, f'=IF(ISNUMBER(C{r}),C{r}+$B${k_band_r}*{uy},"")')
        ws.cell(r, 6).number_format = "0.000000"
        ws.cell(r, 7, 0.0).number_format = "0.000000"
        for c in range(1, 8):
            ws.cell(r, c).border = BORDER

    GRID_PT = DATA_LAST + 2  # 52
    for i in range(100):
        r = GRID_PT + i
        ws.cell(r, 1, f"=MIN('01_Ввод'!$B$5:$B$16)+(ROW()-{r})*(MAX('01_Ввод'!$B$5:$B$16)-MIN('01_Ввод'!$B$5:$B$16))/99")
        ws.cell(r, 1).number_format = "0.000000"
        ws.cell(r, 3, f'=IF(ISNUMBER(A{r}),$B${R_B1}*A{r}+$B${R_B0},"")')
        ws.cell(r, 3).number_format = "0.000000"
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 56
    return ws, DATA_FIRST, DATA_LAST, GRID_PT


# ---------------------------------------------------------------------------
# ЛИСТ 03_ANOVA
# ---------------------------------------------------------------------------
def fill_anova_live(wb, levels, inputs):
    ws = wb.create_sheet("03_ANOVA")
    ws["A1"] = "КЛАССИЧЕСКАЯ ANOVA ПО GUM (грубый учёт табл.1/2 ГОСТ EN 13132) — ЖИВЫЕ ФОРМУЛЫ"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:H1")
    ws["A2"] = ("σr — из повторностей (столбец J на 01_Ввод); σR/σL — по табл.1 (режим "
                "01_Ввод!B29: 1=step VLOOKUP, 2=smooth INDEX/MATCH по узлам R-05); "
                "u_repr(mean2)=√(R²−r²/2)/2,8.")
    ws["A2"].font = Font(italic=True, color=SUP)
    ws.merge_cells("A2:H2")

    # --- Компоненты (фиксированные адреса) ---
    for c, h in enumerate(["Компонента", "Символ", "Значение", "Ед.", "Источник/формула"], 1):
        ws.cell(4, c, h)
        ws.cell(4, c).fill = HDR_FILL
        ws.cell(4, c).font = HDR_FONT
        ws.cell(4, c).border = BORDER
    ws.cell(5, 1, "SS_within (pure error)")
    set_cell(ws, f"B{R_SSW}", "=SUM('01_Ввод'!J5:J16)", num_fmt="0.000000")
    ws.cell(5, 3, "Y²")
    ws.cell(5, 5, "Σ_i DEVSQ(C_i:E_i) — вспом. столбец J на 01_Ввод")
    ws.cell(6, 1, "df_r")
    set_cell(ws, f"B{R_DFR}",
             "=SUMIF('01_Ввод'!G5:G16,\">0\",'01_Ввод'!G5:G16)-COUNTIF('01_Ввод'!G5:G16,\">0\")",
             num_fmt="0")
    ws.cell(6, 3, "1")
    ws.cell(6, 5, "Σ(n_i−1)")
    ws.cell(7, 1, "Повторяемость инъекций")
    ws.cell(7, 2, "σr")
    set_cell(ws, f"C{R_SIGR}", f"=IF(B{R_DFR}>0,SQRT(B{R_SSW}/B{R_DFR}),NA())", num_fmt="0.000000")
    ws.cell(7, 4, "Y")
    ws.cell(7, 5, "√(SS_within/df_r); GUM H.5 / 5725-6 4.1.4")
    for r in (5, 6, 7):
        for c in range(1, 6):
            ws.cell(r, c).border = BORDER

    ws.cell(9, 1, "w0 (рабочая точка, % мас.)")
    set_cell(ws, f"B{R_W0}", f"=IF(ISNUMBER('06_Итог'!$B${I_W0}),'06_Итог'!$B${I_W0},\"\")",
             num_fmt="0.0000")
    ws.cell(9, 3, "% мас.")
    ws.cell(9, 5, "w0 = X0·K_G2PCT (06_Итог!B19)")
    for c in range(1, 6):
        ws.cell(9, c).border = BORDER

    # --- табл.1 (step): A13:D20 ---
    ws.cell(11, 1, "Табл.1 ГОСТ EN 13132 (поддиапазоны; lo — столбец VLOOKUP)")
    ws.cell(11, 1).font = HDR_FONT
    ws.merge_cells("A11:H11")
    for c, h in enumerate(["lo", "hi", "r", "R"], 1):
        ws.cell(12, c, h)
        ws.cell(12, c).fill = HDR_FILL
        ws.cell(12, c).font = HDR_FONT
        ws.cell(12, c).border = BORDER
    for i, (lo, hi, r, R) in enumerate(ag.TABLE1_RANGES):
        rw = TAB_FIRST + i
        ws.cell(rw, 1, lo)
        ws.cell(rw, 2, hi)
        ws.cell(rw, 3, r)
        ws.cell(rw, 4, R)
        for c in range(1, 5):
            ws.cell(rw, c).border = BORDER

    # --- узлы smooth (w|r|R): F24:H32 ---
    ws.cell(22, 1, "Узлы R-05 для smooth (w, r, R)")
    ws.cell(22, 1).font = HDR_FONT
    ws.merge_cells("A22:H22")
    for c, h in enumerate(["w", "r", "R"], 6):
        ws.cell(23, c, h)
        ws.cell(23, c).fill = HDR_FILL
        ws.cell(23, c).font = HDR_FONT
        ws.cell(23, c).border = BORDER
    for i, (w, r, R) in enumerate(ag._smooth_nodes(ag.TABLE1_RANGES)):
        rw = NODE_FIRST + i
        ws.cell(rw, 6, w)
        ws.cell(rw, 7, r)
        ws.cell(rw, 8, R)
        for c in range(6, 9):
            ws.cell(rw, c).border = BORDER

    # --- R/r/u_repr по режиму (B35..B40) ---
    w_expr = f"B{R_W0}"
    tab_range = f"$A${TAB_FIRST}:$D${TAB_LAST}"
    R_step = f"IFERROR(VLOOKUP({w_expr},{tab_range},4,TRUE),0.1)"
    r_step = f"IFERROR(VLOOKUP({w_expr},{tab_range},3,TRUE),0.05)"
    R_smooth = smooth_formula_r(w_expr, rcol="H", wcol="F")
    r_smooth = smooth_formula_r(w_expr, rcol="G", wcol="F")
    pairs = [
        (R_RSTEP, "R(w0) — step", R_step, "1"),
        (R_RSTEP2, "r(w0) — step", r_step, "1"),
        (R_UREPRSTEP, "u_repr(mean2) — step", f"SQRT(MAX(0,{R_step}^2-{r_step}^2/2))/2.8", "% мас."),
        (R_RSMOOTH, "R(w0) — smooth", R_smooth, "1"),
        (R_RSMOOTH2, "r(w0) — smooth", r_smooth, "1"),
        (R_UREPRSMOOTH, "u_repr(mean2) — smooth",
         f"SQRT(MAX(0,{R_smooth}^2-{r_smooth}^2/2))/2.8", "% мас."),
    ]
    for rw, label, formula, unit in pairs:
        ws.cell(rw, 1, label)
        set_cell(ws, f"B{rw}", "=" + formula, num_fmt="0.000000")
        ws.cell(rw, 3, unit)
        for c in range(1, 4):
            ws.cell(rw, c).border = BORDER

    # --- σR / σL по активному режиму ---
    ws.cell(R_SIGMA_R, 1, "Воспроизводимость табл.1")
    ws.cell(R_SIGMA_R, 2, "σR")
    set_cell(ws, f"C{R_SIGMA_R}",
             f"=IF({MODE_CELL}=1,{R_step}/2.8,{R_smooth}/2.8)", num_fmt="0.000000")
    ws.cell(R_SIGMA_R, 4, "% мас.")
    ws.cell(R_SIGMA_R, 5, "R(w0)/2,8 (5725-6 4.1.2/4.1.4)")
    ws.cell(R_SIGMA_L, 1, "Межлабораторная составляющая")
    ws.cell(R_SIGMA_L, 2, "σL")
    set_cell(ws, f"C{R_SIGMA_L}",
             f"=IF(AND(ISNUMBER(C{R_SIGMA_R}),ISNUMBER(C{R_SIGR})),"
             f"SQRT(MAX(0,C{R_SIGMA_R}^2-C{R_SIGR}^2)),NA())", num_fmt="0.000000")
    ws.cell(R_SIGMA_L, 4, "% мас.")
    ws.cell(R_SIGMA_L, 5, "√(max(0, σR²−σr²)); GUM 4.2; грубый учёт (шкалы σr/σR разные)")
    for c in range(1, 6):
        ws.cell(R_SIGMA_R, c).border = BORDER
        ws.cell(R_SIGMA_L, c).border = BORDER

    # --- Таблица ANOVA ---
    av_hdr = R_SIGMA_L + 2  # 46
    ws.cell(av_hdr, 1, "Таблица ANOVA (внутриуровневая, источник: инъекции)")
    ws.merge_cells(f"A{av_hdr}:F{av_hdr}")
    ws.cell(av_hdr, 1).font = HDR_FONT
    for c, h in enumerate(["Источник", "SS", "df", "MS", "F", "Примечание"], 1):
        ws.cell(av_hdr + 1, c, h)
        ws.cell(av_hdr + 1, c).fill = HDR_FILL
        ws.cell(av_hdr + 1, c).font = HDR_FONT
        ws.cell(av_hdr + 1, c).border = BORDER
    ms_pe = (f"=IF(AND(ISNUMBER(B{av_hdr + 2}),ISNUMBER(C{av_hdr + 2}),C{av_hdr + 2}>0),"
             f"B{av_hdr + 2}/C{av_hdr + 2},NA())")
    av_rows = [
        ("Внутри уровней (pure error)",
         "=SUM('01_Ввод'!J5:J16)",
         "=SUMIF('01_Ввод'!G5:G16,\">0\",'01_Ввод'!G5:G16)-COUNTIF('01_Ввод'!G5:G16,\">0\")",
         ms_pe,
         "",
         "SS_PE=ΣDEVSQ(C:E); MS=SS/df"),
        ("Между уровнями (регрессия)",
         "",
         f"=IF(ISNUMBER('02_Регрессия'!$B${R_N}),'02_Регрессия'!$B${R_N}-1,NA())",
         "",
         "",
         "не случайный эффект (систематика калибровки)"),
        ("Lack-of-fit",
         f"='02_Регрессия'!$B${R_SSLOF}",
         f"='02_Регрессия'!$B${R_DFLOF}",
         f"='02_Регрессия'!$B${R_MSLOF}",
         f"='02_Регрессия'!$B${R_FLOF}",
         "F=MS_LoF/MS_PE; Fкр=02_Регрессия!B33"),
        ("Итого",
         "",
         "=SUM('01_Ввод'!G5:G16)",
         "",
         "",
         "N — всего инъекций"),
    ]
    for i, (label, ss, df, ms, f, note) in enumerate(av_rows):
        r = av_hdr + 2 + i
        ws.cell(r, 1, label)
        if ss:
            set_cell(ws, f"B{r}", ss, num_fmt="0.000000")
        if df:
            set_cell(ws, f"C{r}", df, num_fmt="0")
        if ms:
            set_cell(ws, f"D{r}", ms, num_fmt="0.000000")
        if f:
            set_cell(ws, f"E{r}", f, num_fmt="0.000000")
        ws.cell(r, 6, note)
        for c in range(1, 7):
            ws.cell(r, c).border = BORDER

    note_r = av_hdr + 2 + len(av_rows) + 1
    ws.cell(note_r, 1, "ПРИМЕЧАНИЯ (GUM H.5.3.3, контракт разд.3/7):")
    ws.cell(note_r, 1).font = HDR_FONT
    for i, t in enumerate([
        "· Полный иерархический ANOVA (σr/σL/σR) невозможен на однодневных данных (1 день, 1 лаборатория) — GUM H.5.3.3.",
        "· Межсерийная σL данных лаборатории не оцениваема (QC BUD-17, вне плана).",
        "· Применён «грубый учёт»: σR — из табл.1/2 ГОСТ EN 13132 (5725-6 §7, РМГ 61 §5.5).",
    ]):
        ws.cell(note_r + 1 + i, 1, t).font = Font(italic=True, color=SUP)
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 8
    ws.column_dimensions["F"].width = 22
    return ws


# ---------------------------------------------------------------------------
# ЛИСТ 04_Профиль U(X)
# ---------------------------------------------------------------------------
def fill_profile_live(wb, levels, inputs):
    ws = wb.create_sheet("04_Профиль U(X)")
    ws["A1"] = "ПРОФИЛЬ РАСШИРЕННОЙ НЕОПРЕДЕЛЁННОСТИ U(X) — ЖИВЫЕ ФОРМУЛЫ"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:H1")
    ws["A2"] = (
        "U(x)=k(x)·u_c(x); сетка 200 точек (Xmax−Xmin)/199; модель s²·(1/p+1/n+(x−x̄)²/Sxx)/b1²; "
        "u_repr — по режиму 01_Ввод!B29 (1=step VLOOKUP, 2=smooth INDEX/MATCH по узлам R-05); "
        "k=TINV(0.05;ROUND(ν_eff;0)). Все столбцы — формулы, пересчитываются при изменении входов."
    )
    ws.merge_cells("A2:H2")
    ws["A2"].font = Font(italic=True, color=SUP)
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[2].height = 44

    ws["A3"] = "шаг сетки"
    set_cell(ws, "B3", "=(MAX('01_Ввод'!$B$5:$B$16)-MIN('01_Ввод'!$B$5:$B$16))/199",
             num_fmt="0.000000")

    for c, h in enumerate(["№", "X (г/г)", "u_model(x)", "u_repr(x) г/г", "u_c(x)",
                           "ν_eff(x)", "k(x)", "U(x)"], 1):
        ws.cell(PROF_HDR, c, h)
        ws.cell(PROF_HDR, c).fill = HDR_FILL
        ws.cell(PROF_HDR, c).font = HDR_FONT
        ws.cell(PROF_HDR, c).border = BORDER

    K = f"'06_Итог'!$B${I_KG2PCT}"
    S2 = f"'02_Регрессия'!$B${R_S2}"
    N = f"'02_Регрессия'!$B${R_N}"
    XBAR = f"'02_Регрессия'!$B${R_XBAR}"
    SXX = f"'02_Регрессия'!$B${R_SXX}"
    B1 = f"'02_Регрессия'!$B${R_B1}"
    DF = f"'02_Регрессия'!$B${R_DFRES}"

    for i in range(N_PROFILE):
        r = PROF_FIRST + i
        ws.cell(r, 1, f"=ROW()-{PROF_HDR}")
        ws.cell(r, 2, f"=MIN('01_Ввод'!$B$5:$B$16)+(ROW()-{PROF_FIRST})*$B$3")
        ws.cell(r, 2).number_format = "0.000000"
        # J (скрытый): w(x) = X·K_G2PCT (% мас.) — отдельная ячейка, чтобы D ссылалась
        # на простую ссылку (formulas-библиотека не считает MATCH(выражение,диапазон,1))
        ws.cell(r, 10, f"=IF(ISNUMBER(B{r}),B{r}*{K},\"\")")
        ws.cell(r, 10).number_format = "0.0000"
        # C: u_model(x)
        ws.cell(r, 3, f"=IF(ISNUMBER({S2}),SQRT({S2}*(1/{P_INJ}+1/{N}+(B{r}-{XBAR})^2/{SXX}))/ABS({B1}),\"\")")
        ws.cell(r, 3).number_format = "0.000000"
        # D: u_repr(x) г/г  (u_repr_pct(J{r}) / K_G2PCT)
        ws.cell(r, 4, f"=IF(ISNUMBER(J{r}),{urepr_formula(f'J{r}')}/{K},\"\")")
        ws.cell(r, 4).number_format = "0.000000"
        # E: u_c(x)
        ws.cell(r, 5, f"=IF(ISNUMBER(C{r}),SQRT(C{r}^2+(B{r}*{U_REL_ETAL})^2+(B{r}*'05_Бюджет'!$B${B_UREL_PUR})^2+(B{r}*'05_Бюджет'!$B${B_UREL_MASS})^2+D{r}^2),\"\")")
        ws.cell(r, 5).number_format = "0.000000"
        # F: ν_eff(x) — Welch–Satterthwaite: только модель (df_res) и эталоны (ν=50) конечны
        ws.cell(r, 6, f"=IF(OR(NOT(ISNUMBER(E{r})),E{r}=0),NA(),E{r}^4/(C{r}^4/{DF}+(B{r}*{U_REL_ETAL})^4/50))")
        ws.cell(r, 6).number_format = "0.0"
        # G: k(x)
        ws.cell(r, 7, f"=IF(ISNUMBER(F{r}),TINV(0.05,MAX(ROUND(F{r},0),1)),NA())")
        ws.cell(r, 7).number_format = "0.0000"
        # H: U(x) = k·u_c  (живая формула)
        ws.cell(r, 8, f"=G{r}*E{r}")
        ws.cell(r, 8).number_format = "0.000000"
        for c in range(1, 11):
            ws.cell(r, c).border = BORDER

    # скрыть вспомогательный столбец J
    ws.column_dimensions["J"].hidden = True
    ws.column_dimensions["J"].width = 12

    # маркер x_pred (колонка I)
    ws.cell(PROF_HDR, 9, "маркер")
    ws.cell(PROF_HDR, 9).fill = HDR_FILL
    ws.cell(PROF_HDR, 9).font = HDR_FONT
    for i in range(N_PROFILE):
        r = PROF_FIRST + i
        ws.cell(r, 9, f"=IF(ABS(B{r}-'06_Итог'!$B${I_X0})<=$B$3/2,\"← x_pred\",\"\")")
        ws.cell(r, 9).font = Font(bold=True, color="C00000")

    ws.cell(PROF_HDR, 8).comment = comment(
        "Значения U в столбце H — ЖИВЫЕ формулы Excel (=k·u_c, G*E). "
        "Кэша значений в книге нет; при чтении без пересчёта используйте ядро anova_gum.py."
    )

    # Маркер x_pred для диаграммы профиля: xVal/yVal серии маркера должны лежать
    # на листе ПРОФИЛЯ (Excel-правило «серия — один лист»), а не на «Графики».
    # K4 = X0 (06_Итог!B19), L4 = U (05_Бюджет!B14) — ЖИВЫЕ формулы.
    ws.cell(PROF_HDR - 1, 11, "Маркер x_pred (X0; U(x0)) — данные серии на этом листе").font = \
        Font(italic=True, color=SUP)
    set_cell(ws, "K4", f"='06_Итог'!$B${I_X0}", num_fmt="0.000000",
             cmt="xVal серии-маркера 'x_pred' диаграммы профиля; на листе 04_Профиль "
                 "U(X) вместе с yVal (L4) — Excel-правило (иначе X -> 1..n).")
    set_cell(ws, "L4", f"='05_Бюджет'!$B${B_U}", num_fmt="0.000000",
             cmt="yVal серии-маркера 'x_pred' (= U(06_Итог), TD-DEV-206); на листе "
                 "04_Профиль U(X) вместе с xVal (K4).")
    ws.column_dimensions["A"].width = 6
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["I"].width = 12
    ws.column_dimensions["K"].width = 14
    ws.column_dimensions["L"].width = 14
    return ws, PROF_FIRST, PROF_LAST


# ---------------------------------------------------------------------------
# ЛИСТ 05_Бюджет
# ---------------------------------------------------------------------------
def fill_budget_live(wb, levels, inputs):
    ws = wb.create_sheet("05_Бюджет")
    ws["A1"] = "БЮДЖЕТ НЕОПРЕДЕЛЁННОСТИ РЕЗУЛЬТАТА ПРОБЫ — ЖИВЫЕ ФОРМУЛЫ"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:F1")
    ws["A2"] = ("Вклады: модель (обратная градуировка с var/cov, u_y0²=s²_res/p — эквивалент "
                "интервальной формы), эталоны, чистота, весы, прецизионность (табл.1 по режиму); "
                "u_c=√Σu², ν_eff (Welch–Satterthwaite), k=TINV, U=k·u_c.")
    ws["A2"].font = Font(italic=True, color=SUP)
    ws.merge_cells("A2:F2")

    # служебные коэффициенты (строки 17..21)
    ws.cell(17, 1, "Служебные коэффициенты (живые):")
    ws.cell(17, 1).font = HDR_FONT
    ws.cell(18, 1, "u_rel(m) = u(m_net)/m_ИС")
    set_cell(ws, f"B{B_UREL_MASS}", f"={U_MNET}/{M_IS}", num_fmt="0.000000")
    ws.cell(18, 3, "1")
    ws.cell(19, 1, "u_rel(pA/pИС) = √((u_pa/pA)²+(u_pa/pИС)²)")
    set_cell(ws, f"B{B_UREL_PUR}",
             f"=SQRT(({U_PA}/{P_A})^2+({U_PA}/{P_IS})^2)", num_fmt="0.000000")
    ws.cell(19, 3, "1")
    ws.cell(20, 1, "u_y0 = √(s²_res/p)")
    set_cell(ws, f"B{B_UY0}", f"=IF(ISNUMBER('02_Регрессия'!$B${R_S2}),SQRT('02_Регрессия'!$B${R_S2}/{P_INJ}),0)",
             num_fmt="0.000000")
    ws.cell(20, 3, "Y")

    # заголовок таблицы вкладов
    for c, h in enumerate(["Вклад", "u (г/г)", "df", "Тип", "Доля, %", "Источник (НД)"], 1):
        ws.cell(3, c, h)
        ws.cell(3, c).fill = HDR_FILL
        ws.cell(3, c).font = HDR_FONT
        ws.cell(3, c).border = BORDER

    X0 = f"'06_Итог'!$B${I_X0}"
    VARB1 = f"'02_Регрессия'!$B${R_VARB1}"
    VARB0 = f"'02_Регрессия'!$B${R_VARB0}"
    COV = f"'02_Регрессия'!$B${R_COV}"
    B1 = f"'02_Регрессия'!$B${R_B1}"
    DF = f"'02_Регрессия'!$B${R_DFRES}"
    w0 = f"'06_Итог'!$B${I_W0}"
    urepr_w0 = urepr_formula(w0)

    u_model = f"=SQRT((B{B_UY0}^2+{X0}^2*{VARB1}+{VARB0}+2*{X0}*{COV}))/ABS({B1})"
    rows = [
        ("модель (OLS: обратная градуировка)",
         u_model, f"={DF}", "A",
         "u²=(u_y0²+x0²·var_a+var_b+2·x0·cov)/b1²; u_y0²=s²/p — эквивалент интервальной формы"),
        ("эталоны", f"=ABS({X0})*{U_REL_ETAL}", "50", "B",
         "u_rel эталонов (вход 01_Ввод); ν=50"),
        ("чистота (pA, pИС)", f"=ABS({X0})*$B${B_UREL_PUR}", "∞", "B",
         "u_rel(pA/pИС) — формула выше; равномерное; ν=∞"),
        ("весы (навески)", f"=ABS({X0})*$B${B_UREL_MASS}", "∞", "B",
         "u_rel(m)=u(m_net)/m_ИС; СЦЕНАРИЙ; ν=∞; TD-DEV-33"),
        ("прецизионность u_repr (mean2)",
         f"=IF(ISNUMBER({w0}),IF({MODE_CELL}=1,'03_ANOVA'!$B${R_UREPRSTEP},'03_ANOVA'!$B${R_UREPRSMOOTH})/'06_Итог'!$B${I_KG2PCT},\"\")",
         "∞", "B",
         f"√(R²−r²/2)/2,8 (5725-6 4.2.3), табл.1; режим {MODE_CELL}; анти-двойной счёт σr"),
    ]
    for i, (name, u, df, tp, src) in enumerate(rows):
        r = B_ROW_MODEL + i
        ws.cell(r, 1, name)
        set_cell(ws, f"B{r}", u, num_fmt="0.0000000")
        ws.cell(r, 3, df)
        ws.cell(r, 4, tp)
        ws.cell(r, 6, src).alignment = Alignment(wrap_text=True)
        for c in range(1, 7):
            ws.cell(r, c).border = BORDER

    # доли дисперсии (столбец E)
    for i in range(5):
        r = B_ROW_MODEL + i
        ws.cell(r, 5,
                f"=IF(AND(ISNUMBER(B{r}),ISNUMBER($B${B_UC}),$B${B_UC}>0),B{r}^2/$B${B_UC}^2*100,0)").number_format = "0.00"

    # итоги
    totals = [
        (B_UC, "u_c (г/г)", f"=SQRT(SUMSQ(B{B_ROW_MODEL}:B{B_ROW_REPR}))", "0.0000000",
         "√Σu² — корень квадратный из суммы (GUM 6.2.1)"),
        (B_NUEFF, "ν_eff (Welch–Satterthwaite)",
         f"=IF(B{B_UC}=0,NA(),B{B_UC}^4/(B{B_ROW_MODEL}^4/{DF}+B{B_ROW_ETAL}^4/50))", "0.0",
         "ν_eff = u_c⁴/Σ(u_i⁴/ν_i); GUM G.2b (чистота/весы/прецизионность — ν=∞)"),
        (B_K, "k (t_{0.975}(ν_eff))",
         f"=IF(ISNUMBER(B{B_NUEFF}),TINV(0.05,MAX(ROUND(B{B_NUEFF},0),1)),NA())", "0.0000",
         "GUM 6.3/G.4"),
        (B_U, "U (г/г) = k·u_c", f"=B{B_K}*B{B_UC}", "0.0000000", "GUM 6.2.1; P=0.95"),
        (B_UPCT, "U (% мас.) = U·K_G2PCT", f"=B{B_U}*'06_Итог'!$B${I_KG2PCT}", "0.0000",
         "перевод шкалы (контракт разд.9.2)"),
        (B_UREL, "U_rel, % (отн.)",
         f"=IF(OR(NOT(ISNUMBER(B{B_UPCT})),NOT(ISNUMBER({w0})),{w0}=0),"
         f"NA(),B{B_UPCT}/{w0}*100)", "0.00", "U% / w0 · 100"),
    ]
    for r, label, formula, fmt, note in totals:
        ws.cell(r, 1, label).font = Font(bold=True)
        set_cell(ws, f"B{r}", formula, num_fmt=fmt)
        ws.cell(r, 6, note).alignment = Alignment(wrap_text=True)
        for c in range(1, 7):
            ws.cell(r, c).border = BORDER

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 6
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 58
    return ws


# ---------------------------------------------------------------------------
# ЛИСТ 06_Итог
# ---------------------------------------------------------------------------
def fill_result_live(wb, levels, inputs):
    ws = wb.create_sheet("06_Итог")
    X0_f = (f"=IF(ISNUMBER('02_Регрессия'!$B${R_B1}),"
            f"({Y0_CELL}-'02_Регрессия'!$B${R_B0})/'02_Регрессия'!$B${R_B1},NA())")
    ws["A1"] = "РЕЗУЛЬТАТ ОПРЕДЕЛЕНИЯ MTBE (ГОСТ EN 13132-2012) — ЖИВЫЕ ФОРМУЛЫ"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:F1")
    ws["A2"] = ("Все расчётные ячейки — формулы со ссылками на 01_Ввод/02/03/05. "
                "Изменение входов (X/Y, массы, чистота, Y0, p) пересчитывает результат в Excel.")
    ws["A2"].font = Font(italic=True, color=SUP)
    ws.merge_cells("A2:F2")

    ws["A3"] = "K_G2PCT = m_ИС·p_ИС/m_пробы·100"
    set_cell(ws, f"B{I_KG2PCT}",
             f"={M_IS}*{P_IS}/{M_PROBY}*100", num_fmt="0.0000")
    ws["C3"] = "% мас./(г/г)"
    for c in range(1, 4):
        ws.cell(3, c).border = BORDER

    # --- 1. ПАРАМЕТРЫ ---
    hr = 5
    ws.cell(hr, 1, "1. ПАРАМЕТРЫ (ввод, лист 01_Ввод)").font = HDR_FONT
    hr += 1
    for c, h in enumerate(["Параметр", "Значение", "Ед.", "Источник"], 1):
        ws.cell(hr, c, h)
        ws.cell(hr, c).fill = HDR_FILL
        ws.cell(hr, c).font = HDR_FONT
        ws.cell(hr, c).border = BORDER
    param_cells = [
        ("Масса пробы m_пробы, г", f"={M_PROBY}", "0.000", "01_Ввод (поле ввода)"),
        ("Масса ИС в пробу m_ИС, г", f"={M_IS}", "0.000", "01_Ввод (поле ввода)"),
        ("Чистота ИС p_ИС", f"={P_IS}", "0.0000", "01_Ввод; паспорт партии"),
        ("Чистота аналита p_A", f"={P_A}", "0.0000", "01_Ввод"),
        ("Y0 пробы (среднее p инъекций)", f"={Y0_CELL}", "0.000000", "01_Ввод; измерение"),
        ("Число инъекций пробы p", f"={P_INJ}", "0", "01_Ввод; ГОСТ EN 13132 §10.1"),
        ("u(m_net) весы, г", f"={U_MNET}", "0.00000000", "01_Ввод; СЦЕНАРИЙ (TD-DEV-33)"),
        ("u(pA)=u(pИС), доля", f"={U_PA}", "0.000000", "01_Ввод; равномерное /√12"),
        ("u_rel эталонов, доля", f"={U_REL_ETAL}", "0.0000", "01_Ввод; свидетельство"),
    ]
    for i, (name, formula, fmt, src) in enumerate(param_cells):
        r = hr + 1 + i
        ws.cell(r, 1, name)
        set_cell(ws, f"B{r}", formula, num_fmt=fmt)
        ws.cell(r, 4, src)
        for c in range(1, 5):
            ws.cell(r, c).border = BORDER

    # --- 2. РЕЗУЛЬТАТ ---
    r = hr + 1 + len(param_cells) + 1
    ws.cell(r, 1, "2. РЕЗУЛЬТАТ").font = HDR_FONT
    r += 1
    for c, h in enumerate(["Параметр", "Значение", "Ед.", "Источник"], 1):
        ws.cell(r, c, h)
        ws.cell(r, c).fill = HDR_FILL
        ws.cell(r, c).font = HDR_FONT
        ws.cell(r, c).border = BORDER
    res_cells = [
        ("X0 (обратная градуировка), г/г", X0_f, "0.000000",
         "X0 = (Y0 − b0)/b1 (GUM 5.1.2); Y0 из 01_Ввод, b1/b0 из 02_Регрессия"),
        ("w(MTBE) = X0·m_ИС·p_ИС/m_пробы·100, % мас.",
         f"=IF(ISNUMBER(B{I_X0}),B{I_X0}*$B${I_KG2PCT},\"\")",
         "0.00", "ГОСТ EN 13132 §8; K_G2PCT (B3)"),
        ("u(X0) модель, г/г", f"='05_Бюджет'!$B${B_ROW_MODEL}", "0.000000",
         "вклад модели 05_Бюджет"),
        ("u_c результата, г/г", f"='05_Бюджет'!$B${B_UC}", "0.0000000",
         "бюджет 05 (анти-двойной счёт)"),
        ("ν_eff", f"='05_Бюджет'!$B${B_NUEFF}", "0.0", "Welch–Satterthwaite, GUM G.2b"),
        ("k", f"='05_Бюджет'!$B${B_K}", "0.0000", "t_{0.975}(ν_eff), GUM 6.3"),
        ("U, г/г", f"='05_Бюджет'!$B${B_U}", "0.0000000", "k·u_c; = U(x_pred) маркер «Графики»!D4"),
        ("U, % мас.", f"='05_Бюджет'!$B${B_UPCT}", "0.0000", "k·u_c в % мас. (перевод U3)"),
        ("U_rel, % (отн.)", f"='05_Бюджет'!$B${B_UREL}", "0.00", "U/w·100"),
    ]
    for i, (name, formula, fmt, src) in enumerate(res_cells):
        rr2 = r + 1 + i
        ws.cell(rr2, 1, name)
        set_cell(ws, f"B{rr2}", formula, num_fmt=fmt)
        ws.cell(rr2, 4, src).alignment = Alignment(wrap_text=True)
        for c in range(1, 5):
            ws.cell(rr2, c).border = BORDER
        ws.cell(rr2, 4).comment = comment(src)

    # --- 3. ПРОТОКОЛЬНАЯ ЗАПИСЬ ---
    r = r + 1 + len(res_cells) + 1
    ws.cell(r, 1, "3. ПРОТОКОЛЬНАЯ ЗАПИСЬ (ГОСТ EN 13132 §9, точность 0,1 % мас.)").font = HDR_FONT
    proto = (f'="w(MTBE) = "&TEXT(ROUND(B{I_W0},1),"0.0")&" ± "&TEXT(ROUND(B{I_UPCT},1),"0.0")'
             f'&" % мас. (P=0,95, k="&TEXT(ROUND(B{I_K},2),"0.00")&", ν_eff="&TEXT(ROUND(B{I_NUEFF},0),"0")&")"')
    ws.cell(r + 1, 1, proto).font = Font(bold=True, size=12, color="1F4E79")
    ws.cell(r + 1, 1).border = BORDER
    ws.cell(r + 2, 1, "Протокольное округление до 0,1 % мас. (§9 ГОСТ EN 13132); "
                      "U округление — GUM 7.2.6").font = Font(italic=True, color=SUP)

    # --- 4. ТАБЛИЦА ВКЛАДОВ ---
    r = r + 4
    ws.cell(r, 1, "4. ТАБЛИЦА ВКЛАДОВ (формат Word-отчёта v15)").font = HDR_FONT
    r += 1
    for c, h in enumerate(["Составляющая", "u (г/г)", "Доля дисперсии, %", "Источник"], 1):
        ws.cell(r, c, h)
        ws.cell(r, c).fill = HDR_FILL
        ws.cell(r, c).font = HDR_FONT
        ws.cell(r, c).border = BORDER
    contrib = [
        ("Модель (калибровка)", f"='05_Бюджет'!$B${B_ROW_MODEL}", f"='05_Бюджет'!$E${B_ROW_MODEL}",
         "обратная градуировка с var/cov"),
        ("Эталоны (u(x_cal))", f"='05_Бюджет'!$B${B_ROW_ETAL}", f"='05_Бюджет'!$E${B_ROW_ETAL}",
         "свидетельство калибровки стандарта"),
        ("Чистота аналита/ИС", f"='05_Бюджет'!$B${B_ROW_PUR}", f"='05_Бюджет'!$E${B_ROW_PUR}",
         "u(pA/pИС) из входов 01_Ввод"),
        ("Масса навесок (весы)", f"='05_Бюджет'!$B${B_ROW_W}", f"='05_Бюджет'!$E${B_ROW_W}",
         "СЦЕНАРИЙ (TD-DEV-33)"),
        ("Прецизионность u_repr (mean2)", f"='05_Бюджет'!$B${B_ROW_REPR}", f"='05_Бюджет'!$E${B_ROW_REPR}",
         "табл.1 ГОСТ EN 13132; анти-двойной счёт (контракт разд.6, R3)"),
    ]
    for i, (name, u, share, src) in enumerate(contrib):
        rr2 = r + 1 + i
        ws.cell(rr2, 1, name)
        set_cell(ws, f"B{rr2}", u, num_fmt="0.0000000")
        set_cell(ws, f"C{rr2}", share, num_fmt="0.00")
        ws.cell(rr2, 4, src).alignment = Alignment(wrap_text=True)
        for c in range(1, 5):
            ws.cell(rr2, c).border = BORDER
    rr2 = r + 1 + len(contrib)
    ws.cell(rr2, 1, "U = k·u_c").font = Font(bold=True)
    set_cell(ws, f"B{rr2}", f"='05_Бюджет'!$B${B_U}", num_fmt="0.0000000")
    ws.cell(rr2, 4, "GUM 6.2.1/6.3; ν_eff по Welch–Satterthwaite")
    for c in range(1, 5):
        ws.cell(rr2, c).border = BORDER

    # --- 5. ПРИМЕЧАНИЯ ---
    r = rr2 + 2
    ws.cell(r, 1, "5. ПРИМЕЧАНИЯ").font = HDR_FONT
    notes = [
        "· Режим u_repr: 01_Ввод!B29. По умолчанию 2=smooth (кусочно-линейная по узлам R-05, "
        "как ядро anova_gum.py) — U=0.0862 г/г; режим 1=step (ступенчатая, РМГ 61 §5.5) — "
        "на границах поддиапазонов табл.1 U может отличаться (для w0=3.97: ~0.0795 г/г).",
        "· Допущение: «σR принято по табл.1 ГОСТ EN 13132 (N межлаб. округление 2,8); σL "
        "данных лаборатории не оценивалась за отсутствием межсерийных данных» (17025 §7.6).",
        "· Статус: ПРЕДВАРИТЕЛЬНО — весы/партии не подтверждены (TD-DEV-33); протокол "
        "блокируется до паспортных данных.",
        "· Книга ЖИВАЯ: формулы пересчитываются Excel-ем при изменении входов 01_Ввод; "
        "кэша значений нет — программное чтение без пересчёта через ядро anova_gum.py.",
    ]
    for i, note in enumerate(notes):
        ws.cell(r + 1 + i, 1, note).font = Font(italic=True, color=SUP)
        ws.cell(r + 1 + i, 1).alignment = Alignment(wrap_text=True)
        ws.merge_cells(f"A{r + 1 + i}:F{r + 1 + i}")
        ws.row_dimensions[r + 1 + i].height = 30
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 58
    return ws


# ---------------------------------------------------------------------------
# ЛИСТ ГРАФИКИ
# ---------------------------------------------------------------------------
def fill_charts_live(wb, levels, inputs, ws_reg, prof_range):
    DATA_FIRST, DATA_LAST, GRID_PT = prof_range[0], prof_range[1], prof_range[2]
    PROF_FIRST, PROF_LAST = prof_range[3], prof_range[4]
    gch = wb.create_sheet("Графики")
    gch["A1"] = "ГРАФИЧЕСКАЯ ДИАГНОСТИКА МОДЕЛИ — ЖИВЫЕ ДИАПАЗОНЫ"
    gch["A1"].font = TITLE_FONT
    gch["A2"] = ("Диаграммы построены на живых диапазонах листов 02_Регрессия и 04_Профиль U(X); "
                 "обновляются при пересчёте книги (в Excel — автоматически).")
    gch["A2"].font = Font(italic=True, color=SUP)
    gch["A3"] = ("Примечание: U(x) — живые формулы =k·u_c; при программном чтении без "
                 "пересчёта использовать ядро anova_gum.py.")
    gch["A3"].font = Font(italic=True, color=SUP)

    gch["A4"] = "Маркер x_pred (результат пробы)"
    gch["A4"].font = Font(bold=True)
    gch["B4"] = "x_pred и U(x_pred) — ссылки на 06_Итог (живые формулы)"
    gch["B4"].font = Font(bold=True, color="C00000")
    set_cell(gch, "C4", f"='06_Итог'!$B${I_X0}", num_fmt="0.000000")
    set_cell(gch, "D4", f"='06_Итог'!$B${I_U}", num_fmt="0.000000")
    gch["A5"] = "Нули линии 0 (график остатков) — на 02_Регрессия (столбец G); " \
                "xVal/yVal серии на ОДНОМ листе (Excel-правило)"
    gch["A5"].font = Font(italic=True, color=SUP)

    # --- Диаграмма 1: Калибровочная прямая ---
    ch1 = ScatterChart()
    ch1.title = "Калибровочная прямая: точки, линия Y=b1·X+b0, полоса ŷ±k·u(ŷ)"
    ch1.style = 13
    ch1.x_axis.title = "X, г/г"
    ch1.y_axis.title = "Y (отношение площадей)"
    ch1.x_axis.delete = False
    ch1.y_axis.delete = False
    ch1.height = 10
    ch1.width = 16
    xdata1 = Reference(ws_reg, min_col=1, min_row=DATA_FIRST, max_row=DATA_LAST)
    ydata1 = Reference(ws_reg, min_col=2, min_row=DATA_FIRST, max_row=DATA_LAST)
    s_pts = Series(ydata1, xdata1, title="точки (X, Y_ср)")
    s_pts.marker.symbol = "circle"
    s_pts.marker.size = 6
    s_pts.graphicalProperties.line.noFill = True
    ch1.series.append(s_pts)
    xline1 = Reference(ws_reg, min_col=1, min_row=GRID_PT, max_row=GRID_PT + 99)
    yline1 = Reference(ws_reg, min_col=3, min_row=GRID_PT, max_row=GRID_PT + 99)
    s_line = Series(yline1, xline1, title="Y = b1·X + b0")
    s_line.graphicalProperties.line.solidFill = "C00000"
    s_line.graphicalProperties.line.width = 22000
    s_line.marker.symbol = "none"
    ch1.series.append(s_line)
    yup1 = Reference(ws_reg, min_col=6, min_row=DATA_FIRST, max_row=DATA_LAST)
    ydn1 = Reference(ws_reg, min_col=5, min_row=DATA_FIRST, max_row=DATA_LAST)
    s_up = Series(yup1, xdata1, title="ŷ + k·u(ŷ)")
    s_up.graphicalProperties.line.solidFill = "9DC3E6"
    s_up.graphicalProperties.line.width = 12000
    s_up.graphicalProperties.line.dashStyle = "dash"
    s_up.marker.symbol = "none"
    ch1.series.append(s_up)
    s_dn = Series(ydn1, xdata1, title="ŷ − k·u(ŷ)")
    s_dn.graphicalProperties.line.solidFill = "9DC3E6"
    s_dn.graphicalProperties.line.width = 12000
    s_dn.graphicalProperties.line.dashStyle = "dash"
    s_dn.marker.symbol = "none"
    ch1.series.append(s_dn)
    gch.add_chart(ch1, "C2")

    # --- Диаграмма 2: Профиль U(x) ---
    ch2 = ScatterChart()
    ch2.title = "Профиль неопределённости U(x) = k·u_c(x) и u_c(x)"
    ch2.style = 13
    ch2.x_axis.title = "X, г/г"
    ch2.y_axis.title = "неопределённость, г/г"
    ch2.x_axis.delete = False
    ch2.y_axis.delete = False
    ch2.height = 10
    ch2.width = 16
    ws_prof = wb["04_Профиль U(X)"]
    xdata2 = Reference(ws_prof, min_col=2, min_row=PROF_FIRST, max_row=PROF_LAST)
    yuc2 = Reference(ws_prof, min_col=5, min_row=PROF_FIRST, max_row=PROF_LAST)
    yu2 = Reference(ws_prof, min_col=8, min_row=PROF_FIRST, max_row=PROF_LAST)
    s_uc = Series(yuc2, xdata2, title="u_c(x)")
    s_uc.graphicalProperties.line.solidFill = "2E75B6"
    s_uc.graphicalProperties.line.width = 18000
    s_uc.marker.symbol = "none"
    ch2.series.append(s_uc)
    s_u = Series(yu2, xdata2, title="U(x)")
    s_u.graphicalProperties.line.solidFill = "C00000"
    s_u.graphicalProperties.line.width = 22000
    s_u.marker.symbol = "none"
    ch2.series.append(s_u)
    # Маркер x_pred: xVal/yVal на листе 04_Профиль U(X) (K4/L4) — Excel-правило
    # «серия — один лист»; иначе Excel рендерит X как 1..n.
    xm = Reference(ws_prof, min_col=11, min_row=4, max_row=4)
    ym = Reference(ws_prof, min_col=12, min_row=4, max_row=4)
    s_mark = Series(ym, xm, title="x_pred")
    s_mark.marker.symbol = "square"
    s_mark.marker.size = 8
    s_mark.graphicalProperties.line.noFill = True
    ch2.series.append(s_mark)
    gch.add_chart(ch2, "C26")

    # --- Диаграмма 3: Остатки ---
    ch3 = ScatterChart()
    ch3.title = "Остатки регрессии: e = Y_ср − (b1·X + b0)"
    ch3.style = 13
    ch3.x_axis.title = "X, г/г"
    ch3.y_axis.title = "остаток e"
    ch3.x_axis.delete = False
    ch3.y_axis.delete = False
    ch3.height = 10
    ch3.width = 16
    xdata3 = Reference(ws_reg, min_col=1, min_row=DATA_FIRST, max_row=DATA_LAST)
    yres3 = Reference(ws_reg, min_col=4, min_row=DATA_FIRST, max_row=DATA_LAST)
    s_res = Series(yres3, xdata3, title="e = Y_ср − (b1·X + b0)")
    s_res.marker.symbol = "diamond"
    s_res.marker.size = 6
    s_res.graphicalProperties.line.noFill = True
    ch3.series.append(s_res)
    # Линия нуля: yVal на 02_Регрессия (столбец G, константа 0) ВМЕСТЕ с xVal
    # (столбец A) — Excel-правило «серия — один лист» (иначе X -> 1..n).
    yzero3 = Reference(ws_reg, min_col=7, min_row=DATA_FIRST, max_row=DATA_LAST)
    s_zero = Series(yzero3, xdata3, title="0 (линия нуля)")
    s_zero.graphicalProperties.line.solidFill = "808080"
    s_zero.graphicalProperties.line.width = 10000
    s_zero.marker.symbol = "none"
    ch3.series.append(s_zero)
    gch.add_chart(ch3, "C50")

    gch.column_dimensions["A"].width = 30
    gch.column_dimensions["B"].width = 64
    gch.column_dimensions["C"].width = 12
    gch.column_dimensions["D"].width = 12
    gch.column_dimensions["H"].width = 10
    return gch


# ---------------------------------------------------------------------------
# ЛИСТ ИНСТРУКЦИЯ + ЛИСТ ЯДРО_СВЕРКА
# ---------------------------------------------------------------------------
def fill_instruction(wb):
    ws = wb.create_sheet("Инструкция")
    ws["A1"] = "ИНСТРУКЦИЯ ПО ЖИВОЙ КНИГЕ EN13132_v16_LIVE.xlsx"
    ws["A1"].font = TITLE_FONT
    lines = [
        ("Как работает книга", "Каждая расчётная ячейка — формула Excel со ссылками на входы "
         "листа «01_Ввод». Формулы не содержат кэша значений — при открытии в Excel "
         "(или LibreOffice) они пересчитываются автоматически."),
        ("Что менять", "Все входы — на листе «01_Ввод»: X (г/г), инъекции Y1..Y3, массы, чистоты, "
         "Y0 пробы, число инъекций пробы p, u(m_net), u(pA)=u(pИС), u_rel эталонов, режим u_repr."),
        ("Что пересчитается", "После изменения входов Excel автоматически пересчитает: "
         "02_Регрессия (OLS, LoF), 03_ANOVA (σr/σR/σL, ANOVA-таблица), 04_Профиль U(X) (200 точек), "
         "05_Бюджет, 06_Итог (X0, w, U, протокол), графики."),
        ("Режим u_repr", "Ячейка 01_Ввод!B29: 1=step (ступенчатая по табл.1 ГОСТ EN 13132, РМГ 61 §5.5), "
         "2=smooth (кусочно-линейная по узлам R-05, как ядро anova_gum.py). По умолчанию 2 — "
         "совпадает с книгой CALC и ядром (U=0.0862 г/г)."),
        ("Верификация чисел", "Лист «Ядро_сверка» (скрытый) содержит значения ядра anova_gum.py "
         "параллельно — для ручной сверки. Автоматический верификатор "
         "verify_en13132_v16_charts.py пересчитывает формулы библиотекой formulas и сверяет с ядром."),
        ("Примечание", "Статус ПРЕДВАРИТЕЛЬНО (TD-DEV-33): весы/партии не подтверждены — протокол "
         "блокируется до паспортных данных."),
    ]
    r = 3
    for title, body in lines:
        ws.cell(r, 1, title).font = Font(bold=True)
        ws.cell(r, 1).alignment = Alignment(vertical="top")
        ws.cell(r, 2, body).alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(r, 1).border = BORDER
        ws.cell(r, 2).border = BORDER
        ws.row_dimensions[r].height = 40
        r += 1
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 90
    return ws


def fill_kernel_check(wb, res):
    ws = wb.create_sheet("Ядро_сверка")
    ws["A1"] = "DEV: значения ядра anova_gum.py (НЕ формулы) — только для сверки с живыми формулами"
    ws["A1"].font = Font(bold=True, color="808080")
    ws["A2"] = "Эти числа генерируются скриптом при сборке; книга считается живой по формулам листов 01..06."
    ws["A2"].font = Font(italic=True, color=SUP)
    hdr = 4
    for c, h in enumerate(["Параметр", "Ядро", "Ед."], 1):
        ws.cell(hdr, c, h)
        ws.cell(hdr, c).fill = HDR_FILL
        ws.cell(hdr, c).font = HDR_FONT
    rows = [
        ("b1", res["SNAP"]["a"], "Y/X"),
        ("b0", res["SNAP"]["b"], "Y"),
        ("R²", res["SNAP"]["r2"], "1"),
        ("s²_res", res["SNAP"]["s2_res"], "Y²"),
        ("X0", res["X0"], "г/г"),
        ("w0", res["W0"], "% мас."),
        ("σr", res["SIGMA_R"], "Y"),
        ("u_repr(w0)", res["u_repr_w0"], "% мас."),
        ("σL", res["sigma_L"], "% мас."),
        ("u_c", res["U_C_RES"], "г/г"),
        ("ν_eff", res["NU_EFF_RES"], "1"),
        ("k", res["K_RES"], "1"),
        ("U", res["U_RES_G"], "г/г"),
        ("U%", res["U_RES_PCT"], "% мас."),
        ("U_rel", res["U_REL"], "1"),
    ]
    for i, (name, val, unit) in enumerate(rows):
        r = hdr + 1 + i
        ws.cell(r, 1, name)
        ws.cell(r, 2, float(val)).number_format = "0.00000000"
        ws.cell(r, 3, unit)
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 18
    ws.sheet_state = "hidden"
    return ws


# ---------------------------------------------------------------------------
# СБОРКА
# ---------------------------------------------------------------------------
def build_live_workbook(levels, inputs, mode, out_path, with_kernel_check=True):
    levels4 = []
    for item in levels:
        x, ys, *rest = item
        mA = rest[0] if len(rest) > 0 else None
        mIS = rest[1] if len(rest) > 1 else None
        levels4.append((float(x), list(ys), mA, mIS))
    levels2 = [(x, ys) for x, ys, _mA, _mIS in levels4]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    write_input_sheet_live(wb, levels4, inputs, mode)
    ws_reg, DATA_FIRST, DATA_LAST, GRID_PT = fill_regression_live(wb, levels2, inputs)
    fill_anova_live(wb, levels2, inputs)
    ws_prof, PROF_FIRST, PROF_LAST = fill_profile_live(wb, levels2, inputs)
    fill_budget_live(wb, levels2, inputs)
    fill_result_live(wb, levels2, inputs)
    fill_charts_live(wb, levels2, inputs, ws_reg,
                     (DATA_FIRST, DATA_LAST, GRID_PT, PROF_FIRST, PROF_LAST))
    fill_instruction(wb)

    res = None
    if with_kernel_check:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", RuntimeWarning)
            res = bld.compute_calibration(levels2, inputs, "smooth" if mode == 2 else "step")
        fill_kernel_check(wb, res)

    bld.assert_chart_series_sheets(wb)  # Excel-правило: xVal/yVal серии — один лист
    wb.save(out_path)
    return wb, res, (DATA_FIRST, DATA_LAST, GRID_PT, PROF_FIRST, PROF_LAST)


def load_existing_inputs(path):
    if not os.path.exists(path):
        return None, None, 2
    try:
        wb = openpyxl.load_workbook(path, data_only=False)
        ws = wb["01_Ввод"]
        levels = bld.read_levels(ws)
        inputs = bld.read_inputs(ws)
        mode = 2
        r = None
        for row in range(1, ws.max_row + 1):
            v = ws.cell(row, 1).value
            if isinstance(v, str) and v.startswith("Режим u_repr"):
                r = row
                break
        if r is not None and isinstance(ws.cell(r, 2).value, (int, float)):
            mode = int(ws.cell(r, 2).value)
        if levels is not None and inputs is not None:
            return levels, inputs, mode
    except Exception as exc:  # noqa: BLE001
        print(f"  (предупреждение: существующая книга не перечитана — {exc})")
    return None, None, 2


def main():
    levels, inputs, mode = load_existing_inputs(OUT_XLSX)
    source = "EN13132_v16_LIVE.xlsx (01_Ввод, существующая книга)"
    if levels is None:
        v15 = bld.load_v15_levels(XLSX_V15)
        levels = [(x, ys, mA, mIS) for x, ys, mA, mIS in v15]
        inputs = bld.default_inputs(levels)
        mode = 2
        source = "EN13132_v15_MTBE.xlsx (01_Ввод)"
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", RuntimeWarning)
        wb, res, prof_range = build_live_workbook(levels, inputs, mode, OUT_XLSX)
    print("Живая книга сохранена:", OUT_XLSX)
    print("Источник данных:", source)
    print("Листы:", wb.sheetnames)
    print("Режим u_repr:", "smooth" if mode == 2 else "step")
    print("Уровней:", len(levels), "| p пробы:", inputs["n_inj"])
    if res is not None:
        print("--- Значения ядра (для сверки) ---")
        print("a=%.6f b=%.6f R2=%.5f s2=%.3e" % (res["SNAP"]["a"], res["SNAP"]["b"],
                                                  res["SNAP"]["r2"], res["SNAP"]["s2_res"]))
        print("X0=%.6f г/г, w0=%.3f %%мас" % (res["X0"], res["W0"]))
        print("u_c=%.6f, nu_eff=%.1f, k=%.4f, U=%.6f г/г = %.4f %%мас"
              % (res["U_C_RES"], res["NU_EFF_RES"], res["K_RES"], res["U_RES_G"], res["U_RES_PCT"]))


if __name__ == "__main__":
    main()
