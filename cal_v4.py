"""
cal_v4.py — генератор метрологической калибровочной модели (EURACHEM/CITAC)

Наследует cal_v3.py и закрывает технические долги TD-01…TD-06:

  TD-01  x=0 в WLS: при весах 1/x, 1/x² нулевой уровень получает ПУСТОЙ вес
         (не #N/A!) и корректно исключается из взвешенных сумм через
         индикатор участия I (столбец E). ОМНК по-прежнему использует все
         точки (контроль b0 по холостой пробе). Валидация проверяет, что
         n_eff = n − (число нулевых уровней).
  TD-02  Ковариационная матрица WLS — полные формулы (XᵀWX)⁻¹:
         var(b0)_w = φ_w·Σw²x²/(Σw·Sxx_w), cov = −φ_w·Σwx/(Σw·Sxx_w);
         добавлен коэффициент дисперсии D = φ_w/s²_ОМНК (проверка модели
         дисперсии; PASS при 0.5 < D < 2).
  TD-03  Новый лист «Диагностика»: Lack-of-Fit тест по повторностям
         (SS_PE, SS_LoF, F, F_crit через FINV) для ОМНК и активного WLS;
         анализ остатков WLS (корреляция |e|·√w с x; ожидание ≈ 0).
  TD-04  Коэффициент охвата k по Стьюденту (TINV) с ν_eff по Уэлчу–
         Саттертуэйту (модель + эталоны, ν_эталонов = 50). U = k·u_c.
  TD-05  Все расчётные диапазоны — динамические: SUMIF/SUMPRODUCT со
         множителем-индикатором участия; статистики не зависят от пустых
         строк таблицы (до границы 108).
  TD-06  Внешний валидатор validate_v4.py: numpy-эталон WLS/ОМНК/LoF/ν_eff,
         численная валидация (конечные разности) и MC-покрытие; статус
         пишется в лист «Валидация».

Листы: Ввод | Регрессия | Взвешенная регрессия | Диагностика |
       Неопределенность | Неопределенность по точкам | Профиль U(x) | Валидация
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import LineChart, ScatterChart, Reference, Series
from openpyxl.workbook.defined_name import DefinedName

import datetime
import os
import subprocess
import sys

# TD-09/12: центральная конфигурация (без магических чисел) и метаданные версии.
CONFIG = dict(
    VERSION="5.5.0",
    MAX_POINTS=100,       # макс. строк калибровки (лимит TD-07)
    GRID_N=200,           # точек в профиле U(x)
    FIRST_DATA=9,         # первая строка данных на «Вводе»
    STD_DF=50,            # степени свободы эталонов (Welch–Satterthwaite)
    EPS=1e-12,            # защита от деления на ~0
    D_LOW=0.5,            # границы коэффициента дисперсии D
    D_HIGH=2.0,
)


def get_git_commit():
    try:
        r = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            capture_output=True, text=True,
            cwd=os.path.dirname(os.path.abspath(__file__)),
        )
        return r.stdout.strip() or "n/a"
    except Exception:
        return "n/a"


def create_metrology_excel_v4(output_path="Metrology_Core_EURACHEM_v4.xlsx"):
    wb = Workbook()

    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="4F81BD")
    input_fill = PatternFill("solid", fgColor="E2EFDA")
    calc_fill = PatternFill("solid", fgColor="DDEBF7")
    pass_fill = PatternFill("solid", fgColor="C6EFCE")
    fail_fill = PatternFill("solid", fgColor="FFC7CE")
    warn_fill = PatternFill("solid", fgColor="FFEB9C")

    def style_header(sh, row, cols):
        for c in range(1, cols + 1):
            cell = sh.cell(row, c)
            cell.fill = header_fill
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center",
                                       wrap_text=True)
            cell.border = border

    # ============================================================
    # TD-07: именованные диапазоны (динамические области, без 9:108)
    # ============================================================
    L = CONFIG["MAX_POINTS"]
    F = CONFIG["FIRST_DATA"]
    LAST = F + L - 1          # 108
    DLAST = F + L - 2         # 107 (таблица Диагностики начинается с 8)
    WD = 9 + L - 1            # 108 (WLS-таблица с 9)
    DD = 8 + L - 1            # 107 (Диагностика с 8)
    PD = 4 + CONFIG["GRID_N"] - 1   # 203 (Профиль с 4)
    Cfg = dict(L=L, F=F, LAST=LAST, DLAST=DLAST, WD=WD, DD=DD, PD=PD)

    # Ячейка итога u(x_cal)/x на листе «Стандарты» (вычисляется заранее)
    STD_FIRST = 17                      # первая строка таблицы эталонов
    STD_LAST = STD_FIRST + L - 1        # 116
    STD_ITOG = STD_LAST + 8             # 124
    STD_UREL = f"'Стандарты'!$B${STD_ITOG}"
    STD_NU = f"'Стандарты'!$B${STD_LAST + 2}"   # ν_эталонов (R-09), ячейка 118

    def add_name(name, ref):
        wb.defined_names.add(DefinedName(name, attr_text=ref))

    add_name("CAL_X", f"'Ввод'!$B${F}:$B${LAST}")     # концентрации
    add_name("CAL_Y", f"'Ввод'!$F${F}:$F${LAST}")     # средние отклики
    add_name("CAL_E2", f"'Ввод'!$I${F}:$I${LAST}")    # e²
    add_name("CAL_SY", f"'Ввод'!$J${F}:$J${LAST}")    # s(y_i) повторн.
    add_name("CAL_XBACK", f"'Ввод'!$M${F}:$M${LAST}")  # x_back (A-01)
    add_name("CAL_DELTA", f"'Ввод'!$N${F}:$N${LAST}")  # Δ% back (A-01)
    add_name("WLS_X", f"'Взвешенная регрессия'!$A$9:$A${WD}")
    add_name("WLS_Y", f"'Взвешенная регрессия'!$B$9:$B${WD}")
    add_name("WLS_W", f"'Взвешенная регрессия'!$C$9:$C${WD}")
    add_name("WLS_WY", f"'Взвешенная регрессия'!$D$9:$D${WD}")
    add_name("WLS_I", f"'Взвешенная регрессия'!$E$9:$E${WD}")
    add_name("WLS_XY", f"'Взвешенная регрессия'!$G$9:$G${WD}")
    add_name("WLS_WX2", f"'Взвешенная регрессия'!$H$9:$H${WD}")
    add_name("WLS_W2X2", f"'Взвешенная регрессия'!$J$9:$J${WD}")
    add_name("WLS_WE2", f"'Взвешенная регрессия'!$L$9:$L${WD}")
    add_name("WLS_WOLS2", f"'Взвешенная регрессия'!$O$9:$O${WD}")  # w·e_ols² (A-06)
    add_name("DG_X", f"'Диагностика'!$A$8:$A${DD}")
    add_name("DG_NI", f"'Диагностика'!$C$8:$C${DD}")
    add_name("DG_PE", f"'Диагностика'!$D$8:$D${DD}")
    add_name("DG_LOF_OLS", f"'Диагностика'!$G$8:$G${DD}")
    add_name("DG_LOF_WLS", f"'Диагностика'!$H$8:$H${DD}")
    add_name("DG_ABS", f"'Диагностика'!$J$8:$J${DD}")   # |e|·√w
    add_name("DG_XP", f"'Диагностика'!$K$8:$K${DD}")    # x участн.
    add_name("DG_J2", f"'Диагностика'!$L$8:$L${DD}")
    add_name("DG_K2", f"'Диагностика'!$M$8:$M${DD}")
    add_name("DG_JK", f"'Диагностика'!$N$8:$N${DD}")
    add_name("PROF_U", f"'Профиль U(x)'!$F$4:$F${PD}")

    # ============================================================
    # ЛИСТ 1. ВВОД
    # Данные: B9:B108 x; C:E отклики; F среднее; J s(y_i).
    # Параметры: B3 y_obs, B4 p, B5 u_rel эталонов, B6 режим весов,
    # B7 доверительная вероятность.
    # ============================================================
    ws = wb.active
    ws.title = "Ввод"

    ws["A1"] = "ПАРАМЕТРЫ ПРОБЫ И КАЛИБРОВОЧНЫЕ ДАННЫЕ"
    ws["A1"].font = Font(bold=True, size=15)

    # B3..B7 — трассируемые: y_obs из повторностей пробы (B114:B116),
    # u_rel из листа «Стандарты» (B15). Значения защищены формулами,
    # ввод осуществляется в подсвеченных областях.
    params = [
        ("Средний отклик пробы (y_obs)",
         "=IF(COUNT('Ввод'!B113:B115)>=2,AVERAGE('Ввод'!B113:B115),NA())"),
        ("Число повторностей пробы (p)",
         "=COUNT('Ввод'!B113:B115)"),
        ("Относительная стандартная неопределённость эталонов u(x_cal)/x",
         "=" + STD_UREL),
        ("Режим весов WLS (1=равные, 2=1/x, 3=1/x²)", 3),
        ("Доверительная вероятность для U", 0.95),
    ]
    for r, (name, value) in enumerate(params, 3):
        ws.cell(r, 1, name)
        cell = ws.cell(r, 2, value)
        cell.fill = calc_fill if isinstance(value, str) and value.startswith("=") else input_fill
    ws["B5"].number_format = "0.00%"
    ws["B7"].number_format = "0.00%"

    # В-01: метаданные калибровки (ISO 17025) — в столбцах P:Q, НЕ в зоне данных (A9:A108)!
    ws["P2"] = "МЕТАДАННЫЕ КАЛИБРОВКИ (В-01)"
    ws["P2"].font = Font(bold=True)
    meta = [
        ("ID калибровки (авто)", '=TEXT(TODAY(),"YYYYMMDD")&"-"&Q6'),
        ("Оператор", ""),
        ("Прибор (серийный №)", ""),
        ("Партия реактивов/эталонов", ""),
        ("Дата калибровки", "=TODAY()"),
        ("Температура помещения, °C", ""),
        ("Срок действия калибровки", ""),
    ]
    for i, (name, value) in enumerate(meta):
        r = 3 + i
        ws.cell(r, 16, name)   # P
        cell = ws.cell(r, 17, value)  # Q
        cell.fill = calc_fill if isinstance(value, str) and value.startswith("=") else input_fill
    ws["Q3"].number_format = "0"
    ws["Q7"].number_format = "DD.MM.YYYY"
    ws["Q9"].number_format = "DD.MM.YYYY"

    dv_w = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False)
    ws.add_data_validation(dv_w)
    dv_w.add(ws["B6"])
    dv_p = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="1")
    ws.add_data_validation(dv_p)
    dv_p.add(ws["B4"])
    dv_u = DataValidation(type="decimal", operator="between",
                          formula1="0", formula2="1")
    ws.add_data_validation(dv_u)
    dv_u.add(ws["B5"])
    dv_u.add(ws["B7"])

    headers = [
        "№", "Концентрация x", "Отклик 1 y", "Отклик 2 y", "Отклик 3 y",
        "Среднее y", "ŷ (ОМНК)", "Остаток e", "e²", "s(y_i) повт.",
        "x_back", "Δ% back",  # A-01
    ]
    for c, h in enumerate(headers, 1):
        ws.cell(8, c, h)
    style_header(ws, 8, len(headers))

    cal = [
        [1, 0.0, 0.002, 0.001, 0.003],
        [2, 0.2, 0.055, 0.058, 0.054],
        [3, 0.4, 0.110, 0.112, 0.108],
        [4, 0.6, 0.165, 0.162, 0.168],
        [5, 0.8, 0.220, 0.225, 0.218],
        [6, 1.0, 0.275, 0.272, 0.278],
    ]
    for r, row in enumerate(cal, 9):
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)
        for c in range(2, 6):
            ws.cell(r, c).fill = input_fill
        ws.cell(r, 10, f'=IF(B{r}="","",IF(COUNT(C{r}:E{r})>=2,'
                       f'STDEV(C{r}:E{r}),NA()))').fill = calc_fill

    for r in range(9, 109):
        ws.cell(r, 6, f'=IF(B{r}="","",AVERAGE(C{r}:E{r}))').fill = calc_fill
        ws.cell(r, 7, f'=IF(B{r}="","",'
                      f"'Регрессия'!$B$7*B{r}+'Регрессия'!$B$8)").fill = calc_fill
        ws.cell(r, 8, f'=IF(F{r}="","",F{r}-G{r})').fill = calc_fill
        ws.cell(r, 9, f'=IF(H{r}="","",H{r}^2)').fill = calc_fill
        # A-01: back-calculation x_back = (F − b0)/b1, Δ% = |x_back − x|/x (ISO 11095)
        ws.cell(r, 13,
                f'=IF(OR(B{r}="",F{r}=""),"",(F{r}-\'Регрессия\'!$B$8)'
                f'/\'Регрессия\'!$B$7)').fill = calc_fill
        ws.cell(r, 14,
                f'=IF(OR(M{r}="",B{r}=0),"",ABS(M{r}-B{r})/B{r})').fill = calc_fill
        # K,L: доверительная полоса ŷ ± k·u(ŷ_i) для графика диапазона неопределённости
        # «Неопределенность по точкам» начинается со строки 6 (эквивалент Ввода 9), сдвиг -3.
        ws.cell(r, 11, f"=IF(B{r}=\"\",\"\",G{r}+'Неопределенность'!$B$12*"
                       f"'Неопределенность по точкам'!H{r-3})").fill = calc_fill
        ws.cell(r, 12, f"=IF(B{r}=\"\",\"\",G{r}-'Неопределенность'!$B$12*"
                       f"'Неопределенность по точкам'!H{r-3})").fill = calc_fill

    # Повторности пробы (трассируемый источник y_obs): строки 112-114.
    ws["A111"] = "ОТКЛИК ПРОБЫ — повторности (ввод)"
    ws["A111"].font = Font(bold=True, size=12)
    for c, h in enumerate(["№", "Повторность y_obs"], 1):
        ws.cell(112, c, h)
    style_header(ws, 112, 2)
    for i in range(3):
        r = 113 + i
        ws.cell(r, 1, i + 1)
        ws.cell(r, 2, "").fill = input_fill
    # значения по умолчанию (чтобы модель была рабочей сразу после генерации)
    for i, v in enumerate([0.148, 0.152, 0.150]):
        ws.cell(113 + i, 2, v).fill = input_fill
    ws["A116"] = ("y_obs = среднее повторов (B3); p = число повторов (B4). "
                  "Стандартная неопределённость u(x_cal)/x — на листе «Стандарты».")
    ws["A116"].font = Font(italic=True, color="808080")

    ws.freeze_panes = "A9"

    # ============================================================
    # ЛИСТ 1Б. СТАНДАРТЫ (трассируемая оценка u(x_cal)/x, с/без корреляции)
    # Каждый эталон регрессии расписан: концентрация (связана с «Ввод»),
    # источник/сертификат, u_cert, распределение, множитель, u_rel.
    # Маточный раствор: концентрация, u_rel, и эталоны из него разводятся
    # (общий источник → корреляция). Итог B37 = u(x_cal)/x на «Ввод» B5.
    # ============================================================
    std = wb.create_sheet("Стандарты")
    std["A1"] = "ОЦЕНКА НЕОПРЕДЕЛЁННОСТИ СТАНДАРТОВ (u(x_cal)/x)"
    std["A1"].font = Font(bold=True, size=15)
    std["A2"] = ("Каждый стандарт калибровочной прямой расписан: концентрация, источник, "
                 "сертификатная неопределённость, распределение. Эталоны, приготовленные из "
                 "общего маточного раствора, коррелируют — это учитывается. Итог u(x_cal)/x "
                 "используется на листе «Ввод» (B5).")
    std["A2"].font = Font(italic=True, color="808080")

    # --- Блок 0: маточный раствор (источник корреляции) ---
    std["A4"] = "0. МАТОЧНЫЙ РАСТВОР (источник эталонов, задаёт корреляцию)"
    std["A4"].font = Font(bold=True, color="C00000")
    std["A5"] = "Наименование маточного раствора"
    std["B5"] = "Маточник M1"
    std["B5"].fill = input_fill
    std["A6"] = "Концентрация маточника C_stock"
    std["B6"] = 100.0
    std["B6"].fill = input_fill
    std["A7"] = "Единицы концентрации"
    std["B7"] = "мг/л"
    std["B7"].fill = input_fill
    std["A8"] = "u(C_stock) из сертификата (абс.)"
    std["B8"] = 1.0
    std["B8"].fill = input_fill
    std["A9"] = "Распределение маточника"
    std["B9"] = "нормальное"
    std["B9"].fill = input_fill
    std["A10"] = "Множитель k / √ для маточника"
    std["B10"] = 2
    std["B10"].fill = input_fill
    std["A11"] = "u_rel маточника (расчёт)"
    std["B11"] = ("=IF(AND(B6<>0,ISNUMBER(B8)),"
                  "IF(B9=\"нормальное\",B8/B6/B10,"
                  "IF(B9=\"прямоугольное\",B8/B6/SQRT(3),"
                  "IF(B9=\"треугольное\",B8/B6/SQRT(6),B8/B6))),NA())")
    std["B11"].fill = calc_fill
    std["B11"].number_format = "0.00%"
    std["A12"] = ("Способы разведения: 1=прямое (C_i=C_stock·k_i), "
                  "2=объёмное (C_i=C_stock·V_i/V_колб)")
    std["B12"] = 1
    std["B12"].fill = input_fill
    std["A13"] = "u_rel процедуры разведения (пипетки, колбы)"
    std["B13"] = 0.005
    std["B13"].fill = input_fill
    std["B13"].number_format = "0.00%"

    # --- Блок 1: эталоны регрессии (расписать каждый) ---
    std["A15"] = "1. ЭТАЛОНЫ КАЛИБРОВОЧНОЙ ПРЯМОЙ (каждый стандарт)"
    std["A15"].font = Font(bold=True)
    sh_std = ["№", "Конц. эталона x_cal", "Связан с листом «Ввод»",
              "Источник/сертификат", "Приготовлен из маточника?",
              "Коэффициент разведения k_i", "u(x) серт. (абс.)",
              "Распределение", "Множитель k/√", "u(x) станд. (абс.)",
              "u_rel = u/x", "Участвует в WLS?", "ν_i (из сертификата)"]
    for c, h in enumerate(sh_std, 1):
        std.cell(16, c, h)
    style_header(std, 16, len(sh_std))

    # Связываем каждый эталон регрессии с листом «Ввод» (CAL_X, строки 9..LAST)
    # Таблица стандартов: строки 17..(17+MAX_POINTS-1)
    for i in range(CONFIG["MAX_POINTS"]):
        r = STD_FIRST + i
        src_r = CONFIG["FIRST_DATA"] + i          # строка на «Вводе»
        # №
        std.cell(r, 1, i + 1)
        # Концентрация — связана с Ввод
        std.cell(r, 2, f'=IF(\'Ввод\'!B{src_r}="","",\'Ввод\'!B{src_r})').fill = calc_fill
        # Ссылка на Ввод (текст)
        std.cell(r, 3, f"'Ввод'!B{src_r}").fill = calc_fill
        # Источник (ввод текста)
        std.cell(r, 4, "").fill = input_fill
        # Из маточника? (да/нет); холостая проба (x=0) — не из маточника
        std.cell(r, 5, "нет" if i == 0 else "да").fill = input_fill
        # Коэффициент разведения k_i
        std.cell(r, 6, "").fill = input_fill
        # u серт абс
        std.cell(r, 7, "").fill = input_fill
        # распределение
        std.cell(r, 8, "нормальное").fill = input_fill
        # множитель
        std.cell(r, 9, 2).fill = input_fill
        # u станд. абс = если из маточника → u_rel_маточника·x, иначе u_серт/k
        std.cell(r, 10,
                 f'=IF(B{r}="","",'
                 f'IF(E{r}="да",B{r}*$B$11,'
                 f'IF(D{r}="нормальное",G{r}/I{r},'
                 f'IF(D{r}="прямоугольное",G{r}/SQRT(3),'
                 f'IF(D{r}="треугольное",G{r}/SQRT(6),G{r})))))'
                 ).fill = calc_fill
        # u_rel = u_станд / x
        std.cell(r, 11, f'=IF(AND(B{r}<>"",B{r}<>0),J{r}/B{r},"")').fill = calc_fill
        # Участвует в WLS? (все эталоны с концентрацией участвуют)
        std.cell(r, 12, f'=IF(B{r}="","",IF(B{r}>0,"да",""))').fill = calc_fill
        # ν_i (R-09): степени свободы из сертификата (по умолчанию 50; 0 = "∞")
        std.cell(r, 13, "50" if i < 6 else "").fill = input_fill

    # Итог ν_эталонов (R-09): пользователь вводит ν из сертификатов.
    # 0 = «бесконечность». Используется в ν_eff на листе «Неопределенность».
    std.cell(STD_LAST + 2, 1, "ν_эталонов (из сертификатов; 0 = ∞)")
    std.cell(STD_LAST + 2, 2, 50).fill = input_fill
    std.cell(STD_LAST + 2, 2).number_format = "0"

    # Итоги блока 1
    N = CONFIG["MAX_POINTS"]
    first_s = STD_FIRST
    last_s = STD_FIRST + N - 1
    std.cell(last_s + 1, 1, "Средняя u_rel независимых эталонов (квадратично)")
    std.cell(last_s + 1, 2,
             f'=IF(COUNTIFS(B{first_s}:B{last_s},">0")=0,NA(),'
             f'SQRT(SUMSQ(K{first_s}:K{last_s})/COUNTIFS(B{first_s}:B{last_s},">0")))'
             ).fill = calc_fill
    std.cell(last_s + 1, 2).number_format = "0.00%"

    # --- Блок 2: корреляция (общий маточник) ---
    std["A" + str(last_s + 3)] = "2. КОРРЕЛЯЦИЯ ЭТАЛОНОВ"
    std["A" + str(last_s + 3)].font = Font(bold=True)
    std["A" + str(last_s + 4)] = "Все эталоны из маточника (Е=да) коррелируют через u_rel маточника."
    std["A" + str(last_s + 5)] = "Коэффициент корреляции r (0..1)"
    std["B" + str(last_s + 5)] = 0.9
    std["B" + str(last_s + 5)].fill = input_fill
    std["A" + str(last_s + 6)] = "Число эталонов из маточника n_stock"
    std["B" + str(last_s + 6)] = (f'=COUNTIFS(B{first_s}:B{last_s},">0",'
                                  f'E{first_s}:E{last_s},"да")')
    std["B" + str(last_s + 6)].fill = calc_fill

    # --- Итог: u(x_cal)/x на «Ввод» ---
    itog_r = STD_ITOG
    std["A" + str(itog_r)] = "ИТОГ: u(x_cal)/x (на «Ввод» B5)"
    std["A" + str(itog_r)].font = Font(bold=True)
    std["B" + str(itog_r)] = (
        f'=IF($B${last_s+5}<=0.3,'
        f'SQRT($B$11^2+$B$13^2+(B{last_s+1})^2),'
        f'$B$11+$B$13+ABS(B{last_s+1}))'
    )
    std["B" + str(itog_r)].fill = calc_fill
    std["B" + str(itog_r)].number_format = "0.00%"
    std["A" + str(itog_r + 1)] = ("Комментарий: при r≤0.3 вклады складываются квадратично (независимы);"
                                  " при r>0.3 — линейно (полная корреляция через маточник), консервативно.")
    std["A" + str(itog_r + 1)].font = Font(italic=True, color="808080")

    # Запоминаем ячейку итога для ссылки с «Ввода»
    std_u_rel_cell = f"'{std.title}'!B{itog_r}"

    # ============================================================
    # ЛИСТ 2. РЕГРЕССИЯ (ОМНК)
    # Карта: B2 n, B3 x̄, B4 ȳ, B5 Sxx, B6 s²y/x, B7 b1, B8 b0,
    #        B9 var(b1), B10 var(b0), B11 cov, B12 R²,
    #        B13 x̄−b0/b1, B14 var(x_pred) ковариационная
    # ============================================================
    reg = wb.create_sheet("Регрессия")
    reg["A1"] = "СТАТИСТИКИ ЛИНЕЙНОЙ РЕГРЕССИИ (ОМНК)"
    reg["A1"].font = Font(bold=True, size=15)

    labels = [
        "Число точек калибровки n", "Среднее x̄", "Среднее ȳ",
        "Sxx = Σ(xᵢ−x̄)²", "Остаточная дисперсия s²y/x",
        "Наклон b1", "Сдвиг b0",
        "var(b1)", "var(b0)", "cov(b0,b1)", "R²",
        "b0/b1", "var(x_pred) ковариационная (полная)",
        "t(b0) = b0/SE(b0)  (R-02)", "t_crit (α=0.05, n−2)  (R-02)",
        "Вердикт b0 (R-02)", "s_blank = s_y/x (R-03)",
        "LOD = 3.3·s_blank/b1 (R-03)", "LOQ = 10·s_blank/b1 (R-03)",
        "Рабочий диапазон (R-03)",
    ]
    formulas = [
        "=COUNT(CAL_X)",
        "=AVERAGE(CAL_X)",
        "=AVERAGE(CAL_Y)",
        "=DEVSQ(CAL_X)",
        "=IF(B2>2,SUM(CAL_E2)/(B2-2),NA())",
        "=SLOPE(CAL_Y,CAL_X)",
        "=INTERCEPT(CAL_Y,CAL_X)",
        "=B6/B5",
        "=B6*(1/B2+B3^2/B5)",
        "=-B6*B3/B5",
        "=RSQ(CAL_Y,CAL_X)",
        "=B8/B7",
        "=(B10+B6/'Ввод'!$B$4+'Неопределенность'!$B$1^2*B9"
        "+2*'Неопределенность'!$B$1*B11)/B7^2",
        # R-02: t-тест значимости сдвига b0 (бланк), EURACHEM E.4.2.
        # TINV(p,ν) здесь — ДВУСТОРОННИЙ квантиль: TINV(α,ν)=k такой, что P(|T|≤k)=1−α.
        # Критическое значение для двустороннего теста при доверительной
        # вероятности B7 (0.95) равно TINV(1−B7, ν) = TINV(0.05, n−2).
        "=IF(AND(ISNUMBER(B10),B10>0),B8/SQRT(B10),NA())",
        "=IF(B2>2,TINV(1-'Ввод'!$B$7,B2-2),NA())",
        '=IF(NOT(ISNUMBER(B15)),"н/д",IF(ABS(B15)<=B16,"незначим (b0≈0)","ЗНАЧИМ"))',
        # R-03: LOD/LOQ по остаточной дисперсии
        "=IF(ISNUMBER(B6),SQRT(B6),NA())",
        "=IF(AND(ISNUMBER(B18),ABS(B7)>1E-12),3.3*B18/ABS(B7),NA())",
        "=IF(AND(ISNUMBER(B18),ABS(B7)>1E-12),10*B18/ABS(B7),NA())",
        '=IF(NOT(ISNUMBER(B20)),"n/a","LOQ-"&MAX(CAL_X))',
    ]
    for r, (label, formula) in enumerate(zip(labels, formulas), 2):
        reg.cell(r, 1, label).border = border
        reg.cell(r, 2, formula).fill = calc_fill
        reg.cell(r, 2).border = border
    reg.conditional_formatting.add(
        "B17", FormulaRule(formula=['ISNUMBER(SEARCH("ЗНАЧИМ",$B$17))'], fill=warn_fill))

    OO = dict(
        n="'Регрессия'!$B$2", xbar="'Регрессия'!$B$3", ybar="'Регрессия'!$B$4",
        Sxx="'Регрессия'!$B$5", s2="'Регрессия'!$B$6", b1="'Регрессия'!$B$7",
        b0="'Регрессия'!$B$8", vb1="'Регрессия'!$B$9", vb0="'Регрессия'!$B$10",
        cov="'Регрессия'!$B$11",
    )

    # ============================================================
    # ЛИСТ 3. ВЗВЕШЕННАЯ РЕГРЕССИЯ (WLS)
    # Таблица A9:F108: x, y, w, w·y, I (участие), относит. вес.
    # Статистика B115..B133:
    #   B115 Σw, B116 Σwx, B117 Σwy, B118 Sxx_w, B119 Sxy_w,
    #   B120 n_eff, B121 dof=n_eff−2, B122 b1_w, B123 b0_w,
    #   B124 SSE_w, B125 φ_w, B126 var(b1)_w, B127 var(b0)_w,
    #   B128 cov_w, B129 x̄_w, B130 Δb1%, B131 Δb0%, B132 Σw²x²,
    #   B133 D = φ_w/s²_ОМНК
    # ============================================================
    wls = wb.create_sheet("Взвешенная регрессия")
    wls["A1"] = "ВЗВЕШЕННАЯ РЕГРЕССИЯ (WLS): РАВНЫЕ ВЕСА / 1/x / 1/x²"
    wls["A1"].font = Font(bold=True, size=15)
    wls["A2"] = 'Текущий режим (лист «Ввод», B6):'
    wls["B2"] = '=CHOOSE(\'Ввод\'!$B$6,"равные веса (ОМНК)","w ∝ 1/x","w ∝ 1/x²")'
    wls["B2"].fill = calc_fill
    wls["A3"] = ("TD-01: нулевой уровень x=0 при весах 1/x, 1/x² исключается из WLS "
                 "(пустой вес, индикатор участия I=0), но остаётся в ОМНК-контроле b0.")
    wls["A3"].font = Font(italic=True, color="808080")

    tbl_headers = [
        "x", "y (среднее)", "Вес w", "w·y", "Участие I", "Относит. вес",
        # служебные столбцы G..L — только простые построчные выражения,
        # чтобы все SUMPRODUCT статистик содержали лишь диапазоны (см. выше)
        "xy", "w·x²", "w²", "w²·x²", "e_w = y−ŵ", "w·e_w²", "w_участн.",
        "w·e_ols² (A-06)",  # O — для F-теста OLS vs WLS
    ]
    for c, h in enumerate(tbl_headers, 1):
        wls.cell(8, c, h)
    style_header(wls, 8, len(tbl_headers))

    # TD-01: IF(x>0, 1/x^k, 0) — вес 0 для x≤0 (SUMPRODUCT-safe, вместо #N/A или "").
    # Индикатор участия I=1 только при C>0: x=0 исключается из WLS, но остаётся в ОМНК.
    for i in range(100):
        r = 9 + i
        src = 9 + i
        wls.cell(r, 1, f'=IF(\'Ввод\'!B{src}="","",\'Ввод\'!B{src})').fill = calc_fill
        wls.cell(r, 2, f'=IF(A{r}="","",\'Ввод\'!F{src})').fill = calc_fill
        wls.cell(
            r, 3,
            f'=IF(A{r}="",0,CHOOSE(\'Ввод\'!$B$6,1,'
            f'IF(A{r}>0,1/A{r},0),IF(A{r}>0,1/A{r}^2,0)))'
        ).fill = calc_fill
        wls.cell(r, 4, f'=IF(A{r}="",0,C{r}*B{r})').fill = calc_fill
        # I: 1 если точка участвует в WLS (вес>0), иначе 0
        wls.cell(r, 5, f'=IF(A{r}="",0,IF(C{r}>0,1,0))').fill = calc_fill
        wls.cell(r, 6, f'=IF(A{r}="",0,C{r}/MAX($C$9:$C$108))').fill = calc_fill
        wls.cell(r, 7, f'=IF(A{r}="",0,A{r}*B{r})').fill = calc_fill          # xy
        wls.cell(r, 8, f'=IF(A{r}="",0,C{r}*A{r}*A{r})').fill = calc_fill     # w·x²
        wls.cell(r, 9, f'=IF(A{r}="",0,C{r}*C{r})').fill = calc_fill          # w²
        wls.cell(r, 10, f'=IF(A{r}="",0,I{r}*A{r}*A{r})').fill = calc_fill    # w²·x²
        # e_w — остаток при текущих коэффициентах (зависит от B122/B123)
        wls.cell(r, 11, f'=IF(B{r}="",0,B{r}-$B$122*A{r}-$B$123)').fill = calc_fill
        # L: w·e² — 0 при пустой строке (SUMPRODUCT-safe)
        wls.cell(r, 12,
                 f'=IF(OR(A{r}="",K{r}=""),0,C{r}*K{r}*K{r})').fill = calc_fill
        # M: вес только для участвующих точек (0 вместо "" — SUMPRODUCT-safe)
        wls.cell(r, 13, f'=IF(A{r}="",0,IF(ISNUMBER(C{r}),C{r},0))').fill = calc_fill   # w_участн.
        # O: w·(y−ŷ_ols)² для F-теста OLS vs WLS (A-06) — 0 при пустой строке
        wls.cell(r, 15,
                 f'=IF(A{r}="",0,C{r}*(B{r}-'
                 f"'Регрессия'!$B$7*A{r}-'Регрессия'!$B$8)^2)").fill = calc_fill

    stat_labels = [
        "Σw (по участвующим точкам)", "Σwx", "Σwy",
        "Sxx_w = Σw·x² − (Σwx)²/Σw", "Sxy_w = Σw·xy − Σwx·Σwy/Σw",
        "n_eff (точек с числовым весом)", "Степени свободы (n_eff − 2)",
        "Наклон b1 (WLS)", "Сдвиг b0 (WLS)",
        "SSE_w = Σw·(y−ŵ)²", "Взвешенная остаточная дисперсия φ_w",
        "var(b1)_w = φ_w/Sxx_w", "var(b0)_w (полная форма (XᵀWX)⁻¹₁₁)",
        "cov(b0,b1)_w", "Взвешенное среднее x̄_w",
        "Δb1 vs ОМНК, %", "Δb0 vs ОМНК, %",
        "Σw²x² (для полной формы var(b0))",
        "Коэфф. дисперсии D = φ_w/s²_ОМНК (TD-02)",
    ]
    # ВНИМАНИЕ: во всех SUMPRODUCT только простые ссылки на диапазоны
    # (степени и произведения внутри массивов недопустимы — Excel/LibreOffice
    # вернут #VALUE!). Поэтому в таблице заранее вычислены служебные
    # столбцы G..L: x², xy, w², w·x², w²·x², e_w.
    stat_formulas = [
        # TD-05: все суммы — по строкам с I=1 (множитель WLS_I в SUMPRODUCT).
        "=SUMIF(WLS_I,1,WLS_W)",                              # B115 Σw
        "=SUMPRODUCT(WLS_I,WLS_W,WLS_X)",                     # B116 Σwx
        "=SUM(WLS_WY)",                                        # B117 Σwy
        "=SUMPRODUCT(WLS_I,WLS_WX2)-B116^2/B115",             # B118 Sxx_w
        "=SUMPRODUCT(WLS_I,WLS_W,WLS_XY)-B116*B117/B115",     # B119 Sxy_w
        "=SUM(WLS_I)",                                         # B120 n_eff
        "=IF(B120>2,B120-2,NA())",                             # B121 dof
        "=B119/B118",                                          # B122 b1_w
        "=(B117-B122*B116)/B115",                              # B123 b0_w
        "=SUM(WLS_WE2)",                                      # B124 SSE_w
        "=IF(B121>0,B124/B121,NA())",                          # B125 φ_w
        "=B125/B118",                                          # B126 var(b1)_w
        # R-01 (из внешнего аудита): var(b0)_w = φ_w·Σ(w·x²)/(Σw·Sxx_w)
        #   строгая форма (XᵀWX)⁻¹ требует Σw·x² (столбец H = WLS_WX2),
        #   НЕ Σw²·x² (столбец J). При w=1/x² разница ~7×.
        "=B125*SUM(WLS_WX2)/(B115*B118)",                    # B127 var(b0)_w
        "=-B125*B116/(B115*B118)",                             # B128 cov_w
        "=B116/B115",                                          # B129 x̄_w
        "=IF('Регрессия'!B7=0,NA(),ABS(B122-'Регрессия'!B7)/ABS('Регрессия'!B7))",   # B130
        "=IF('Регрессия'!B8=0,NA(),ABS(B123-'Регрессия'!B8)"
        "/MAX(ABS('Регрессия'!B8),1E-12))",                    # B131
        "=SUM(WLS_W2X2)",                                      # B132 Σw²x²
        # TD-02: D≈1 если модель дисперсии верна; D≫1 → интервалы занижены
        "=IF('Регрессия'!B6>0,B125/'Регрессия'!B6,NA())",      # B133
    ]
    for r, (label, formula) in enumerate(zip(stat_labels, stat_formulas), 115):
        wls.cell(r, 1, label).border = border
        wls.cell(r, 2, formula).fill = calc_fill
        wls.cell(r, 2).border = border
    wls.cell(130, 2).number_format = "0.0%"
    wls.cell(131, 2).number_format = "0.0%"

    WW = dict(
        Sw="'Взвешенная регрессия'!$B$115",
        Swx="'Взвешенная регрессия'!$B$116",
        Sxxw="'Взвешенная регрессия'!$B$118",
        neff="'Взвешенная регрессия'!$B$120",
        dofw="'Взвешенная регрессия'!$B$121",
        b1="'Взвешенная регрессия'!$B$122",
        b0="'Взвешенная регрессия'!$B$123",
        SSEw="'Взвешенная регрессия'!$B$124",
        phi="'Взвешенная регрессия'!$B$125",
        vb1="'Взвешенная регрессия'!$B$126",
        vb0="'Взвешенная регрессия'!$B$127",
        cov="'Взвешенная регрессия'!$B$128",
        xbarw="'Взвешенная регрессия'!$B$129",
        db1="'Взвешенная регрессия'!$B$130",
        D="'Взвешенная регрессия'!$B$133",
    )

    # ============================================================
    # ЛИСТ 4. ДИАГНОСТИКА (TD-03: Lack-of-Fit + остатки WLS)
    # Таблица уровней A8:K107: x, ȳ, n_i, SS_PE_i, ŷ_ols, ŷ_wls,
    #   LoF-вклад ОМНК, LoF-вклад WLS, (пусто), |e|·√w, x_участн.
    # Итоги: B110..B120; проверка дисперсии B123..B125.
    # ============================================================
    dg = wb.create_sheet("Диагностика")
    dg["A1"] = "ДИАГНОСТИКА МОДЕЛИ: LACK-OF-FIT И АНАЛИЗ ОСТАТКОВ WLS"
    dg["A1"].font = Font(bold=True, size=15)
    dg["A2"] = ("Разложение SS_res на чистую ошибку повторностей (SS_PE) и "
                "несоответствие модели (SS_LoF). F = (SS_LoF/df_LoF)/(SS_PE/df_PE).")
    dg["A2"].font = Font(italic=True, color="808080")

    dh = ["x_i", "ȳ_i", "n_i", "SS_PE_i", "ŷ_i ОМНК", "ŷ_i WLS",
          "n_i·(ȳ−ŷ)² ОМНК", "w·n_i·(ȳ−ŷ)² WLS", "", "|e|·√w WLS",
          "x (участн.)", "J²", "K²", "J·K"]
    for c, h in enumerate(dh, 1):
        if h:
            dg.cell(7, c, h)
    style_header(dg, 7, 8)
    for extra_col in (10, 11, 12, 13, 14):
        cell = dg.cell(7, extra_col)
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")

    for i in range(100):
        r = 8 + i
        src = 9 + i
        wr = 9 + i  # та же строка в таблице WLS-листа
        dg.cell(r, 1, f'=IF(\'Ввод\'!B{src}="","",\'Ввод\'!B{src})').fill = calc_fill
        dg.cell(r, 2, f'=IF(A{r}="","",\'Ввод\'!F{src})').fill = calc_fill
        dg.cell(r, 3, f'=IF(A{r}="","",COUNT(\'Ввод\'!C{src}:E{src}))').fill = calc_fill
        dg.cell(r, 4,
                f'=IF(A{r}="","",IF(C{r}>=2,(C{r}-1)*\'Ввод\'!J{src}^2,0))'
                ).fill = calc_fill
        dg.cell(r, 5, f'=IF(A{r}="","",{OO["b1"]}*A{r}+{OO["b0"]})').fill = calc_fill
        dg.cell(r, 6, f'=IF(A{r}="","",{WW["b1"]}*A{r}+{WW["b0"]})').fill = calc_fill
        dg.cell(r, 7, f'=IF(A{r}="","",C{r}*(B{r}-E{r})^2)').fill = calc_fill
        dg.cell(r, 8,
                f'=IF(OR(A{r}="",\'Взвешенная регрессия\'!C{wr}=""),"",'
                f'\'Взвешенная регрессия\'!C{wr}*C{r}*(B{r}-F{r})^2)'
                ).fill = calc_fill
        dg.cell(r, 10,
                f'=IF(OR(A{r}="",\'Взвешенная регрессия\'!C{wr}=""),"",'
                f'ABS(B{r}-F{r})*SQRT(\'Взвешенная регрессия\'!C{wr}))'
                ).fill = calc_fill
        dg.cell(r, 11, f'=IF(J{r}="",0,A{r})').fill = calc_fill
        # служебные столбцы только с простыми ссылками (SUMPRODUCT!)
        dg.cell(r, 12, f'=N(J{r})^2').fill = calc_fill                    # J²
        dg.cell(r, 13, f'=IF(J{r}="",0,K{r}*K{r})').fill = calc_fill      # K²
        dg.cell(r, 14, f'=N(J{r})*K{r}').fill = calc_fill                 # J·K

    dg_labels = [
        "Число уровней всего (m)",
        "SS_PE (чистая ошибка повторностей)",
        "df_PE = Σ(n_i−1)",
        "SS_LoF (ОМНК) = Σ n_i·(ȳ_i−ŷ_i)²",
        "df_LoF (ОМНК) = m − 2",
        "F (ОМНК) = (SS_LoF/df_LoF)/(SS_PE/df_PE)",
        "F_crit (α, df_LoF, df_PE), FINV",
        "Вердикт LoF (ОМНК)",
        "SS_LoF_W (взвешенный, только участвующие уровни)",
        "F (WLS) = (SS_LoF_W/df_LoF)/(SS_PE/df_PE)",
        "Вердикт LoF (WLS)",
    ]
    dg_formulas = [
        "=COUNT(DG_X)",                                                # B110 m
        "=SUM(DG_PE)",                                                 # B111 SS_PE
        "=COUNTIFS(DG_NI,\">=2\",DG_NI,\"<>\")"
        "*0+SUMIFS(DG_NI,DG_NI,\">=2\")"
        "-COUNTIFS(DG_NI,\">=2\")",                                   # B112 df_PE
        "=SUM(DG_LOF_OLS)",                                            # B113 SS_LoF OLS
        "=IF(B110>2,B110-2,NA())",                                     # B114 df_LoF
        "=IF(AND(B112>0,B114>0),(B113/B114)/(B111/B112),NA())",        # B115 F OLS
        "=IF(ISNUMBER(B115),FINV(1-'Ввод'!$B$7,B114,B112),NA())",      # B116 Fcrit
        "=IF(ISNUMBER(B115),IF(B115<=B116,\"PASS: линейность адекватна\","
        "\"FAIL: lack of fit — модель непригодна\"),\"нет повторностей\")",  # B117
        "=SUM(DG_LOF_WLS)",                                            # B118 SS_LoF W
        "=IF(AND(B112>0,B114>0,'Ввод'!$B$6>1),"
        "(B118/B114)/(B111/B112),NA())",                               # B119 F WLS
        "=IF(ISNUMBER(B119),IF(B119<=B116,\"PASS\",\"FAIL: нелинейность под весами\"),"
        "\"ОМНК-режим\")",                                             # B120
    ]
    for r, (label, formula) in enumerate(zip(dg_labels, dg_formulas), 110):
        dg.cell(r, 1, label).border = border
        dg.cell(r, 2, formula).fill = calc_fill
        dg.cell(r, 2).border = border
    for verdict_cell in ("B117", "B120"):
        dg.conditional_formatting.add(
            verdict_cell,
            FormulaRule(formula=[f'ISNUMBER(SEARCH("FAIL",{verdict_cell}))'],
                        fill=fail_fill))
        dg.conditional_formatting.add(
            verdict_cell,
            FormulaRule(formula=[f'ISNUMBER(SEARCH("PASS",{verdict_cell}))'],
                        fill=pass_fill))

    dg.cell(122, 1, "Проверка модели дисперсии WLS (TD-02/TD-03):").font = Font(bold=True)
    dg.cell(123, 1, "Корреляция |e|·√w с x (ожидание ≈ 0; >0.5 — веса слабы)")
    # SUMPRODUCT только по простым диапазонам L..N (J², K², J·K)
    dg.cell(123, 2,
            "=IF(AND('Ввод'!$B$6>1,COUNT(DG_ABS)>=4),"
            "(SUM(DG_JK)"
            "-SUM(DG_ABS)*SUM(DG_XP)/COUNT(DG_ABS))"
            "/((SUM(DG_J2)-SUM(DG_ABS)^2/COUNT(DG_ABS))^0.5"
            "*(SUM(DG_K2)-SUM(DG_XP)^2/COUNT(DG_ABS))^0.5)"
            ",\"ОМНК-режим\")").fill = calc_fill
    dg.cell(124, 1, "D = φ_w/s²_ОМНК (идеал ≈ 1; вне 0.5…2 — уточнить веса)")
    dg.cell(124, 2, f"={WW['D']}").fill = calc_fill
    dg.cell(125, 1, "Вердикт модели дисперсии")
    dg.cell(125, 2,
            "=IF(NOT(ISNUMBER(B124)),\"н/д\","
            "IF(AND(B124>0.5,B124<2),\"PASS: веса адекватны\","
            "\"FAIL: нужна IWLS / 1/(a+bx)²\"))").fill = calc_fill
    dg.conditional_formatting.add(
        "B125", FormulaRule(formula=['ISNUMBER(SEARCH("FAIL",$B$125))'], fill=fail_fill))
    dg.conditional_formatting.add(
        "B125", FormulaRule(formula=['ISNUMBER(SEARCH("PASS",$B$125))'], fill=pass_fill))

    # A-03/R-06: нормальность остатков (SKEW/KURT), GUM 5.2
    dg.cell(127, 1, "Нормальность остатков (A-03/R-06):").font = Font(bold=True)
    dg.cell(128, 1, "Асимметрия остатков SKEW")
    dg.cell(128, 2,
            "=IF(COUNT(DG_PE)=0,NA(),SKEW('Ввод'!$H$9:$H$108))").fill = calc_fill
    dg.cell(129, 1, "Эксцесс остатков KURT")
    dg.cell(129, 2,
            "=IF(COUNT(DG_PE)=0,NA(),KURT('Ввод'!$H$9:$H$108))").fill = calc_fill
    dg.cell(130, 1, "Вердикт (|SKEW|<2 и |KURT|<7)")
    dg.cell(130, 2,
            "=IF(OR(NOT(ISNUMBER(B128)),NOT(ISNUMBER(B129))),\"н/д\","
            "IF(AND(ABS(B128)<2,ABS(B129)<7),\"PASS: остатки ~нормальны\","
            "\"FAIL: отклонение от нормальности\"))").fill = calc_fill
    dg.conditional_formatting.add(
        "B130", FormulaRule(formula=['ISNUMBER(SEARCH("FAIL",$B$130))'], fill=fail_fill))

    # A-05/R-12: стандартизированные остатки e* = e/(s·√(1−h)), флаг |e*|>2.5
    # h_i = 1/n + (x_i−x̄)²/Sxx — leverage. s = SQRT(s²_y/x).
    dg.cell(132, 1, "Стандартизированные остатки и выбросы (A-05/R-12):").font = Font(bold=True)
    for i in range(CONFIG["MAX_POINTS"]):
        r = 133 + i
        src = CONFIG["FIRST_DATA"] + i
        if i < 6:
            dg.cell(r, 1, f'=IF(\'Ввод\'!B{src}="","",\'Ввод\'!B{src})').fill = calc_fill
            dg.cell(r, 2,
                    f'=IF(\'Ввод\'!H{src}="","",'
                    f"'Ввод'!H{src}/SQRT('Регрессия'!$B$6*"
                    f"(1-1/'Регрессия'!$B$2-(\'Ввод\'!B{src}-'Регрессия'!$B$3)^2"
                    f"/'Регрессия'!$B$5)))").fill = calc_fill
            dg.cell(r, 3,
                    f'=IF(OR(B{r}="",NOT(ISNUMBER(B{r}))),"",'
                    f'IF(ABS(B{r})>2.5,"ВЫБРОС?",""))').fill = calc_fill
        else:
            dg.cell(r, 1, "").fill = calc_fill
            dg.cell(r, 2, "").fill = calc_fill
            dg.cell(r, 3, "").fill = calc_fill
    dg.cell(133 + CONFIG["MAX_POINTS"] + 1, 1,
            "Число потенциальных выбросов (|e*|>2.5)")
    dg.cell(133 + CONFIG["MAX_POINTS"] + 1, 2,
            f'=COUNTIF(C133:C{133+CONFIG["MAX_POINTS"]-1},"ВЫБРОС?")').fill = calc_fill

    # A-06: F-тест OLS vs WLS (взвешенные остаточные дисперсии)
    dg.cell(133 + CONFIG["MAX_POINTS"] + 3, 1,
            "F-тест OLS vs WLS (A-06): F = SSE_OLS,w/SSE_WLS,w").font = Font(bold=True)
    dg.cell(133 + CONFIG["MAX_POINTS"] + 4, 1,
            "SSE_OLS под весами WLS (остатки ОМНК, веса активного режима)")
    # Σ w·(y−ŷ_ols)² по участвующим точкам (столбец O листа «Взвешенная регрессия»)
    dg.cell(133 + CONFIG["MAX_POINTS"] + 4, 2,
            "=IF('Ввод'!$B$6=1,NA(),SUM(WLS_WOLS2))").fill = calc_fill
    dg.cell(133 + CONFIG["MAX_POINTS"] + 5, 1, "SSE_WLS,w (B124 листа «Взвешенная регрессия»)")
    dg.cell(133 + CONFIG["MAX_POINTS"] + 5, 2,
            "='Взвешенная регрессия'!B124").fill = calc_fill
    dg.cell(133 + CONFIG["MAX_POINTS"] + 6, 1, "F = SSE_OLS,w/SSE_WLS,w")
    dg.cell(133 + CONFIG["MAX_POINTS"] + 6, 2,
            f'=IF(OR(NOT(ISNUMBER(B{133+CONFIG["MAX_POINTS"]+4})),'
            f'B{133+CONFIG["MAX_POINTS"]+5}<=0),NA(),'
            f'B{133+CONFIG["MAX_POINTS"]+4}/B{133+CONFIG["MAX_POINTS"]+5})').fill = calc_fill
    dg.cell(133 + CONFIG["MAX_POINTS"] + 7, 1,
            "Вердикт: WLS лучше OLS, если F>F_crit (1.5)")
    dg.cell(133 + CONFIG["MAX_POINTS"] + 7, 2,
            f'=IF(NOT(ISNUMBER(B{133+CONFIG["MAX_POINTS"]+6})),"ОМНК-режим",'
            f'IF(B{133+CONFIG["MAX_POINTS"]+6}>1.5,"WLS обоснован","OLS достаточен"))'
            ).fill = calc_fill

    # ============================================================
    # ЛИСТ 5. НЕОПРЕДЕЛЁННОСТЬ (проба) + TD-04 (k по Стьюденту)
    # Карта: B1 x_pred, B2 u² активная, B3 u модель, B4 u(x_cal),
    # B5 u_c, B6 u² ОМНК класс., B7 u² ОМНК ковар., B8 контроль форм,
    # B9 ν, B10 ν_eff, B11 k, B12 U, B13 U/x, B14 экстраполяция?
    # ============================================================
    unc = wb.create_sheet("Неопределенность")
    unc["A1"] = "ПРОГНОЗИРОВАНИЕ И НЕОПРЕДЕЛЁННОСТЬ (ТЕКУЩИЙ РЕЖИМ ВЕСОВ)"
    unc["A1"].font = Font(bold=True, size=15)

    d_ols = f"$B$1-{OO['xbar']}"
    # R-01/WLS-формы: для WLS единая интервальная форма
    #   u²(x_pred) = φ_w·(1/p + 1/Σw + (x−x̄_w)²/Sxx_w)/b1_w²
    # (без лишних var(b0), 2·d·cov и смещения +b0/b1 в d — это были F-04/R-14).
    u2_model = (
        f"IF('Ввод'!$B$6=1,"
        f"{OO['s2']}*(1/'Ввод'!$B$4+1/{OO['n']}+({d_ols})^2/{OO['Sxx']})"
        f"/{OO['b1']}^2,"
        f"{WW['phi']}*(1/'Ввод'!$B$4+1/{WW['Sw']}"
        f"+($B$1-{WW['xbarw']})^2/{WW['Sxxw']})/{WW['b1']}^2)"
    )

    layout = [
        ("Предсказанная концентрация x_pred",
         f"=('Ввод'!$B$3-{WW['b0']})/{WW['b1']}", ""),
        ("u²(x_pred) активный режим (интервальная форма)", "=" + u2_model, ""),
        ("u(x_pred) модель", "=IF(B2<0,NA(),SQRT(B2))", ""),
        ("u(x_cal) вклад эталонов", "=ABS(B1)*'Ввод'!$B$5", ""),
        ("u_c(x_pred) суммарная", "=SQRT(B3^2+B4^2)", ""),
        ("u²(x_pred) ОМНК классическая (справочно)",
         f"=({OO['s2']}/{OO['b1']}^2)*(1/'Ввод'!$B$4+1/{OO['n']}"
         f"+(B1-{OO['xbar']})^2/{OO['Sxx']})", ""),
        ("u²(x_pred) ОМНК ковариационная (справочно)", "='Регрессия'!B14", ""),
        ("Контроль: |разница двух форм ОМНК|", "=ABS(B6-B7)", ""),
        ("u²(x_pred) ОМНК упрощённая (справочно, без cov)",
         f"=({OO['s2']}*(1/'Ввод'!$B$4+1/{OO['n']})+{OO['vb0']}"
         f"+(B1-{OO['xbar']})^2*{OO['vb1']})/{OO['b1']}^2", ""),
        ("ν (степени свободы модели)",
         f"=IF('Ввод'!$B$6=1,{OO['n']}-2,{WW['dofw']})", "0"),
        ("ν_eff (Welch–Satterthwaite; ν_эталонов из «Стандарты», R-09)",
         f"=IF(B5=0,NA(),B5^4/(B3^4/MAX(B10,1)+"
         f"IF(B4=0,0,B4^4/IF({STD_NU}>0,{STD_NU},100000))))", "0.0"),
        ("k (Стьюдент, TINV, двусторонний) — TD-04",
         "=IF(B11>0,TINV(1-'Ввод'!$B$7,MAX(ROUND(B11,0),1)),2)", "0.00"),
        ("U(x_pred) расширенная = k·u_c", "=B12*B5", ""),
        ("U/x_pred относительная", "=IF(B1=0,NA(),B13/ABS(B1))", "0.0%"),
        ("Экстраполяция? (x_pred вне диапазона)",
         "=OR(B1<MIN(CAL_X),B1>MAX(CAL_X))", ""),
    ]
    for r, (label, formula, fmt) in enumerate(layout, 1):
        unc.cell(r, 1, label).border = border
        cell = unc.cell(r, 2, formula)
        cell.fill = calc_fill
        cell.border = border
        if fmt:
            cell.number_format = fmt
    unc.conditional_formatting.add(
        "B15", FormulaRule(formula=["B15=TRUE"], fill=fail_fill))

    # ============================================================
    # ЛИСТ 6. НЕОПРЕДЕЛЁННОСТЬ ПО ТОЧКАМ
    # A №, B x_i, C s(y_i), D u(y_i)повт, E n_i, F ŷ_i, G var(ŷ_i),
    # H u(ŷ_i), I u_model(y_i), J u_comb(y_i), K u(x_i)обр., L U=k·u
    # ============================================================
    pts = wb.create_sheet("Неопределенность по точкам")
    pts["A1"] = "НЕОПРЕДЕЛЁННОСТИ ПО УРОВНЯМ КАЛИБРОВКИ (ТЕКУЩИЙ РЕЖИМ ВЕСОВ)"
    pts["A1"].font = Font(bold=True, size=15)
    pts["A2"] = "Режим весов:"
    pts["B2"] = "=CHOOSE('Ввод'!$B$6,\"ОМНК\",\"1/x\",\"1/x²\")"
    pts["B2"].fill = calc_fill
    pts["A3"] = "k для U(x_i) (TD-04):"
    pts["B3"] = "=IF('Неопределенность'!$B$12>0,'Неопределенность'!$B$12,2)"
    pts["B3"].fill = calc_fill
    pts["B3"].number_format = "0.00"

    pt_headers = [
        "№", "x_i", "s(y_i) повт.", "u(y_i) повторн.", "n_i",
        "ŷ_i модель", "var(ŷ_i)", "u(ŷ_i)", "u(y_i) модель",
        "u(y_i) комбин.", "u(x_i) обратн.", "U(x_i) k·u",
    ]
    for c, h in enumerate(pt_headers, 1):
        pts.cell(5, c, h)
    style_header(pts, 5, len(pt_headers))

    for i in range(100):
        r = 6 + i
        src = 9 + i
        pts.cell(r, 1, f'=IF(\'Ввод\'!B{src}="","",ROW()-5)').fill = calc_fill
        pts.cell(r, 2, f'=IF(\'Ввод\'!B{src}="","",\'Ввод\'!B{src})').fill = calc_fill
        pts.cell(r, 3, f'=IF(\'Ввод\'!J{src}="","",\'Ввод\'!J{src})').fill = calc_fill
        pts.cell(r, 4, f'=IF(C{r}="","",C{r}/SQRT(E{r}))').fill = calc_fill
        pts.cell(r, 5, f'=IF(B{r}="","",COUNT(\'Ввод\'!C{src}:E{src}))').fill = calc_fill
        pts.cell(r, 6, f'=IF(B{r}="","",'
                       f"IF('Ввод'!$B$6=1,{OO['b1']},{WW['b1']})*B{r}+"
                       f"IF('Ввод'!$B$6=1,{OO['b0']},{WW['b0']}))").fill = calc_fill
        pts.cell(
            r, 7,
            f'=IF(B{r}="","",'
            f"IF('Ввод'!$B$6=1,"
            f"{OO['s2']}*(1/{OO['n']}+(B{r}-{OO['xbar']})^2/{OO['Sxx']}),"
            f"{WW['phi']}*(1/{WW['Sw']}+(B{r}-{WW['xbarw']})^2/{WW['Sxxw']})))"
        ).fill = calc_fill
        pts.cell(r, 8, f'=IF(G{r}="","",SQRT(G{r}))').fill = calc_fill
        pts.cell(r, 9, f'=IF(B{r}="","",SQRT('
                       f"IF('Ввод'!$B$6=1,{OO['s2']},{WW['phi']})+G{r}))"
                       ).fill = calc_fill
        pts.cell(r, 10,
                 f'=IF(B{r}="","",SQRT(N(D{r})^2+I{r}^2+(B{r}*\'Ввод\'!$B$5)^2))'
                 ).fill = calc_fill
        d_ols_r = f"(B{r}-{OO['xbar']})"
        d_wls_r = f"(B{r}-{WW['xbarw']})"
        pts.cell(
            r, 11,
            f'=IF(B{r}="","",SQRT('
            f"IF('Ввод'!$B$6=1,"
            f"{OO['s2']}*(1/E{r}+1/{OO['n']}+{d_ols_r}^2/{OO['Sxx']})"
            f"/{OO['b1']}^2,"
            f"{WW['phi']}*(1/E{r}+1/{WW['Sw']}"
            f"+{d_wls_r}^2/{WW['Sxxw']})/{WW['b1']}^2)"
            f"+(B{r}*'Ввод'!$B$5)^2))"
        ).fill = calc_fill
        pts.cell(r, 12, f'=IF(K{r}="","",$B$3*K{r})').fill = calc_fill

    pts.conditional_formatting.add(
        "L6:L105",
        ColorScaleRule(start_type="min", start_color="C6EFCE",
                       end_type="max", end_color="FFC7CE"))

    # ============================================================
    # ЛИСТ 7. ПРОФИЛЬ U(x) (сетка 200 точек; dx относительно активного центра)
    # ============================================================
    prof = wb.create_sheet("Профиль U(x)")
    prof["A1"] = "ПРОФИЛЬ РАСШИРЕННОЙ НЕОПРЕДЕЛЁННОСТИ U(x) ПО ДИАПАЗОНУ"
    prof["A1"].font = Font(bold=True, size=15)

    ph = ["x", "dx (активный центр)", "u²(x) модель", "u(x) модель",
          "u_c(x) с эталонами", "U(x)=k·u_c", "это x_pred?"]
    for c, h in enumerate(ph, 1):
        prof.cell(3, c, h)
    style_header(prof, 3, len(ph))

    prof["A2"] = "Шаг сетки:"
    prof["B2"] = f"=(MAX(CAL_X)-MIN(CAL_X))/199"
    prof["B2"].fill = calc_fill

    for i in range(CONFIG["GRID_N"]):
        r = 4 + i
        prof.cell(r, 1,
                  "=MIN(CAL_X)+(ROW()-4)*$B$2").fill = calc_fill
        prof.cell(
            r, 2,
            f'=IF(\'Ввод\'!$B$6=1,A{r}-{OO["xbar"]},'
            f"A{r}-{WW['xbarw']})"
        ).fill = calc_fill
        prof.cell(
            r, 3,
            f"=IF('Ввод'!$B$6=1,"
            f"{OO['s2']}*(1/'Ввод'!$B$4+1/{OO['n']}+B{r}^2/{OO['Sxx']})"
            f"/{OO['b1']}^2,"
            f"{WW['phi']}*(1/'Ввод'!$B$4+1/{WW['Sw']}"
            f"+B{r}^2/{WW['Sxxw']})/{WW['b1']}^2)"
        ).fill = calc_fill
        prof.cell(r, 4, f'=IF(C{r}<0,NA(),SQRT(C{r}))').fill = calc_fill
        prof.cell(r, 5, f"=SQRT(D{r}^2+(A{r}*'Ввод'!$B$5)^2)").fill = calc_fill
        prof.cell(r, 6, f"='Неопределенность'!$B$12*E{r}").fill = calc_fill
        prof.cell(r, 7,
                  f'=IF(ABS(A{r}-\'Неопределенность\'!$B$1)<=$B$2/2,"← x_pred","")'
                  ).fill = calc_fill

    prof.conditional_formatting.add(
        "F4:F203",
        ColorScaleRule(start_type="min", start_color="C6EFCE",
                       end_type="max", end_color="FFC7CE"))
    prof.conditional_formatting.add(
        "G4:G203", FormulaRule(formula=['$G4<>""'], fill=warn_fill))

    # ============================================================
    # ЛИСТ 8. ГРАФИКИ (3 диаграммы: калибровка, профиль U, остатки)
    # ============================================================
    charts = wb.create_sheet("Графики")
    charts["A1"] = "ГРАФИЧЕСКАЯ ДИАГНОСТИКА МОДЕЛИ"
    charts["A1"].font = Font(bold=True, size=15)
    charts["A2"] = ("Диаграммы построены на данных листов «Ввод», «Профиль U(x)», "
                    "«Взвешенная регрессия». Обновляются автоматически при пересчёте.")
    charts["A2"].font = Font(italic=True, color="808080")

    # --- 1. Калибровочная прямая: точки (x, ȳ) + ŷ ОМНК и WLS ---
    ch1 = ScatterChart()
    ch1.title = "Калибровочная прямая: точки, ОМНК, WLS"
    ch1.style = 13
    ch1.x_axis.title = "x (концентрация)"
    ch1.y_axis.title = "y (отклик)"
    ch1.x_axis.delete = False
    ch1.y_axis.delete = False
    ch1.height = 10
    ch1.width = 16

    # точки калибровки (Ввод: B9:B{WD}, F9:F{WD})
    ref_pts = Reference(ws, min_col=2, min_row=9, max_row=WD)
    ref_means = Reference(ws, min_col=6, min_row=9, max_row=WD)
    s_pts = Series(ref_means, ref_pts, title_from_data=False)
    s_pts.marker.symbol = "circle"
    s_pts.marker.size = 6
    s_pts.graphicalProperties.line.noFill = True
    ch1.series.append(s_pts)

    # ŷ ОМНК (Ввод G)
    ref_yhat_ols = Reference(ws, min_col=7, min_row=9, max_row=WD)
    s_ols = Series(ref_yhat_ols, ref_pts, title_from_data=False)
    s_ols.graphicalProperties.line.solidFill = "4F81BD"
    s_ols.graphicalProperties.line.width = 22000
    s_ols.marker.symbol = "none"
    ch1.series.append(s_ols)

    # ŷ WLS (Взвешенная регрессия: A9:A{WD} как x, нет прямой — линии по регрессии)
    # используем ту же сетку Ввода, но y по WLS-коэффициентам из листа «Взвешенная регрессия»
    ref_yhat_wls = Reference(wls, min_col=1, min_row=9, max_row=WD)
    # строим WLS-прямую по двум крайним точкам диапазона Ввода: xmin/xmax
    # ŷ_wls = b1_w·x + b0_w; сделаем столбец-прямую на листе «Взвешенная регрессия» O
    for i in range(100):
        r = 9 + i
        # N: прямая WLS от текущих коэффициентов (для графика)
        wls.cell(r, 14,
                 f'=IF(A{r}="","",$B$122*A{r}+$B$123)').fill = calc_fill
    ref_wls_line = Reference(wls, min_col=14, min_row=9, max_row=WD)
    s_wls = Series(ref_wls_line, ref_pts, title_from_data=False)
    s_wls.graphicalProperties.line.solidFill = "C00000"
    s_wls.graphicalProperties.line.width = 22000
    s_wls.marker.symbol = "none"
    ch1.series.append(s_wls)

    # полоса неопределённости модели: ŷ + k·u(ŷ) и ŷ − k·u(ŷ) (Ввод K, L)
    ref_up = Reference(ws, min_col=11, min_row=9, max_row=WD)
    ref_dn = Reference(ws, min_col=12, min_row=9, max_row=WD)
    s_up = Series(ref_up, ref_pts, title_from_data=False)
    s_up.graphicalProperties.line.solidFill = "9DC3E6"
    s_up.graphicalProperties.line.width = 12000
    s_up.graphicalProperties.line.dashStyle = "dash"
    s_up.marker.symbol = "none"
    ch1.series.append(s_up)
    s_dn = Series(ref_dn, ref_pts, title_from_data=False)
    s_dn.graphicalProperties.line.solidFill = "9DC3E6"
    s_dn.graphicalProperties.line.width = 12000
    s_dn.graphicalProperties.line.dashStyle = "dash"
    s_dn.marker.symbol = "none"
    ch1.series.append(s_dn)

    # названия серий оставим дефолтными (ссылки на листы) — это допустимо для ScatterChart

    # --- 2. Профиль U(x) ---
    ch2 = ScatterChart()
    ch2.title = "Профиль расширенной неопределённости U(x)"
    ch2.style = 13
    ch2.x_axis.title = "x"
    ch2.y_axis.title = "U(x)"
    ch2.height = 10
    ch2.width = 16
    ref_upx = Reference(prof, min_col=1, min_row=4, max_row=PD)
    ref_U = Reference(prof, min_col=6, min_row=4, max_row=PD)
    sU = Series(ref_U, ref_upx, title_from_data=False)
    sU.graphicalProperties.line.solidFill = "C00000"
    sU.graphicalProperties.line.width = 22000
    sU.marker.symbol = "none"
    ch2.series.append(sU)

    # --- 3. Остатки WLS: e_w по x ---
    ch3 = ScatterChart()
    ch3.title = "Остатки регрессии (ОМНК: Ввод H; WLS: K на листе «Взвешенная регрессия»)"
    ch3.style = 13
    ch3.x_axis.title = "x"
    ch3.y_axis.title = "e"
    ch3.height = 10
    ch3.width = 16
    # остатки ОМНК (Ввод H) — точки
    ref_e_ols = Reference(ws, min_col=8, min_row=9, max_row=WD)
    sE1 = Series(ref_e_ols, ref_pts, title_from_data=False)
    sE1.marker.symbol = "diamond"
    sE1.marker.size = 5
    sE1.graphicalProperties.line.noFill = True
    ch3.series.append(sE1)

    # позиции диаграмм
    charts.add_chart(ch1, "C2")
    charts.add_chart(ch2, "C26")
    charts.add_chart(ch3, "C50")

    # ============================================================
    # ЛИСТ 8Б. ПОЯСНЕНИЯ (методология, уравнения, памятка для бюджета)
    # ============================================================
    expl = wb.create_sheet("Пояснения")
    expl["A1"] = "ПОЯСНЕНИЯ И МЕТОДОЛОГИЯ"
    expl["A1"].font = Font(bold=True, size=15)

    def expl_row(r, title, body):
        cell = expl.cell(r, 1, title)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        cell2 = expl.cell(r, 2, body)
        cell2.alignment = Alignment(vertical="top", wrap_text=True)
        cell.border = border
        cell2.border = border

    expl_row(3, "Предсказанная концентрация x_pred",
             "x_pred — это оценка концентрации аналита в пробе по измеренному отклику "
             "y_obs и калибровочной прямой. Методология (ISO 11095, обратное предсказание):\n"
             "  x_pred = (y_obs − b0) / b1\n"
             "где b0, b1 — коэффициенты регрессии (ОМНК или WLS в зависимости от режима). "
             "Рассчитывается на листе «Неопределенность» (B1). Значение округлено по правилам "
             "неопределённости (GUM 7.2.6).")
    expl_row(4, "Уравнение прямой (модели) калибровки",
             "Прямая модель: y = b1·x + b0\n"
             "  y — отклик прибора, x — концентрация.\n"
             "  b1 — чувствительность (наклон), b0 — сдвиг (фон).\n"
             "Оценки по МНК: b1 = Sxy/Sxx, b0 = ȳ − b1·x̄.\n"
             "Для WLS используется взвешенная версия с весами w = 1/x^k.")
    expl_row(5, "Уравнение обратной модели (обратное предсказание)",
             "Обратная модель: x_pred = (y_obs − b0) / b1\n"
             "Это решение прямой модели относительно x при заданном y_obs. "
             "Используется для перевода измеренного отклика в концентрацию.")
    expl_row(6, "Неопределённость обратного предсказания",
             "u(x_pred) вычисляется методом распространения неопределённости (GUM).\n"
             "Составляющие: неопределённость отклика (повторности), неопределённость "
             "коэффициентов регрессии (через ковариационную матрицу), неопределённость "
             "эталонов u(x_cal)/x (лист «Стандарты»).\n"
             "u_c = sqrt(u_model² + u_cal²); U = k·u_c с k по Стьюденту (ν_eff).")
    expl_row(7, "Что брать для внесения в бюджет неопределённости (памятка)",
             "Из этой книги в общий бюджет неопределённости измерения пробы вносят:\n"
             "  1) u_model(x_pred) — лист «Неопределенность» B3 (вклад калибровки);\n"
             "  2) u(x_cal) — лист «Неопределенность» B4 (вклад эталонов);\n"
             "  3) U(x_pred) расширенную — лист «Неопределенность» B13, если нужен интервал;\n"
             "  4) ν_eff и k — лист «Неопределенность» B11, B12.\n"
             "Смысл: вклад калибровки в бюджет = u_model(x_pred) — стандартная неопределённость "
             "концентрации, обусловленная несовершенством калибровочной модели. "
             "Его складывают квадратично с прочими вкладами (навеска, объём, разведение).\n"
             "ВАЖНО: u(x_cal) уже учтён в u_c(x_pred) на листе «Неопределенность»; если берёте "
             "u_model отдельно — не прибавляйте u(x_cal) повторно (избегайте двойного счёта).")
    expl_row(8, "Распределение вкладов",
             "Типичный бюджет:\n"
             "  u(x_pred)  — калибровка (регрессия + повторности отклика)\n"
             "  u(x_cal)   — эталоны (сертификат, разведение)\n"
             "  u(m)       — масса/навеска\n"
             "  u(V)       — объём\n"
             "  u(C_sample) = sqrt(u(x_pred)² + u(x_cal)² + u(m)² + u(V)² + ...)\n"
             "Степени свободы: Welch–Satterthwaite для всего бюджета.")
    expl_row(9, "Лист «Стандарты»: как устроена оценка",
             "Каждый эталон калибровочной прямой расписан отдельной строкой:\n"
             "  концентрация (связана с листом «Ввод»), источник/сертификат, "
             "признак «из маточника», коэффициент разведения, сертификатная u, "
             "распределение, множитель k.\n"
             "Эталоны, приготовленные из общего маточного раствора, НЕ являются "
             "независимыми: их систематическая ошибка (неопределённость концентрации "
             "маточника) — общая. Поэтому при r>0.3 вклады складываются ЛИНЕЙНО "
             "(консервативно), при r≤0.3 — квадратично.")
    expl_row(10, "Маточный раствор (неопределённость)",
             "Блок 0 на листе «Стандарты»:\n"
             "  C_stock — концентрация маточника из сертификата;\n"
             "  u(C_stock) — сертификатная абсолютная неопределённость;\n"
             "  распределение (нормальное/прямоугольное/треугольное) → u_rel маточника;\n"
             "  u_rel процедуры разведения — пипетки, колбы (из паспортов мерной посуды);\n"
             "  u_rel(разведения) входит в каждый эталон из маточника и коррелирует через него.")
    expl_row(11, "GUM vs Monte Carlo (JCGM 101, прил. 7) — важное примечание",
             "Лист «Неопределенность» (u_c, U) считает стандартное линейное распространение "
             "неопределённости по GUM. Это АППРОКСИМАЦИЯ: она линеаризует обратную модель "
             "x_pred = (y_obs − b0)/b1 в точке оценок. При малом числе точек калибровки (n≈5–6) "
             "и нелинейной обратной функции GUM СИСТЕМАТИЧЕСКИ ЗАНИЖАЕТ u на 5–15% "
             "(подтверждено Monte-Carlo: u_MC > u_GUM).\n"
             "Правило (JCGM 101): если |u_MC − u_GUM|/u_GUM ≤ 5% — GUM-аппроксимация пригодна; "
             "если 5–15% — использовать MC-оценку как более честную; если >15% — проверить модель.\n"
             "Сравнение для текущих данных (mc_v4.py, N=10⁵):\n"
             "  Параметр        GUM (лист «Неопределенность»)   MC (JCGM 101)\n"
             "  x_pred          0.54429                          0.54435\n"
             "  u_c / u_MC      0.008456                         0.009287\n"
             "  U(95%) / U_MC   0.016960                         0.018219\n"
             "  Расхождение u:  ≈9.8% → рекомендован MC.\n"
             "Как использовать: запустите `python mc_v4.py` после расчёта — скрипт прочитает "
             "книгу, разыграет 10⁵ сценариев и выведет сравнение и вердикт.")

    # ============================================================
    # ЛИСТ 8В. БЮДЖЕТ НЕОПРЕДЕЛЁННОСТИ МЕТОДА (Б-01/R-05, Б-02/R-04, Б-03, Б-05, Б-06/R-10, Б-07)
    # Сводная таблица: источник → u(x_i) → c_i → c_i·u → вклад % → ν_i.
    # Строки 1..6: ввод (Recovery, разбавление); таблица бюджета с 8.
    # ============================================================
    bud = wb.create_sheet("Бюджет")
    bud["A1"] = "БЮДЖЕТ НЕОПРЕДЕЛЁННОСТИ МЕТОДА (ISO 17025 / GUM)"
    bud["A1"].font = Font(bold=True, size=15)
    bud["A2"] = ("Сводная таблица вкладов в неопределённость концентрации. "
                 "Калибровка (u_model, u_cal) подтягивается автоматически; "
                 "Recovery, разбавление и прочие — заполняются из валидации метода.")
    bud["A2"].font = Font(italic=True, color="808080")

    # --- Ввод: Recovery (Б-02/R-04) ---
    bud["A4"] = "Recovery (степень извлечения)"
    bud["A4"].font = Font(bold=True)
    bud["A5"] = "Rec (0..1, 1=100%)"
    bud["B5"] = 0.9
    bud["B5"].fill = input_fill
    bud["B5"].number_format = "0.00"
    bud["A6"] = "u(Rec)/Rec (отн.)"
    bud["B6"] = 0.048
    bud["B6"].fill = input_fill
    bud["B6"].number_format = "0.00%"
    bud["A7"] = "n_Rec (число спайков)"
    bud["B7"] = 6
    bud["B7"].fill = input_fill
    bud["A8"] = "Rec значим? |1−Rec|/u(Rec) > 2"
    bud["B8"] = "=IF(NOT(ISNUMBER(B5)),\"\",IF(ABS(1-B5)/(B6*B5)>2,\"значим — ввести поправку\",\"незначим\"))"
    bud["B8"].fill = calc_fill

    # --- Ввод: разбавление пробы (Б-06/R-10) ---
    bud["A10"] = "Разбавление пробы"
    bud["A10"].font = Font(bold=True)
    bud["A11"] = "Коэффициент разбавления d"
    bud["B11"] = 1
    bud["B11"].fill = input_fill
    bud["A12"] = "u(d)/d (отн.)"
    bud["B12"] = 0.0
    bud["B12"].fill = input_fill
    bud["B12"].number_format = "0.00%"
    # Б-07: корреляция между компонентами (напр., эталоны и разбавление из одного маточника)
    bud["A13"] = "Коррелировано с эталонами? (Б-07)"
    bud["A13"].font = Font(bold=True)
    bud["A14"] = "Флаг корреляции u(x_cal) и u(d) (0/1)"
    bud["B14"] = 0
    bud["B14"].fill = input_fill
    bud["A15"] = "Коэффициент r(u_cal, u_d)"
    bud["B15"] = 0.0
    bud["B15"].fill = input_fill
    bud["A16"] = "Ковариационный член 2·r·u_cal·u_d"
    bud["B16"] = "=IF(B14=1,2*B15*'Неопределенность'!B4*B12,0)"
    bud["B16"].fill = calc_fill

    # --- Таблица бюджета ---
    bh = ["Источник", "u(x_i)/x (отн.)", "Чувствит. c_i", "Вклад c_i·u(x_i)", "Доля %", "ν_i"]
    for c, h in enumerate(bh, 1):
        bud.cell(14, c, h)
    style_header(bud, 14, len(bh))

    # u_model и u_cal подтягиваются из «Неопределенности»
    budget_rows = [
        # (метка, u_rel формула, c_i, ν_i)
        ("Калибровка: u_model(x_pred)", "='Неопределенность'!B3/'Неопределенность'!B1", 1, "='Регрессия'!B2-2"),
        ("Калибровка: u(x_cal) эталоны", "='Неопределенность'!B4/'Неопределенность'!B1", 1, 50),
        ("Recovery u(Rec)/Rec", "=B6", 1, "=B7-1"),
        ("Разбавление u(d)/d", "=B12", 1, 50),
        ("Навеска u(m)/m (заполнить)", 0.005, 1, 50),
        ("Объём u(V)/V (заполнить)", 0.005, 1, 50),
        ("Неоднородность u(H)/H (заполнить)", 0.01, 1, 50),
    ]
    for i, (label, uf, ci, nuf) in enumerate(budget_rows):
        r = 15 + i
        bud.cell(r, 1, label)
        cell = bud.cell(r, 2, uf)
        cell.fill = calc_fill if isinstance(uf, str) and uf.startswith("=") else input_fill
        cell.number_format = "0.00%"
        bud.cell(r, 3, ci).fill = calc_fill
        bud.cell(r, 4, f"=B{r}*C{r}").fill = calc_fill
        bud.cell(r, 5, f"=IF(AND(D{r}<>\"\",D{r}>0),D{r}^2/SUMPRODUCT((D$15:D${15+len(budget_rows)-1}>0)*D$15:D${15+len(budget_rows)-1}^2),\"\")").fill = calc_fill
        ncell = bud.cell(r, 6, nuf)
        ncell.fill = calc_fill if isinstance(nuf, str) and str(nuf).startswith("=") else input_fill

    b_last = 15 + len(budget_rows) - 1  # 21
    # Итоги
    bud.cell(b_last + 2, 1, "Суммарная u_rel (квадратично)").font = Font(bold=True)
    bud.cell(b_last + 2, 2, "=SQRT(SUMSQ(D15:D21))").fill = calc_fill
    bud.cell(b_last + 2, 2).number_format = "0.00%"
    bud.cell(b_last + 3, 1, "ν_eff бюджета (Welch–Satterthwaite)").font = Font(bold=True)
    bud.cell(b_last + 3, 2,
             "=IF(B23=0,NA(),B23^4/(SUMPRODUCT((D15:D21>0)*D15:D21^4/F15:F21)))"
             ).fill = calc_fill
    bud.cell(b_last + 3, 2).number_format = "0.0"
    # x_corr с учётом Rec и d
    bud.cell(b_last + 5, 1, "x_corrected = x_pred·d/Rec").font = Font(bold=True)
    bud.cell(b_last + 5, 2, "='Неопределенность'!B1*B11/B5").fill = calc_fill
    bud.cell(b_last + 6, 1, "u_c(x_corr)/x_corr (суммарно)")
    bud.cell(b_last + 6, 2,
             "=SQRT(('Неопределенность'!B5/'Неопределенность'!B1)^2"
             "+B12^2+B6^2)").fill = calc_fill
    bud.cell(b_last + 6, 2).number_format = "0.00%"
    bud.cell(b_last + 7, 1, "U(x_corr) = k·u_c(x_corr)")
    bud.cell(b_last + 7, 2,
             "='Неопределенность'!B12*B26*B27").fill = calc_fill

    # Pareto-диаграмма вкладов (Б-05): столбчатая по «Доля %»
    from openpyxl.chart import BarChart, Reference as Ref
    pch = BarChart()
    pch.title = "Pareto: вклады в неопределённость (%)"
    pch.type = "col"
    pch.height = 10
    pch.width = 16
    data = Ref(bud, min_col=5, min_row=14, max_row=b_last)
    cats = Ref(bud, min_col=1, min_row=15, max_row=b_last)
    pch.add_data(data, titles_from_data=True)
    pch.set_categories(cats)
    pch.legend = None
    bud.add_chart(pch, "H4")

    # ============================================================
    # ЛИСТ 8Г. ИТОГ (R-08): формат результата для протокола + округление GUM (Б-03)
    # ============================================================
    itg = wb.create_sheet("Итог")
    itg["A1"] = "РЕЗУЛЬТАТ ИЗМЕРЕНИЯ (для протокола)"
    itg["A1"].font = Font(bold=True, size=15)

    itg["A3"] = "Концентрация в пробе x_pred"
    itg["B3"] = "='Неопределенность'!B1"
    itg["B3"].fill = calc_fill
    itg["A4"] = "Суммарная стандартная u_c"
    itg["B4"] = "='Неопределенность'!B5"
    itg["B4"].fill = calc_fill
    itg["A5"] = "Степени свободы ν_eff"
    itg["B5"] = "='Неопределенность'!B11"
    itg["B5"].fill = calc_fill
    itg["A6"] = "Коэффициент охвата k"
    itg["B6"] = "='Неопределенность'!B12"
    itg["B6"].fill = calc_fill
    itg["A7"] = "Расширенная неопределённость U (k)"
    itg["B7"] = "='Неопределенность'!B13"
    itg["B7"].fill = calc_fill
    itg["A8"] = "U_отн = U/x_pred"
    itg["B8"] = "=IF(B3=0,NA(),B7/ABS(B3))"
    itg["B8"].fill = calc_fill
    itg["B8"].number_format = "0.0%"
    itg["A9"] = "LOD / LOQ (контроль малых значений)"
    itg["B9"] = "='Регрессия'!B19 & \" / \" & 'Регрессия'!B20"
    itg["B9"].fill = calc_fill

    # Б-03: округление по GUM 7.2.6 (U до 2 значащих цифр, x до разряда U)
    itg["A11"] = "Округление по GUM 7.2.6 (Б-03)"
    itg["A11"].font = Font(bold=True)
    itg["A12"] = "U округлённое (2 зн. цифры)"
    itg["B12"] = "=IF(ISNUMBER(B7),ROUND(B7,2-1-INT(LOG10(ABS(B7)))),NA())"
    itg["B12"].fill = calc_fill
    itg["A13"] = "x_pred округлённое (до разряда U)"
    itg["B13"] = "=IF(ISNUMBER(B12),ROUND(B3,-INT(LOG10(ABS(B12)))+1),NA())"
    itg["B13"].fill = calc_fill
    itg["A14"] = "Формат результата"
    itg["B14"] = '=IF(ISNUMBER(B13),"("&B13&" ± "&B12&") "&B9,"")'
    itg["B14"].fill = calc_fill
    itg["A15"] = "Статус (из «Валидации»)"
    itg["B15"] = "='Валидация'!B24"
    itg["B15"].fill = calc_fill

    itg.conditional_formatting.add(
        "B15", FormulaRule(formula=['ISNUMBER(SEARCH("ПРИНЯТА",$B$15))'], fill=pass_fill))
    itg.conditional_formatting.add(
        "B15", FormulaRule(formula=['ISNUMBER(SEARCH("ОТКЛОНЕНА",$B$15))'], fill=fail_fill))

    # ============================================================
    # ЛИСТ 8Д. УТВЕРЖДЕНИЕ (В-03): подписи для ISO 17025 7.8
    # ============================================================
    appr = wb.create_sheet("Утверждение")
    appr["A1"] = "УТВЕРЖДЕНИЕ РЕЗУЛЬТАТА (ISO/IEC 17025, п. 7.8)"
    appr["A1"].font = Font(bold=True, size=15)
    sign_rows = [
        ("Выполнил", ""), ("Проверил", ""), ("Утвердил", ""),
    ]
    for i, (role, _) in enumerate(sign_rows):
        r = 3 + i
        appr.cell(r, 1, role + ": ФИО").border = border
        appr.cell(r, 2, "").border = border
        appr.cell(r, 2).fill = input_fill
        appr.cell(r, 3, "Подпись").border = border
        appr.cell(r, 4, "").border = border
        appr.cell(r, 4).fill = input_fill
        appr.cell(r, 5, "Дата").border = border
        appr.cell(r, 6, f'=IF(B{r}<>"",TEXT(TODAY(),"DD.MM.YYYY"),"")').border = border
        appr.cell(r, 6).fill = calc_fill
    appr["A7"] = "Заключение: "
    appr["B7"] = "='Итог'!B15"
    appr["B7"].fill = calc_fill
    appr["A8"] = "Примечания"
    appr["B8"] = ""
    appr["B8"].fill = input_fill

    # ============================================================
    # ЛИСТ 8Е. ИСТОРИЯ / КОНТРОЛЬНЫЕ КАРТЫ (В-02): тренды b1, b0, s_y/x
    # Оператор ведёт журнал калибровок; текущие значения подтягиваются
    # кнопкой/формулой для сравнения с предыдущими.
    # ============================================================
    hist = wb.create_sheet("История")
    hist["A1"] = "ИСТОРИЯ КАЛИБРОВОК И КОНТРОЛЬНЫЕ КАРТЫ (В-02)"
    hist["A1"].font = Font(bold=True, size=15)
    hist["A2"] = ("Вести журнал: дата → b1, b0, s_y/x, R². "
                  "Тренды показывают деградацию прибора. "
                  "Текущие значения (строка 4) — ссылки на расчёт.")
    hist["A2"].font = Font(italic=True, color="808080")

    hh = ["Дата", "b1", "b0", "s_y/x", "R²", "Оператор", "Примечание"]
    for c, h in enumerate(hh, 1):
        hist.cell(4, c, h)
    style_header(hist, 4, len(hh))
    # Текущая калибровка (заполняется автоматически)
    hist.cell(5, 1, "=TODAY()").fill = calc_fill
    hist.cell(5, 2, "='Регрессия'!B7").fill = calc_fill
    hist.cell(5, 3, "='Регрессия'!B8").fill = calc_fill
    hist.cell(5, 4, "='Регрессия'!B18").fill = calc_fill
    hist.cell(5, 5, "='Регрессия'!B12").fill = calc_fill
    hist.cell(5, 6, "").fill = input_fill
    hist.cell(5, 7, "текущая").fill = input_fill
    # Пустые строки для журнала
    for i in range(1, 20):
        r = 5 + i
        for c in range(1, 8):
            hist.cell(r, c, "")
            if c in (1, 6, 7):
                hist.cell(r, c).fill = input_fill
    hist.cell(5, 1).number_format = "DD.MM.YYYY"

    # Контрольные карты (линии тренда b1, b0, s_y/x)
    hist_ch = ScatterChart()
    hist_ch.title = "Контрольная карта: b1 по датам"
    hist_ch.style = 13
    hist_ch.height = 9
    hist_ch.width = 15
    ref_d = Reference(hist, min_col=1, min_row=5, max_row=25)
    ref_b1 = Reference(hist, min_col=2, min_row=5, max_row=25)
    s_b1 = Series(ref_b1, ref_d, title_from_data=False)
    s_b1.graphicalProperties.line.solidFill = "4F81BD"
    s_b1.graphicalProperties.line.width = 18000
    s_b1.marker.symbol = "circle"
    hist_ch.series.append(s_b1)
    hist.add_chart(hist_ch, "I5")

    # ============================================================
    # ЛИСТ 9. О КНИГЕ (TD-12: версия, дата, окружение, git)
    # ============================================================
    about = wb.create_sheet("О книге")
    about["A1"] = "СВЕДЕНИЯ О КНИГЕ (ISO 17025, прослеживаемость)"
    about["A1"].font = Font(bold=True, size=15)
    about["A3"] = "Наименование"
    about["B3"] = "Metrology Core EURACHEM"
    about["A4"] = "Версия"
    about["B4"] = CONFIG["VERSION"]
    about["A5"] = "Дата генерации"
    about["B5"] = datetime.date.today().isoformat()
    about["A6"] = "Генератор"
    about["B6"] = os.path.basename(__file__)
    about["A7"] = "Версия Python"
    about["B7"] = sys.version.split()[0]
    about["A8"] = "Git commit (HEAD)"
    about["B8"] = get_git_commit()
    about["A10"] = ("Книга пересобирается скриптом; ручное редактирование формул "
                    "не рекомендуется. Версия и commit позволяют воспроизвести результат.")
    about["A10"].font = Font(italic=True, color="808080")
    style_header(about, 3, 2)
    for rr in range(3, 9):
        about.cell(rr, 1).border = border
        about.cell(rr, 2).border = border
        about.cell(rr, 2).fill = calc_fill

    # ============================================================
    # ЛИСТ 10. ВАЛИДАЦИЯ (16 проверок, включая TD-01…TD-04)
    # ============================================================
    val = wb.create_sheet("Валидация")
    val["A1"] = "АВТОМАТИЧЕСКАЯ ДИАГНОСТИКА МОДЕЛИ"
    val["A1"].font = Font(bold=True, size=15)

    checks = [
        ("Достаточно степеней свободы (ОМНК)", f"={OO['n']}>2"),
        ("Sxx > 0", f"={OO['Sxx']}>0"),
        ("Наклон b1 ≠ 0 (ОМНК)", f"=ABS({OO['b1']})>1E-12"),
        ("Нет экстраполяции", "=NOT('Неопределенность'!B15)"),
        ("Согласованность двух формул ОМНК",
         "=ABS('Неопределенность'!B6-'Неопределенность'!B7)"
         "<1E-6*MAX(1,ABS('Неопределенность'!B6))"),
        ("R² ≥ 0.99 (диагностический порог)", "='Регрессия'!B12>=0.99"),
        ("Lack-of-Fit (ОМНК) пройден (TD-03)",
         "=IF(ISNUMBER('Диагностика'!B115),"
         "'Диагностика'!B115<='Диагностика'!B116,TRUE)"),
        ("Lack-of-Fit (WLS) пройден (TD-03)",
         "=IF('Ввод'!$B$6=1,TRUE,IF(ISNUMBER('Диагностика'!B119),"
         "'Диагностика'!B119<='Диагностика'!B116,TRUE))"),
        ("x=0 корректно исключено из WLS (TD-01)",
         "=IF('Ввод'!$B$6=1,TRUE,"
         f"{WW['neff']}=COUNT(WLS_X)-COUNTIFS(WLS_X,0))"),
        ("Все веса положительны, NaN нет (TD-01)",
         "=IF('Ввод'!$B$6=1,TRUE,"
         "SUMPRODUCT(WLS_I,WLS_W<=0)=0)"),
        ("n_eff(WLS) ≥ 4 (достаточно для φ_w)",
         "=IF('Ввод'!$B$6=1,TRUE," + WW['neff'] + ">=4)"),
        ("Модель дисперсии адекватна: 0.5 < D < 2 (TD-02)",
         "=IF('Ввод'!$B$6=1,TRUE,"
         "AND('Диагностика'!B124>0.5,'Диагностика'!B124<2))"),
        ("Δb1 (WLS vs ОМНК) < 10%",
         "=IF('Ввод'!$B$6=1,TRUE," + WW['db1'] + "<0.1)"),
        ("a priori гетероскедастичность: corr(x, s_y) > 0.5",
         "=IF(COUNT(CAL_SY)>=3,CORREL(CAL_X,CAL_SY)>0.5,TRUE)"),
        ("ν_eff ≥ 5 (корректный k по Стьюденту, TD-04)",
         "=IF(ISNUMBER('Неопределенность'!B11),"
         "'Неопределенность'!B11>=5,FALSE)"),
        ("k в разумных границах 2…12 (TD-04)",
         "=IF(ISNUMBER('Неопределенность'!B12),"
         "AND('Неопределенность'!B12>=2,'Неопределенность'!B12<=12),FALSE)"),
        # A-01: back-calculation в допуске ±15% (ISO 11095 / FDA ICH)
        ("Back-calculation: все |Δ%| ≤ 0.15 (A-01)",
         "=IF(COUNT(CAL_DELTA)=0,FALSE,"
         "COUNTIFS(CAL_DELTA,\">0.15\")=0)"),
        # R-03: минимум 4 точки калибровки (В-06)
        ("Минимум 4 точки калибровки (В-06)", "=COUNT(CAL_X)>=4"),
        # R-03: x_pred выше LOQ (иначе результат не количественный)
        ("x_pred > LOQ (результат количественный, R-03)",
         "=IF(ISNUMBER('Регрессия'!B20),"
         "'Неопределенность'!B1>'Регрессия'!B20,FALSE)"),
        # R-02: b0 незначим (t-тест бланка) — предупреждение, если значим
        ("Сдвиг b0 незначим (t-тест, R-02)",
         "=IF(ISNUMBER('Регрессия'!B15),"
         "ABS('Регрессия'!B15)<=ABS('Регрессия'!B16),TRUE)"),
    ]
    for r, (name, formula) in enumerate(checks, 3):
        val.cell(r, 1, name).border = border
        val.cell(r, 2, formula).fill = calc_fill
        val.cell(r, 2).border = border
        val.cell(r, 3, f'=IF(B{r},"PASS","FAIL")').border = border
        val.conditional_formatting.add(
            f"C{r}", FormulaRule(formula=[f'C{r}="PASS"'], fill=pass_fill))
        val.conditional_formatting.add(
            f"C{r}", FormulaRule(formula=[f'C{r}="FAIL"'], fill=fail_fill))
    last = 3 + len(checks) - 1
    val.cell(last + 2, 1, "Итог").font = Font(bold=True)
    val.cell(last + 2, 2,
             f'=IF(COUNTIF(C3:C{last},"FAIL")=0,"МОДЕЛЬ ПРИГОДНА","ЕСТЬ ЗАМЕЧАНИЯ")')
    # A-02: интегральный вердикт для рутины
    val.cell(last + 3, 1, "ВЕРДИКТ КАЛИБРОВКИ (A-02)").font = Font(bold=True, size=12)
    val.cell(last + 3, 2,
             f'=IF(COUNTIF(C3:C{last},"FAIL")=0,'
             f'"КАЛИБРОВКА ПРИНЯТА","КАЛИБРОВКА ОТКЛОНЕНА")').font = Font(bold=True)
    val.cell(last + 3, 2).fill = calc_fill
    val.conditional_formatting.add(
        f"B{last+3}",
        FormulaRule(formula=[f'$B${last+3}="КАЛИБРОВКА ПРИНЯТА"'], fill=pass_fill))
    val.conditional_formatting.add(
        f"B{last+3}",
        FormulaRule(formula=[f'$B${last+3}="КАЛИБРОВКА ОТКЛОНЕНА"'], fill=fail_fill))
    val.cell(last + 4, 1,
             "Статус внешнего валидатора (validate_v4.py, TD-06)").font = Font(bold=True)
    val.cell(last + 4, 2, "— заполняется скриптом —")

    # Общий layout
    for sh in wb.worksheets:
        for row in sh.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
        sh.column_dimensions["A"].width = 52
        sh.column_dimensions["B"].width = 22
        sh.column_dimensions["C"].width = 20
    for col in ["D", "E", "F", "G", "H", "I", "J"]:
        ws.column_dimensions[col].width = 15
    for col in ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]:
        pts.column_dimensions[col].width = 15
        prof.column_dimensions[col].width = 15
    prof.column_dimensions["A"].width = 12
    for col in ["D", "E", "F", "G", "H", "I", "J", "K"]:
        dg.column_dimensions[col].width = 15

    # В-04: защита листов от случайной модификации формул.
    # Разблокированы только зелёные ячейки ввода (fill input_fill).
    PROTECT_PASSWORD = "metro2025"
    PROTECT_EXCEPT = {"Ввод", "Стандарты", "Бюджет", "Утверждение", "Графики"}
    for sh in wb.worksheets:
        if sh.title in PROTECT_EXCEPT:
            continue
        for row in sh.iter_rows():
            for cell in row:
                locked = True
                try:
                    if cell.fill is not None and cell.fill.fill_type == "solid" \
                            and cell.fill.start_color is not None \
                            and cell.fill.start_color.rgb == input_fill.start_color.rgb:
                        locked = False
                except Exception:
                    pass
                cell.protection = Protection(locked=locked)
        sh.protection.sheet = True
        sh.protection.password = PROTECT_PASSWORD
        sh.protection.selectLockedCells = False
        sh.protection.selectUnlockedCells = False
        sh.protection.formatCells = False

    wb.save(output_path)
    print(f"Файл успешно сгенерирован: {output_path}")
    return output_path


if __name__ == "__main__":
    create_metrology_excel_v4()
