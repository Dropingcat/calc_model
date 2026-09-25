"""
cal_v3.py — генератор метрологической калибровочной модели (EURACHEM/CITAC)

Расширение относительно cal_v2.py:
  * Ввод: добавлен столбец s(y_i) — стандартное отклонение повторностей по уровням
    и переключатель весов WLS (B6): 1 = равные веса (ОМНК), 2 = w∝1/x, 3 = w∝1/x².
  * Новый лист «Взвешенная регрессия»: WLS с выбором весов, сравнение коэффициентов
    WLS/ОМНК, относительный вес каждой точки.
  * Новый лист «Неопределенность по точкам»: для каждого уровня — u(y_i) из
    повторностей, дисперсия предсказания модели var(ŷ_i), u(x_i) — неопределённость
    обратного предсказания (интервальная форма, без вычитания самого x_i).
  * Новый лист «Профиль U(x)»: кривая расширенной неопределённости U(x) по сетке
    концентраций (для текущего режима весов).
  * Лист «Регрессия» дополнен ковариационной формой var(x_pred) (строки B13–B14).
  * Лист «Валидация» дополнен проверками согласованности ОМНК/WLS и a priori
    гетероскедастичности (s(y) растёт с x → 1/x² уместен).

Дисперсия обратного предсказания (интервальная форма, Nychka–Ziegler / Miller):
  u²(x0) = [ s²·(1/p + 1/n + (ȳ−b0)²/Sxx_w) + var(b0) + (x0−x̄_w)²·var(b1)
             + 2·(x0−x̄_w)·cov(b0,b1) ] / b1²   — для ОМНК; для WLS аналогично
  с взвешенными x̄_w, Sxx_w и матрицей ковариаций (XᵀWX)⁻¹.
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation


def create_metrology_excel_v3(output_path="Metrology_Core_EURACHEM_v3.xlsx"):
    wb = Workbook()

    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="4F81BD")
    input_fill = PatternFill("solid", fgColor="E2EFDA")
    calc_fill = PatternFill("solid", fgColor="DDEBF7")
    pass_fill = PatternFill("solid", fgColor="C6EFCE")
    fail_fill = PatternFill("solid", fgColor="FFC7CE")
    warn_fill = PatternFill("solid", fgColor="FFEB9C")

    def style_header(ws, row, cols):
        for c in range(1, cols + 1):
            cell = ws.cell(row, c)
            cell.fill = header_fill
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

    # ============================================================
    # ЛИСТ 1. ВВОД
    # ============================================================
    ws = wb.active
    ws.title = "Ввод"

    ws["A1"] = "ПАРАМЕТРЫ ПРОБЫ И КАЛИБРОВОЧНЫЕ ДАННЫЕ"
    ws["A1"].font = Font(bold=True, size=15)

    params = [
        ("Средний отклик пробы (y_obs)", 0.150),
        ("Число повторностей пробы (p)", 2),
        ("Относительная стандартная неопределённость эталонов u(x_cal)/x", 0.02),
        ("Режим весов WLS (1=равные, 2=1/x, 3=1/x²)", 3),
    ]
    for r, (name, value) in enumerate(params, 3):
        ws.cell(r, 1, name)
        ws.cell(r, 2, value).fill = input_fill
    ws["B5"].number_format = "0.00%"

    dv_w = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False)
    ws.add_data_validation(dv_w)
    dv_w.add(ws["B6"])

    headers = [
        "№", "Концентрация x", "Отклик 1 y", "Отклик 2 y", "Отклик 3 y",
        "Среднее y", "ŷ (ОМНК)", "Остаток e", "e²", "s(y_i) повт.",
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
        # s(y_i): из повторностей, пусто если повторностей < 2
        ws.cell(r, 10, f'=IF(B{r}="","",IF(COUNT(C{r}:E{r})>=2,STDEV(C{r}:E{r}),NA()))').fill = calc_fill

    # До 100 уровней калибровки. Пустые строки не участвуют в расчёте.
    for r in range(9, 109):
        ws.cell(r, 6, f'=IF(B{r}="","",AVERAGE(C{r}:E{r}))').fill = calc_fill
        ws.cell(
            r, 7,
            f'=IF(B{r}="","",\'Регрессия\'!$B$8*B{r}+\'Регрессия\'!$B$9)'
        ).fill = calc_fill
        ws.cell(r, 8, f'=IF(F{r}="","",F{r}-G{r})').fill = calc_fill
        ws.cell(r, 9, f'=IF(H{r}="","",H{r}^2)').fill = calc_fill

    dv_p = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="1")
    ws.add_data_validation(dv_p)
    dv_p.add(ws["B4"])

    dv_u = DataValidation(type="decimal", operator="between", formula1="0", formula2="1")
    ws.add_data_validation(dv_u)
    dv_u.add(ws["B5"])

    ws.freeze_panes = "A9"

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
        "Число точек калибровки n",
        "Среднее x̄",
        "Среднее ȳ",
        "Sxx = Σ(xᵢ−x̄)²",
        "Остаточная дисперсия s²y/x",
        "Наклон b1",
        "Сдвиг b0",
        "var(b1)",
        "var(b0)",
        "cov(b0,b1)",
        "R²",
        "b0/b1",
        "var(x_pred) ковариационная (полная)",
    ]
    formulas = [
        "=COUNT('Ввод'!B9:B108)",
        "=AVERAGE('Ввод'!B9:B108)",
        "=AVERAGE('Ввод'!F9:F108)",
        "=DEVSQ('Ввод'!B9:B108)",
        "=IF(B2>2,SUM('Ввод'!I9:I108)/(B2-2),NA())",
        "=SLOPE('Ввод'!F9:F108,'Ввод'!B9:B108)",
        "=INTERCEPT('Ввод'!F9:F108,'Ввод'!B9:B108)",
        "=B6/B5",
        "=B6*(1/B2+B3^2/B5)",
        "=-B6*B3/B5",
        "=RSQ('Ввод'!F9:F108,'Ввод'!B9:B108)",
        "=B8/B7",
        # Полная ковариационная форма: var(ŷ_obs)/b1² + (x̄−x_pred)²·var(b1)/b1²
        # тождественна классической s²/b1²·(1/p+1/n+(x_pred−x̄)²/Sxx).
        "=(B10+B6/'Ввод'!B4-2*B11)/B7^2+(B13-'Неопределенность'!B1)^2*B9/B7^2",
    ]
    for r, (label, formula) in enumerate(zip(labels, formulas), 2):
        reg.cell(r, 1, label).border = border
        reg.cell(r, 2, formula).fill = calc_fill
        reg.cell(r, 2).border = border

    # ============================================================
    # ЛИСТ 3. ВЗВЕШЕННАЯ РЕГРЕССИЯ (WLS)
    # Таблица: A9:D108 (x, y, w, wy); статистика: A111..
    # Карта итогов: B115 Sw, B116 Sxw, B117 Syw, B118 Sxxw, B119 Sxyw,
    #               B120 b1w, B121 b0w, B122 SSEw, B123 dof, B124 s2w,
    #               B125 var(b1w), B126 var(b0w), B127 covw,
    #               B128 x̄w, B129 Δb1%, B130 Δb0%
    # ============================================================
    wls = wb.create_sheet("Взвешенная регрессия")
    wls["A1"] = "ВЗВЕШЕННАЯ РЕГРЕССИЯ (WLS): РАВНЫЕ ВЕСА / 1/x / 1/x²"
    wls["A1"].font = Font(bold=True, size=15)
    wls["A2"] = 'Текущий режим (задаётся на листе «Ввод», B6):'
    wls["B2"] = '=CHOOSE(\'Ввод\'!$B$6,"равные веса (ОМНК)","w ∝ 1/x","w ∝ 1/x²")'
    wls["B2"].fill = calc_fill

    tbl_headers = ["x", "y (среднее)", "Вес w", "w·y"]
    for c, h in enumerate(tbl_headers, 1):
        wls.cell(10, c, h)
    style_header(wls, 10, len(tbl_headers))

    for i in range(100):
        r = 11 + i          # строка таблицы
        src = 9 + i         # строка на листе «Ввод»
        wls.cell(r, 1, f'=IF(\'Ввод\'!B{src}="","",\'Ввод\'!B{src})').fill = calc_fill
        wls.cell(r, 2, f'=IF(A{r}="","",\'Ввод\'!F{src})').fill = calc_fill
        wls.cell(
            r, 3,
            f'=IF(A{r}="","",CHOOSE(\'Ввод\'!$B$6,1,'
            f'IF(A{r}>0,1/A{r},NA()),IF(A{r}>0,1/A{r}^2,NA())))'
        ).fill = calc_fill
        wls.cell(r, 4, f'=IF(A{r}="","",C{r}*B{r})').fill = calc_fill

    stat_labels = [
        "Σw", "Σwx", "Σwy", "Sxx_w = Σw·x² − (Σwx)²/Σw", "Sxy_w = Σw·xy − Σwx·Σwy/Σw",
        "Наклон b1 (WLS)", "Сдвиг b0 (WLS)",
        "SSE_w = Σw·(y−ŵ)²", "Степени свободы (n_eff − 2)",
        "Взвешенная остаточная дисперсия φ_w",
        "var(b1)_w = φ_w/Sxx_w", "var(b0)_w", "cov(b0,b1)_w",
        "Взвешенное среднее x̄_w",
        "Δb1 vs ОМНК, %", "Δb0 vs ОМНК, %",
    ]
    stat_formulas = [
        "=SUM(C11:C110)",                                          # B115 Σw
        "=SUMPRODUCT(C11:C110,A11:C110)",                         # B116 Σwx
        "=SUM(D11:D110)",                                         # B117 Σwy
        "=SUMPRODUCT(C11:C110,A11:C110^2)-B116^2/B115",           # B118 Sxx_w
        "=SUMPRODUCT(C11:C110,A11:C110,B11:C110)-B116*B117/B115", # B119 Sxy_w
        "=B119/B118",                                             # B120 b1_w
        "=(B117-B120*B116)/B115",                                 # B121 b0_w
        "=SUMPRODUCT(C11:C110,(B11:C110-$B$120*A11:C110-$B$121)^2)",  # B122 SSE_w
        "=COUNT(A11:A110)-2",                                     # B123 dof
        "=IF(B123>0,B122/B123,NA())",                             # B124 φ_w
        "=B124/B118",                                             # B125 var(b1)_w
        "=B124*(1/B115+B116^2/(B115*B118))",                      # B126 var(b0)_w
        "=-B124*B116/(B115*B118)",                                # B127 cov_w
        "=B116/B115",                                             # B128 x̄_w
        "=IF('Регрессия'!B7=0,NA(),ABS(B120-'Регрессия'!B7)/ABS('Регрессия'!B7))",   # B129 Δb1
        "=IF('Регрессия'!B8=0,NA(),ABS(B121-'Регрессия'!B8)/MAX(ABS('Регрессия'!B8),1E-12))",  # B130 Δb0
    ]
    for r, (label, formula) in enumerate(zip(stat_labels, stat_formulas), 115):
        wls.cell(r, 1, label).border = border
        wls.cell(r, 2, formula).fill = calc_fill
        wls.cell(r, 2).border = border
    wls.cell(129, 2).number_format = "0.0%"
    wls.cell(130, 2).number_format = "0.0%"

    # Относительный вес точки (подсветка вклада уровней)
    wls.cell(9, 6, "Относительный вес уровня").font = Font(bold=True)
    for i in range(100):
        r = 11 + i
        wls.cell(r, 6, f'=IF(C{r}="","",C{r}/MAX($C$11:$C$110))').fill = calc_fill

    # ============================================================
    # ЛИСТ 4. НЕОПРЕДЕЛЁННОСТЬ (проба, сводка)
    # ============================================================
    unc = wb.create_sheet("Неопределенность")
    unc["A1"] = "ПРОГНОЗИРОВАНИЕ И НЕОПРЕДЕЛЁННОСТЬ (ТЕКУЩИЙ РЕЖИМ ВЕСОВ)"
    unc["A1"].font = Font(bold=True, size=15)

    WLS_B1 = "'Взвешенная регрессия'!$B$120"   # наклон WLS
    WLS_B0 = "'Взвешенная регрессия'!$B$121"   # сдвиг WLS

    items = [
        # 1. x_pred — по коэффициентам активного режима
        (
            "Предсказанная концентрация x_pred",
            f"=('Ввод'!B3-{WLS_B0})/{WLS_B1}",
        ),
        # 2. var(y_obs)
        (
            "Дисперсия среднего отклика пробы var(y_obs)",
            "=IF('Ввод'!$B$6=1,'Регрессия'!B6,'Взвешенная регрессия'!B124)/'Ввод'!B4",
        ),
        # 3. u²(x_pred) ОМНК, классическая форма
        (
            "u²(x_pred) ОМНК (классическая форма)",
            "=('Регрессия'!B6/'Регрессия'!B7^2)*(1/'Ввод'!B4+1/'Регрессия'!B2"
            "+(B1-'Регрессия'!B3)^2/'Регрессия'!B5)",
        ),
        # 4. u²(x_pred) ОМНК, ковариационная форма
        (
            "u²(x_pred) ОМНК (ковариационная форма)",
            "='Регрессия'!B14",
        ),
        # 5. u(x_cal)
        (
            "Абсолютная u эталонов u(x_cal)",
            "=ABS(B1)*'Ввод'!B5",
        ),
        # 6. u²(x_pred) WLS, интервальная форма
        (
            "u²(x_pred) текущий режим (WLS, интервальная форма)",
            "=IF('Ввод'!$B$6=1,B4,"
            "('Взвешенная регрессия'!$B$124*(1/'Ввод'!B4+1/COUNTIFS("
            "'Взвешенная регрессия'!A11:A110,\">=0\"))+"
            "(B1-'Взвешенная регрессия'!$B$128+'Взвешенная регрессия'!$B$121"
            "/'Взвешенная регрессия'!$B$120)^2*'Взвешенная регрессия'!$B$125+"
            "'Взвешенная регрессия'!$B$126+"
            "2*(B1-'Взвешенная регрессия'!$B$128+'Взвешенная регрессия'!$B$121"
            "/'Взвешенная регрессия'!$B$120)*'Взвешенная регрессия'!$B$127)"
            "/'Взвешенная регрессия'!$B$120^2)",
        ),
        # 7. суммарная
        (
            "Суммарная стандартная u_c(x_pred)",
            "=SQRT(B6+B5^2)",
        ),
        # 8. расширенная
        (
            "Расширенная U (k=2)",
            "=2*B7",
        ),
        # 9. относительная
        (
            "Относительная U",
            "=IF(B1=0,NA(),B8/ABS(B1))",
        ),
    ]
    for r, (label, formula) in enumerate(items, 2):
        unc.cell(r, 1, label).border = border
        unc.cell(r, 2, formula).fill = calc_fill
        unc.cell(r, 2).border = border

    # F2: x_pred по ОМНК (эталонное значение для сверки)
    unc["F1"] = "x_pred (ОМНК, справочно)"
    unc["F2"] = "('Ввод'!B3-'Регрессия'!B8)/'Регрессия'!B7"
    unc["F2"].fill = calc_fill

    unc["A12"] = "Контроль согласованности двух формул ОМНК"
    unc["B12"] = "=IFERROR(ABS(B4-B5)/MAX(ABS(B4),ABS(B5)),NA())"
    unc["A13"] = "Экстраполяция (границы диапазона калибровки)"
    unc["B13"] = "=OR(B1<MIN('Ввод'!B9:B108),B1>MAX('Ввод'!B9:B108))"

    # ============================================================
    # ЛИСТ 5. НЕОПРЕДЕЛЁННОСТЬ ПО ТОЧКАМ
    # Столбцы: A №, B x_i, C s(y_i), D u(y_i)|повторн., E n_i,
    #           F ŷ_i, G var(ŷ_i), H u(ŷ_i), I u_model(y_i),
    #           J u_combined(y_i), K u(x_i) обратн., L U(x_i) k=2
    # ============================================================
    pts = wb.create_sheet("Неопределенность по точкам")
    pts["A1"] = "НЕОПРЕДЕЛЁННОСТИ ПО УРОВНЯМ КАЛИБРОВКИ (ТЕКУЩИЙ РЕЖИМ ВЕСОВ)"
    pts["A1"].font = Font(bold=True, size=15)
    pts["A2"] = 'Режим весов:'
    pts["B2"] = '=CHOOSE(\'Ввод\'!$B$6,"ОМНК","1/x","1/x²")'
    pts["B2"].fill = calc_fill

    pt_headers = [
        "№", "x_i", "s(y_i) повт.", "u(y_i) повторн.", "n_i",
        "ŷ_i модель", "var(ŷ_i)", "u(ŷ_i)", "u(y_i) модель",
        "u(y_i) комбин.", "u(x_i) обратн.", "U(x_i) k=2",
    ]
    for c, h in enumerate(pt_headers, 1):
        pts.cell(4, c, h)
    style_header(pts, 4, len(pt_headers))

    # Ссылки на активный режим
    REG = {
        1: dict(b1="'Регрессия'!$B$7", b0="'Регрессия'!$B$8", s2="'Регрессия'!$B$6",
                n="'Регрессия'!$B$2", xbar="'Регрессия'!$B$3", sxx="'Регрессия'!$B$5",
                vb1="'Регрессия'!$B$9", vb0="'Регрессия'!$B$10", cov="'Регрессия'!$B$11"),
    }
    W = dict(b1="'Взвешенная регрессия'!$B$120", b0="'Взвешенная регрессия'!$B$121",
             s2="'Взвешенная регрессия'!$B$124", xbarw="'Взвешенная регрессия'!$B$128",
             vb1="'Взвешенная регрессия'!$B$125", vb0="'Взвешенная регрессия'!$B$126",
             cov="'Взвешенная регрессия'!$B$127")

    def pick(ols_ref, wls_ref):
        return f"IF('Ввод'!$B$6=1,{ols_ref},{wls_ref})"

    for i in range(100):
        r = 5 + i
        src = 9 + i  # строка во «Вводе»
        # A: номер
        pts.cell(r, 1, f'=IF(\'Ввод\'!B{src}="","",ROW()-4)').fill = calc_fill
        # B: x_i
        pts.cell(r, 2, f'=IF(\'Ввод\'!B{src}="","",\'Ввод\'!B{src})').fill = calc_fill
        # C: s(y_i) из Ввода
        pts.cell(r, 3, f'=IF(\'Ввод\'!J{src}="","",\'Ввод\'!J{src})').fill = calc_fill
        # D: u(y_i) из повторностей = s/sqrt(n_i)
        pts.cell(r, 4, f'=IF(C{r}="","",C{r}/SQRT(E{r}))').fill = calc_fill
        # E: n_i
        pts.cell(r, 5, f"=IF(B{r}=\"\",\"\",COUNT('Ввод'!C{src}:E{src}))").fill = calc_fill
        # F: ŷ_i (активный режим)
        pts.cell(r, 6, f'=IF(B{r}="","",{pick(REG[1]["b1"], W["b1"])}*B{r}+'
                       f'{pick(REG[1]["b0"], W["b0"])})').fill = calc_fill
        # G: var(ŷ_i) = s²·(1/n + (x−x̄)²/Sxx)  |  WLS: s²w·(1/Σw + (x−x̄w)²/Sxxw)
        pts.cell(
            r, 7,
            f'=IF(B{r}="","",'
            f'IF(\'Ввод\'!$B$6=1,'
            f"'Регрессия'!$B$6*(1/'Регрессия'!$B$2+(B{r}-'Регрессия'!$B$3)^2/'Регрессия'!$B$5),"
            f"'Взвешенная регрессия'!$B$124*(1/'Взвешенная регрессия'!$B$115+"
            f"(B{r}-'Взвешенная регрессия'!$B$128)^2/'Взвешенная регрессия'!$B$118)))"
        ).fill = calc_fill
        # H: u(ŷ_i)
        pts.cell(r, 8, f'=IF(G{r}="","",SQRT(G{r}))').fill = calc_fill
        # I: u_model(y_i) = sqrt(s² + var(ŷ_i)) — неопределённость одиночного отклика
        pts.cell(
            r, 9,
            f'=IF(B{r}="","",SQRT({pick(REG[1]["s2"], W["s2"])}+G{r}))'
        ).fill = calc_fill
        # J: комбинированная u(y_i): повторности + модель + эталоны x
        pts.cell(
            r, 10,
            f'=IF(B{r}="","",SQRT(N(D{r})^2+I{r}^2+(B{r}*\'Ввод\'!$B$5)^2))'
        ).fill = calc_fill
        # K: u(x_i) обратного предсказания (интервальная форма, по ŷ_i)
        dx_ols = f"(B{r}-'Регрессия'!$B$3)"
        dx_wls = f"(B{r}-'Взвешенная регрессия'!$B$128)"
        pts.cell(
            r, 11,
            f'=IF(B{r}="","",'
            f'SQRT('
            f'IF(\'Ввод\'!$B$6=1,'
            f"'Регрессия'!$B$6*(1/E{r}+1/'Регрессия'!$B$2+{dx_ols}^2/'Регрессия'!$B$5)+"
            f"'Регрессия'!$B$10+{dx_ols}^2*'Регрессия'!$B$9+2*{dx_ols}*'Регрессия'!$B$11,"
            f"'Взвешенная регрессия'!$B$124*(1/E{r}+1/COUNTIFS('Взвешенная регрессия'!$A$11:$A$110,\">=0\"))"
            f"+{dx_wls}^2*'Взвешенная регрессия'!$B$125+'Взвешенная регрессия'!$B$126"
            f"+2*{dx_wls}*'Взвешенная регрессия'!$B$127)"
            f"/{pick(REG[1]['b1'], W['b1'])}^2+(B{r}*'Ввод'!$B$5)^2))"
        ).fill = calc_fill
        # L: U(x_i)
        pts.cell(r, 12, f'=IF(K{r}="","",2*K{r})').fill = calc_fill

    pts.conditional_formatting.add(
        "L5:L104",
        ColorScaleRule(start_type="min", start_color="C6EFCE",
                       end_type="max", end_color="FFC7CE"),
    )

    # ============================================================
    # ЛИСТ 6. ПРОФИЛЬ U(x)
    # Сетка 200 точек; x_pred и его U(x) помечены флагом.
    # ============================================================
    prof = wb.create_sheet("Профиль U(x)")
    prof["A1"] = "ПРОФИЛЬ РАСШИРЕННОЙ НЕОПРЕДЕЛЁННОСТИ U(x) ПО ДИАПАЗОНУ"
    prof["A1"].font = Font(bold=True, size=15)

    ph = ["x", "dx = x − (x̄_act + b0_act/b1_act)", "u²(x) модель", "u(x) модель",
          "u_c(x) с учётом эталонов", "U(x) k=2", "это x_pred?"]
    for c, h in enumerate(ph, 1):
        prof.cell(3, c, h)
    style_header(prof, 3, len(ph))

    prof["A2"] = "Шаг сетки:"
    prof["B2"] = "=(MAX('Ввод'!B9:B108)-MIN('Ввод'!B9:B108))/199"
    prof["B2"].fill = calc_fill

    for i in range(200):
        r = 4 + i
        prof.cell(r, 1, "=MIN('Ввод'!$B$9:$B$108)+(ROW()-4)*$B$2").fill = calc_fill
        # dx: ОМНК — x−x̄; WLS — x − x̄_w + b0_w/b1_w (центр коррекции)
        prof.cell(
            r, 2,
            f'=IF(\'Ввод\'!$B$6=1,A{r}-\'Регрессия\'!$B$3,'
            f"A{r}-'Взвешенная регрессия'!$B$128+'Взвешенная регрессия'!$B$121"
            f"/'Взвешенная регрессия'!$B$120)"
        ).fill = calc_fill
        prof.cell(
            r, 3,
            f'=IF(\'Ввод\'!$B$6=1,'
            f"'Регрессия'!$B$6*(1/'Ввод'!$B$4+1/'Регрессия'!$B$2+B{r}^2/'Регрессия'!$B$5)+"
            f"'Регрессия'!$B$10+B{r}^2*'Регрессия'!$B$9+2*B{r}*'Регрессия'!$B$11,"
            f"'Взвешенная регрессия'!$B$124*(1/'Ввод'!$B$4+1/COUNTIFS('Взвешенная регрессия'!$A$11:$A$110,\">=0\"))"
            f"+B{r}^2*'Взвешенная регрессия'!$B$125+'Взвешенная регрессия'!$B$126"
            f"+2*B{r}*'Взвешенная регрессия'!$B$127)"
            f"/IF('Ввод'!$B$6=1,'Регрессия'!$B$7,'Взвешенная регрессия'!$B$120)^2"
        ).fill = calc_fill
        prof.cell(r, 4, f'=IF(C{r}<0,NA(),SQRT(C{r}))').fill = calc_fill
        prof.cell(r, 5, f"=SQRT(D{r}^2+(A{r}*'Ввод'!$B$5)^2)").fill = calc_fill
        prof.cell(r, 6, f"=2*E{r}").fill = calc_fill
        prof.cell(
            r, 7,
            f'=IF(ABS(A{r}-\'Неопределенность\'!$B$1)<=$B$2/2,"← x_pred","")'
        ).fill = calc_fill

    prof.conditional_formatting.add(
        "F4:F203",
        ColorScaleRule(start_type="min", start_color="C6EFCE",
                       end_type="max", end_color="FFC7CE"),
    )
    prof.conditional_formatting.add(
        "G4:G203",
        FormulaRule(formula=['$G4<>""'], fill=warn_fill),
    )

    # ============================================================
    # ЛИСТ 7. ВАЛИДАЦИЯ
    # ============================================================
    val = wb.create_sheet("Валидация")
    val["A1"] = "АВТОМАТИЧЕСКАЯ ДИАГНОСТИКА МОДЕЛИ"
    val["A1"].font = Font(bold=True, size=15)

    checks = [
        ("Достаточно степеней свободы (ОМНК)", "='Регрессия'!B2>2"),
        ("Sxx > 0", "='Регрессия'!B5>0"),
        ("Наклон b1 ≠ 0 (ОМНК)", "=ABS('Регрессия'!B7)>1E-12"),
        ("Нет экстраполяции", "=NOT('Неопределенность'!B13)"),
        ("Согласованность двух формул ОМНК", "='Неопределенность'!B12<1E-6"),
        ("R² ≥ 0.99 (диагностический порог)", "='Регрессия'!B12>=0.99"),
        ("Нулевой уровень x=0 задан (для весов 1/x обязателен)",
         "=COUNTIFS('Ввод'!B9:B108,0)>0"),
        ("Все веса положительны (нет #N/A в столбце w)",
         "=IF('Ввод'!$B$6=1,TRUE,COUNT('Взвешенная регрессия'!C11:C110)=COUNT('Взвешенная регрессия'!A11:A110))"),
        ("SSE_WLS ≤ SSE_ОМНК (взвешенная сумма квадратов меньше невозвешенной)",
         "=IF('Ввод'!$B$6=1,TRUE,'Взвешенная регрессия'!B122<=SUM('Ввод'!I9:I108)+1E-12)"),
        ("Δb1 (WLS vs ОМНК) < 10% — умеренное влияние весов",
         "=IF('Ввод'!$B$6=1,TRUE,'Взвешенная регрессия'!B129<0.1)"),
        ("a priori гетероскедастичность: s(y) растёт с x (кор. рангов > 0.5)",
         "=IF(COUNT('Ввод'!J9:J108)>=3,"
         "CORREL('Ввод'!B9:B108,'Ввод'!J9:J108)>0.5,TRUE)"),
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

    val.cell(15, 1, "Итог").font = Font(bold=True)
    val.cell(15, 2, '=IF(COUNTIF(C3:C13,"FAIL")=0,"МОДЕЛЬ ПРИГОДНА","ЕСТЬ ЗАМЕЧАНИЯ")')

    # Общий layout
    for sh in wb.worksheets:
        for row in sh.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
        sh.column_dimensions["A"].width = 50
        sh.column_dimensions["B"].width = 22
        sh.column_dimensions["C"].width = 20

    ws.column_dimensions["A"].width = 58
    for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J"]:
        ws.column_dimensions[col].width = 16
    for col in ["D", "E", "F", "G", "H", "I", "J", "K", "L"]:
        pts.column_dimensions[col].width = 15
        prof.column_dimensions[col].width = 15
    prof.column_dimensions["A"].width = 12

    wb.save(output_path)
    print(f"Файл успешно сгенерирован: {output_path}")
    return output_path


if __name__ == "__main__":
    create_metrology_excel_v3()
