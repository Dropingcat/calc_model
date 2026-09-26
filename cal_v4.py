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
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation


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
    # ЛИСТ 1. ВВОД
    # Данные: B9:B108 x; C:E отклики; F среднее; J s(y_i).
    # Параметры: B3 y_obs, B4 p, B5 u_rel эталонов, B6 режим весов,
    # B7 доверительная вероятность.
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
        ("Доверительная вероятность для U", 0.95),
    ]
    for r, (name, value) in enumerate(params, 3):
        ws.cell(r, 1, name)
        ws.cell(r, 2, value).fill = input_fill
    ws["B5"].number_format = "0.00%"
    ws["B7"].number_format = "0.00%"

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
        "Число точек калибровки n", "Среднее x̄", "Среднее ȳ",
        "Sxx = Σ(xᵢ−x̄)²", "Остаточная дисперсия s²y/x",
        "Наклон b1", "Сдвиг b0",
        "var(b1)", "var(b0)", "cov(b0,b1)", "R²",
        "b0/b1", "var(x_pred) ковариационная (полная)",
    ]
    formulas = [
        "=COUNT('Ввод'!$B$9:$B$108)",
        "=AVERAGE('Ввод'!$B$9:$B$108)",
        "=AVERAGE('Ввод'!$F$9:$F$108)",
        "=DEVSQ('Ввод'!$B$9:$B$108)",
        "=IF(B2>2,SUM('Ввод'!$I$9:$I$108)/(B2-2),NA())",
        "=SLOPE('Ввод'!$F$9:$F$108,'Ввод'!$B$9:$B$108)",
        "=INTERCEPT('Ввод'!$F$9:$F$108,'Ввод'!$B$9:$B$108)",
        "=B6/B5",
        "=B6*(1/B2+B3^2/B5)",
        "=-B6*B3/B5",
        "=RSQ('Ввод'!$F$9:$F$108,'Ввод'!$B$9:$B$108)",
        "=B8/B7",
        "=(B10+B6/'Ввод'!$B$4+'Неопределенность'!$B$1^2*B9"
        "+2*'Неопределенность'!$B$1*B11)/B7^2",
    ]
    for r, (label, formula) in enumerate(zip(labels, formulas), 2):
        reg.cell(r, 1, label).border = border
        reg.cell(r, 2, formula).fill = calc_fill
        reg.cell(r, 2).border = border

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
    ]
    for c, h in enumerate(tbl_headers, 1):
        wls.cell(8, c, h)
    style_header(wls, 8, len(tbl_headers))

    # TD-01: IF(A>0, 1/A^k, "") — пустая ячейка веса вместо #N/A.
    for i in range(100):
        r = 9 + i
        src = 9 + i
        wls.cell(r, 1, f'=IF(\'Ввод\'!B{src}="","",\'Ввод\'!B{src})').fill = calc_fill
        wls.cell(r, 2, f'=IF(A{r}="","",\'Ввод\'!F{src})').fill = calc_fill
        wls.cell(
            r, 3,
            f'=IF(A{r}="","",CHOOSE(\'Ввод\'!$B$6,1,'
            f'IF(A{r}>0,1/A{r},""),IF(A{r}>0,1/A{r}^2,"")))'
        ).fill = calc_fill
        wls.cell(r, 4, f'=IF(C{r}="","",C{r}*B{r})').fill = calc_fill
        # I: 1 если точка участвует в WLS (числовой вес), иначе 0
        wls.cell(r, 5, f'=IF(ISNUMBER(C{r}),1,0)').fill = calc_fill
        wls.cell(r, 6, f'=IF(C{r}="","",C{r}/MAX($C$9:$C$108))').fill = calc_fill
        wls.cell(r, 7, f'=IF(A{r}="","",A{r}*B{r})').fill = calc_fill          # xy
        wls.cell(r, 8, f'=IF(C{r}="","",C{r}*A{r}*A{r})').fill = calc_fill     # w·x²
        wls.cell(r, 9, f'=IF(C{r}="","",C{r}*C{r})').fill = calc_fill          # w²
        wls.cell(r, 10, f'=IF(I{r}="","",I{r}*A{r}*A{r})').fill = calc_fill    # w²·x²
        # e_w — остаток при текущих коэффициентах (зависит от B122/B123)
        wls.cell(r, 11, f'=IF(B{r}="","",B{r}-$B$122*A{r}-$B$123)').fill = calc_fill
        # L: w·e² — безопасно для пустых весов ("" не даст #VALUE!)
        wls.cell(r, 12,
                 f'=IF(OR(C{r}="",K{r}=""),"",C{r}*K{r}*K{r})').fill = calc_fill
        # M: вес только для участвующих точек (0 вместо "" — SUMPRODUCT-safe)
        wls.cell(r, 13, f'=IF(ISNUMBER(C{r}),C{r},0)').fill = calc_fill           # w_участн.

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
        # TD-05: все суммы — по строкам с I=1 (множитель E в SUMPRODUCT).
        "=SUMIF($E$9:$E$108,1,$C$9:$C$108)",                         # B115 Σw
        "=SUMPRODUCT($E$9:$E$108,$C$9:$C$108,$A$9:$A$108)",           # B116 Σwx
        "=SUM($D$9:$D$108)",                                          # B117 Σwy
        "=SUMPRODUCT($E$9:$E$108,$H$9:$H$108)-B116^2/B115",          # B118 Sxx_w
        "=SUMPRODUCT($E$9:$E$108,$C$9:$C$108,$G$9:$G$108)"
        "-B116*B117/B115",                                            # B119 Sxy_w
        "=SUM($E$9:$E$108)",                                          # B120 n_eff
        "=IF(B120>2,B120-2,NA())",                                    # B121 dof
        "=B119/B118",                                                 # B122 b1_w
        "=(B117-B122*B116)/B115",                                     # B123 b0_w
        "=SUM($L$9:$L$108)",                                         # B124 SSE_w
        "=IF(B121>0,B124/B121,NA())",                                 # B125 φ_w
        "=B125/B118",                                                 # B126 var(b1)_w
        # TD-02: полная форма: var(b0)_w = φ_w·Σw²x²/(Σw·Sxx_w)
        "=B125*B132/(B115*B118)",                                     # B127 var(b0)_w
        "=-B125*B116/(B115*B118)",                                    # B128 cov_w
        "=B116/B115",                                                 # B129 x̄_w
        "=IF('Регрессия'!B7=0,NA(),ABS(B122-'Регрессия'!B7)/ABS('Регрессия'!B7))",   # B130
        "=IF('Регрессия'!B8=0,NA(),ABS(B123-'Регрессия'!B8)"
        "/MAX(ABS('Регрессия'!B8),1E-12))",                           # B131
        "=SUM($J$9:$J$108)",                                          # B132 Σw²x²
        # TD-02: D≈1 если модель дисперсии верна; D≫1 → интервалы занижены
        "=IF('Регрессия'!B6>0,B125/'Регрессия'!B6,NA())",             # B133
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
        "=COUNT($A$8:$A$107)",                                         # B110 m
        "=SUM($D$8:$D$107)",                                           # B111 SS_PE
        "=COUNTIFS($C$8:$C$107,\">=2\",$C$8:$C$107,\"<>\")"
        "*0+SUMIFS($C$8:$C$107,$C$8:$C$107,\">=2\")"
        "-COUNTIFS($C$8:$C$107,\">=2\")",                              # B112 df_PE
        "=SUM($G$8:$G$107)",                                           # B113 SS_LoF OLS
        "=IF(B110>2,B110-2,NA())",                                     # B114 df_LoF
        "=IF(AND(B112>0,B114>0),(B113/B114)/(B111/B112),NA())",        # B115 F OLS
        "=IF(ISNUMBER(B115),FINV(1-'Ввод'!$B$7,B114,B112),NA())",      # B116 Fcrit
        "=IF(ISNUMBER(B115),IF(B115<=B116,\"PASS: линейность адекватна\","
        "\"FAIL: lack of fit — модель непригодна\"),\"нет повторностей\")",  # B117
        "=SUM($H$8:$H$107)",                                           # B118 SS_LoF W
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
            "=IF(AND('Ввод'!$B$6>1,COUNT($J$8:$J$107)>=4),"
            "(SUM($N$8:$N$107)"
            "-SUM($J$8:$J$107)*SUM($K$8:$K$107)/COUNT($J$8:$J$107))"
            "/((SUM($L$8:$L$107)-SUM($J$8:$J$107)^2/COUNT($J$8:$J$107))^0.5"
            "*(SUM($M$8:$M$107)-SUM($K$8:$K$107)^2/COUNT($J$8:$J$107))^0.5)"
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
    d_wls = f"$B$1-{WW['xbarw']}+{WW['b0']}/{WW['b1']}"
    u2_model = (
        f"IF('Ввод'!$B$6=1,"
        f"{OO['s2']}*(1/'Ввод'!$B$4+1/{OO['n']}+({d_ols})^2/{OO['Sxx']}),"
        f"{WW['phi']}*(1/'Ввод'!$B$4+1/{WW['Sw']})"
        f"+({d_wls})^2*{WW['vb1']}+{WW['vb0']}+2*({d_wls})*{WW['cov']})"
        f"/IF('Ввод'!$B$6=1,{OO['b1']},{WW['b1']})^2"
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
        ("ν_eff (Welch–Satterthwaite; ν_эталонов=50)",
         "=IF(B5=0,NA(),B5^4/(B3^4/MAX(B10,1)+IF(B4=0,0,B4^4/50)))", "0.0"),
        ("k (Стьюдент, TINV, двусторонний) — TD-04",
         "=IF(B11>0,TINV(1-'Ввод'!$B$7,MAX(ROUND(B11,0),1)),2)", "0.00"),
        ("U(x_pred) расширенная = k·u_c", "=B12*B5", ""),
        ("U/x_pred относительная", "=IF(B1=0,NA(),B13/ABS(B1))", "0.0%"),
        ("Экстраполяция? (x_pred вне диапазона)",
         "=OR(B1<MIN('Ввод'!$B$9:$B$108),B1>MAX('Ввод'!$B$9:$B$108))", ""),
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
        d_wls_r = f"(B{r}-{WW['xbarw']}+{WW['b0']}/{WW['b1']})"
        pts.cell(
            r, 11,
            f'=IF(B{r}="","",SQRT('
            f"IF('Ввод'!$B$6=1,"
            f"{OO['s2']}*(1/E{r}+1/{OO['n']}+{d_ols_r}^2/{OO['Sxx']}),"
            f"{WW['phi']}*(1/E{r}+1/{WW['Sw']})"
            f"+{d_wls_r}^2*{WW['vb1']}+{WW['vb0']}+2*{d_wls_r}*{WW['cov']})"
            f"/IF('Ввод'!$B$6=1,{OO['b1']},{WW['b1']})^2"
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
    prof["B2"] = "=(MAX('Ввод'!$B$9:$B$108)-MIN('Ввод'!$B$9:$B$108))/199"
    prof["B2"].fill = calc_fill

    for i in range(200):
        r = 4 + i
        prof.cell(r, 1,
                  "=MIN('Ввод'!$B$9:$B$108)+(ROW()-4)*$B$2").fill = calc_fill
        prof.cell(
            r, 2,
            f'=IF(\'Ввод\'!$B$6=1,A{r}-{OO["xbar"]},'
            f"A{r}-{WW['xbarw']}+{WW['b0']}/{WW['b1']})"
        ).fill = calc_fill
        prof.cell(
            r, 3,
            f"=IF('Ввод'!$B$6=1,"
            f"{OO['s2']}*(1/'Ввод'!$B$4+1/{OO['n']}+B{r}^2/{OO['Sxx']}),"
            f"{WW['phi']}*(1/'Ввод'!$B$4+1/{WW['Sw']})"
            f"+B{r}^2*{WW['vb1']}+{WW['vb0']}+2*B{r}*{WW['cov']})"
            f"/IF('Ввод'!$B$6=1,{OO['b1']},{WW['b1']})^2"
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
    # ЛИСТ 8. ВАЛИДАЦИЯ (16 проверок, включая TD-01…TD-04)
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
         f"{WW['neff']}=COUNT('Взвешенная регрессия'!$A$9:$A$108)"
         "-COUNTIFS('Взвешенная регрессия'!$A$9:$A$108,0))"),
        ("Все веса положительны, NaN нет (TD-01)",
         "=IF('Ввод'!$B$6=1,TRUE,"
         "SUMPRODUCT('Взвешенная регрессия'!$E$9:$E$108,"
         "'Взвешенная регрессия'!$C$9:$C$108<=0)=0)"),
        ("n_eff(WLS) ≥ 4 (достаточно для φ_w)",
         "=IF('Ввод'!$B$6=1,TRUE," + WW['neff'] + ">=4)"),
        ("Модель дисперсии адекватна: 0.5 < D < 2 (TD-02)",
         "=IF('Ввод'!$B$6=1,TRUE,"
         "AND('Диагностика'!B124>0.5,'Диагностика'!B124<2))"),
        ("Δb1 (WLS vs ОМНК) < 10%",
         "=IF('Ввод'!$B$6=1,TRUE," + WW['db1'] + "<0.1)"),
        ("a priori гетероскедастичность: corr(x, s_y) > 0.5",
         "=IF(COUNT('Ввод'!$J$9:$J$108)>=3,"
         "CORREL('Ввод'!$B$9:$B$108,'Ввод'!$J$9:$J$108)>0.5,TRUE)"),
        ("ν_eff ≥ 5 (корректный k по Стьюденту, TD-04)",
         "=IF(ISNUMBER('Неопределенность'!B11),"
         "'Неопределенность'!B11>=5,FALSE)"),
        ("k в разумных границах 2…12 (TD-04)",
         "=IF(ISNUMBER('Неопределенность'!B12),"
         "AND('Неопределенность'!B12>=2,'Неопределенность'!B12<=12),FALSE)"),
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

    wb.save(output_path)
    print(f"Файл успешно сгенерирован: {output_path}")
    return output_path


if __name__ == "__main__":
    create_metrology_excel_v4()
