import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation


def create_metrology_excel(output_path="Metrology_Core_EURACHEM_v2.xlsx"):
    wb = Workbook()

    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="4F81BD")
    input_fill = PatternFill("solid", fgColor="E2EFDA")
    calc_fill = PatternFill("solid", fgColor="DDEBF7")
    pass_fill = PatternFill("solid", fgColor="C6EFCE")
    fail_fill = PatternFill("solid", fgColor="FFC7CE")

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
    ]
    for r, (name, value) in enumerate(params, 3):
        ws.cell(r, 1, name)
        ws.cell(r, 2, value).fill = input_fill
    ws["B5"].number_format = "0.00%"

    headers = [
        "№", "Концентрация x", "Отклик 1 y", "Отклик 2 y", "Отклик 3 y",
        "Среднее y", "ŷ", "Остаток e", "e²",
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
    # ЛИСТ 2. РЕГРЕССИЯ
    # Карта ячеек:
    # B2 n, B3 xmean, B4 ymean, B5 Sxx, B6 s2_yx,
    # B7 b1, B8 b0, B9 var(b1), B10 var(b0), B11 cov, B12 R2
    # ============================================================
    reg = wb.create_sheet("Регрессия")
    reg["A1"] = "СТАТИСТИКИ ЛИНЕЙНОЙ РЕГРЕССИИ (МНК)"
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
    ]
    for r, (label, formula) in enumerate(zip(labels, formulas), 2):
        reg.cell(r, 1, label).border = border
        reg.cell(r, 2, formula).fill = calc_fill
        reg.cell(r, 2).border = border

    # ============================================================
    # ЛИСТ 3. НЕОПРЕДЕЛЁННОСТЬ
    # ============================================================
    unc = wb.create_sheet("Неопределенность")
    unc["A1"] = "ПРОГНОЗИРОВАНИЕ И НЕОПРЕДЕЛЁННОСТЬ"
    unc["A1"].font = Font(bold=True, size=15)

    items = [
        (
            "Предсказанная концентрация x_pred",
            "=('Ввод'!B3-'Регрессия'!B8)/'Регрессия'!B7",
        ),
        (
            "Дисперсия среднего отклика пробы var(y_obs)",
            "='Регрессия'!B6/'Ввод'!B4",
        ),
        (
            "u²(x_pred), ковариационный",
            "=(B3+'Регрессия'!B10+B2^2*'Регрессия'!B9+2*B2*'Регрессия'!B11)/'Регрессия'!B7^2",
        ),
        (
            "u²(x_pred), классический",
            "=('Регрессия'!B6/'Регрессия'!B7^2)*(1/'Ввод'!B4+1/'Регрессия'!B2+(B2-'Регрессия'!B3)^2/'Регрессия'!B5)",
        ),
        (
            "Абсолютная u эталонов u(x_cal)",
            "=ABS(B2)*'Ввод'!B5",
        ),
        (
            "Суммарная стандартная u_c(x_pred)",
            "=SQRT(B5+B6^2)",
        ),
        (
            "Расширенная U (k=2)",
            "=2*B7",
        ),
        (
            "Относительная U",
            "=IF(B2=0,NA(),B8/ABS(B2))",
        ),
    ]
    for r, (label, formula) in enumerate(items, 2):
        unc.cell(r, 1, label).border = border
        unc.cell(r, 2, formula).fill = calc_fill
        unc.cell(r, 2).border = border

    unc["A11"] = "Контроль согласованности двух формул"
    unc["B11"] = "=IFERROR(ABS(B4-B5)/MAX(ABS(B4),ABS(B5)),NA())"
    unc["A12"] = "Экстраполяция"
    unc["B12"] = "=OR(B2<MIN('Ввод'!B9:B108),B2>MAX('Ввод'!B9:B108))"

    # ============================================================
    # ЛИСТ 4. ВАЛИДАЦИЯ
    # ============================================================
    val = wb.create_sheet("Валидация")
    val["A1"] = "АВТОМАТИЧЕСКАЯ ДИАГНОСТИКА МОДЕЛИ"
    val["A1"].font = Font(bold=True, size=15)

    checks = [
        ("Достаточно степеней свободы", "='Регрессия'!B2>2"),
        ("Sxx > 0", "='Регрессия'!B5>0"),
        ("Наклон b1 ≠ 0", "=ABS('Регрессия'!B7)>1E-12"),
        ("Нет экстраполяции", "=NOT('Неопределенность'!B12)"),
        ("Согласованность двух формул", "='Неопределенность'!B11<1E-6"),
        ("R² ≥ 0.99 (диагностический порог)", "='Регрессия'!B12>=0.99"),
    ]
    for r, (name, formula) in enumerate(checks, 3):
        val.cell(r, 1, name).border = border
        val.cell(r, 2, formula).fill = calc_fill
        val.cell(r, 2).border = border
        val.cell(r, 3, f'=IF(B{r},"PASS","FAIL")').border = border
        val.conditional_formatting.add(
            f"C{r}",
            FormulaRule(formula=[f'C{r}="PASS"'], fill=pass_fill),
        )
        val.conditional_formatting.add(
            f"C{r}",
            FormulaRule(formula=[f'C{r}="FAIL"'], fill=fail_fill),
        )

    # Общий layout
    for sh in wb.worksheets:
        for row in sh.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
        sh.column_dimensions["A"].width = 50
        sh.column_dimensions["B"].width = 22
        sh.column_dimensions["C"].width = 20

    ws.column_dimensions["A"].width = 58
    for col in ["B", "C", "D", "E", "F", "G", "H", "I"]:
        ws.column_dimensions[col].width = 16

    wb.save(output_path)
    print(f"Файл успешно сгенерирован: {output_path}")
    return output_path


if __name__ == "__main__":
    create_metrology_excel()
