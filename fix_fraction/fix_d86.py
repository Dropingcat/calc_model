# -*- coding: utf-8 -*-
"""
ФР-02, ФР-03 — АСТМ Д 86-1-4.xlsx.

ФР-02: 'расчёт'!P5..P24 — воспроизводимость R группы 1 по Table 8 ASTM D86-20a:
       5% -> 3.1 + 1.74*L (сейчас 2.0+)
       10-80% -> 2.0 + 1.74*L (верно, не трогаем)
       90% -> 0.8 + 1.74*L (верно, не трогаем)
       95% -> 1.1 + 1.74*L (сейчас 2.0+)
       FBP(100%) -> 7.2 (верно, не трогаем)
       Условие группы: IF('вывод данных'!G$5=расчёт!W$3, ..., (1.33*L+1.8))

ФР-03: блок температур B36:B47 — K37..K47, L36..L47, M36..M47, N/O/Q —
       обернуть расчёт в IFERROR(...,"—") чтобы при B=0 (температура не введена)
       не было #DIV/0!. Для заполненных строк (B36 НК, B37, B46, B47) логика не меняется.
"""
import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fix_common import D86, norm_formula

# Проценты по строкам P (B5..B24 = 5..100)
PCT_BY_ROW = {5: 5, 6: 10, 8: 20, 10: 30, 12: 40, 14: 50, 16: 60, 18: 70, 20: 80, 22: 90, 23: 95, 24: 100}
# Для группы 1 (Table 8): свободный член R по % отгона
R_COEF_BY_PCT = {
    5: 3.1, 10: 2.0, 20: 2.0, 30: 2.0, 40: 2.0, 50: 2.0,
    60: 2.0, 70: 2.0, 80: 2.0, 90: 0.8, 95: 1.1, 100: 7.2,
}


def formula_R_d86(row, pct):
    """Формула R (колонка P) для строки row и процента pct (группа 1)."""
    coef = R_COEF_BY_PCT[pct]
    if pct == 100:
        return "=ROUNDUP(INT((IF('вывод данных'!G5=расчёт!W3,(7.2),(0.42*L%d+3.1)))*2+0.4999)/2,2)" % row
    return ("=ROUNDUP(INT((IF('вывод данных'!G$5=расчёт!W$3,((1.74*L%d+%s)),"
            "(1.33*L%d+1.8)))*2+0.4999)/2,2)") % (row, coef, row)


def fix_file(path):
    print('\n==== D86: %s ====' % os.path.basename(path))
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb['расчёт']
    changes = []

    # ---------- ФР-02: P5 и P23 (5% и 95%) ----------
    w3 = ws['W3'].value
    if w3 != 'Группа 1':
        print('  !! ВНИМАНИЕ: W3 =', repr(w3), '— ожидалось "Группа 1"')
    for row in (5, 23):
        old_p = ws['P%d' % row].value
        new_p = formula_R_d86(row, PCT_BY_ROW[row])
        if norm_formula(old_p) != norm_formula(new_p):
            ws['P%d' % row] = new_p
            changes.append(('ФР-02', 'расчёт!P%d' % row, old_p, new_p))
        else:
            print('  ФР-02: P%d уже корректна (пропуск)' % row)
    # Убедиться, что P6..P20 (10-80%) и P22 (90%), P24 (FBP) не тронуты
    for row in (6, 8, 10, 12, 14, 16, 18, 20, 22, 24):
        if row in (5, 23):
            continue
        old_p = ws['P%d' % row].value
        expected = formula_R_d86(row, PCT_BY_ROW[row])
        if norm_formula(old_p) != norm_formula(expected):
            print('  !! P%d отличается от эталона (не трогаем, но фиксируем):' % row)
            print('     было   :', repr(old_p))
            print('     эталон :', repr(expected))
    print('  ФР-02: P5=3.1+, P23=1.1+ (остальные группы 1 без изменений)')

    # ---------- ФР-03: IFERROR в блоке температур ----------
    # K37..K47 — крутизна; L36..L47, M36..M47 — r/R; N/O/Q производные.
    # Стратегия: если в строке B=0 или формула K даёт #DIV/0! при пустых данных —
    # обернуть в IFERROR(...,"—"). Чтобы не менять логику заполненных строк,
    # оборачиваем КАЖДУЮ формулу K/L/M/P (для строк 36..47) в IFERROR.
    # Проверка: для заполненных строк (B36,B37,B46,B47) формула не изменит результат.
    temp_rows = range(36, 48)
    for r in temp_rows:
        for col in ('K', 'L', 'M'):
            cell = ws['%s%d' % (col, r)]
            v = cell.value
            if v is None or not (isinstance(v, str) and v.startswith('=')):
                continue
            if 'IFERROR' in v.upper():
                continue
            cell.value = '=IFERROR(%s,"—")' % v[1:]  # убрать '=' и обернуть
            changes.append(('ФР-03', 'расчёт!%s%d' % (col, r), v, cell.value))
    # N,O,Q уже содержат IFERROR (проверено) — пропускаем
    print('  ФР-03: K/L/M 36..47 обёрнуты в IFERROR (N/O/Q уже были)')

    wb.save(path)
    print('  SAVED:', path)
    wb.close()
    return changes


def main():
    changes = fix_file(D86)
    print('\nИТОГО изменений:', len(changes))
    for c in changes:
        print('  %s | %s | было=%r | стало=%r' % c[:4])
    return 0


if __name__ == '__main__':
    sys.exit(main())