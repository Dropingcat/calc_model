# -*- coding: utf-8 -*-
"""
ФР-01, ФР-04, ФР-16 — ГОСТ ISO 3405 / ГОСТ ИСО 3405 (два файла).

ФР-04: 'Группы 1-4'!D23: '='Ввод данных'!D23327' -> '='Ввод данных'!D23'
ФР-01: 'Лист1'!O5..O23 — поправки воспроизводимости R группы 1 по табл.6 ГОСТ ISO 3405-2022:
       5% -> R1+1.11; 10-80% -> R1 (без константы); 90% -> R1-1.22; 95% -> R1-0.94
       Условие группы: IF('Группы 1-4'!G$5=Лист1!V$3, ..., (1.33*K+1.8))
ФР-16: 'Группы 1-4'!K6..K26: '=Лист1!G4..G24' -> '=Лист1!I4..I24' (скорректированный на потери % отгона)
"""
import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fix_common import ISO1, ISO2, norm_formula

# Проценты отгона по строкам O на Лист1 (A5..A23 = 5..95)
PCT_BY_ROW = {5: 5, 6: 10, 8: 20, 10: 30, 12: 40, 14: 50, 16: 60, 18: 70, 20: 80, 22: 90, 23: 95}

# Поправка для R1 = 1.736*K + 1.994 по % отгона (табл.6 ГОСТ ISO 3405-2022)
# None -> без константы
CORR_BY_PCT = {
    5: '+1.11',
    10: None, 20: None, 30: None, 40: None, 50: None, 60: None, 70: None, 80: None,
    90: '-1.22',
    95: '-0.94',
}


def formula_R(row, pct):
    """Собрать формулу R (колонка O) для строки row и процента pct."""
    corr = CORR_BY_PCT.get(pct)
    if corr is None:
        r1 = '((1.736*K%d+1.994))' % row
    else:
        r1 = '((1.736*K%d+1.994)%s)' % (row, corr)
    return "=ROUNDUP(INT((IF('Группы 1-4'!G$5=Лист1!V$3,%s,(1.33*K%d+1.8)))*2+0.4999)/2,2)" % (r1, row)


def fix_file(path, tag):
    print('\n==== %s: %s ====' % (tag, os.path.basename(path)))
    wb = openpyxl.load_workbook(path, data_only=False)
    changes = []

    # ---------- ФР-04: D23 на листе 'Группы 1-4' ----------
    ws_g = wb['Группы 1-4']
    old_d23 = ws_g['D23'].value
    if norm_formula(old_d23) == norm_formula("='Ввод данных'!D23327"):
        ws_g['D23'] = "='Ввод данных'!D23"
        changes.append(('ФР-04', 'Группы 1-4!D23', old_d23, ws_g['D23'].value))
        print('  ФР-04: D23 исправлена')
    elif norm_formula(old_d23) == norm_formula("='Ввод данных'!D23"):
        print('  ФР-04: D23 уже корректна (пропуск)')
    else:
        print('  ФР-04: НЕОЖИДАННОЕ значение D23 =', repr(old_d23))
        # не блокируем — просто записываем в отчёт

    # ---------- ФР-16: K6..K26 на 'Группы 1-4' = Лист1!I4..I24 ----------
    for i, row in enumerate(range(6, 27)):
        src_row = 4 + i  # I4..I24 (K6->I4, K7->I5, ..., K26->I24)
        old_k = ws_g['K%d' % row].value
        new_k = '=Лист1!I%d' % src_row
        if norm_formula(old_k) != norm_formula(new_k) and norm_formula(old_k).startswith('=ЛИСТ1!G'):
            ws_g['K%d' % row] = new_k
            changes.append(('ФР-16', 'Группы 1-4!K%d' % row, old_k, new_k))
        elif norm_formula(old_k) == norm_formula(new_k):
            pass  # уже
        else:
            print('  ФР-16: K%d неожиданно: %r (пропуск)' % (row, old_k))
    print('  ФР-16: K6..K26 -> Лист1!I4..I24 (%d ячеек)' % sum(1 for c in changes if c[0] == 'ФР-16'))

    # ---------- ФР-01: O5..O23 на 'Лист1' ----------
    ws1 = wb['Лист1']
    # проверить условие группы — V3 должно быть 'Группа 1'
    v3 = ws1['V3'].value
    if v3 != 'Группа 1':
        print('  !! ВНИМАНИЕ: V3 =', repr(v3), '— ожидалось "Группа 1"')
    fr01 = 0
    for row, pct in PCT_BY_ROW.items():
        old_o = ws1['O%d' % row].value
        new_o = formula_R(row, pct)
        if norm_formula(old_o) != norm_formula(new_o):
            ws1['O%d' % row] = new_o
            changes.append(('ФР-01', 'Лист1!O%d' % row, old_o, new_o))
            fr01 += 1
        else:
            print('  ФР-01: O%d уже корректна (пропуск)' % row)
    print('  ФР-01: изменено %d ячеек колонки O' % fr01)

    wb.save(path)
    print('  SAVED:', path)
    wb.close()
    return changes


def main():
    all_changes = []
    all_changes += fix_file(ISO1, 'ISO1')
    all_changes += fix_file(ISO2, 'ISO2')
    print('\nИТОГО изменений:', len(all_changes))
    for c in all_changes:
        print('  %s | %s | было=%r | стало=%r' % c[:4])
    return 0


if __name__ == '__main__':
    sys.exit(main())