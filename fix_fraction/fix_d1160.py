# -*- coding: utf-8 -*-
"""
ФР-05, ФР-06, + ДОВОДКА (переключение L по давлению) — астм д 1160.xlsx (ASTM D1160).

ФАКТЫ (проверены инспекцией 2026-10-01):
- Условие в M/N: IF('АСТМ Д 1160'!$E$1=Лист1!$N$2, ...) где Лист1!N2 = 1 (число).
  При E1=10 (10 мм рт.ст.) условие FALSE -> ветка 10 мм — ВЕРНО.
  При E1=1 -> TRUE -> ветка 1 мм — ВЕРНО.

- Константы 12.5.1 на Лист1 (для r):
   1 мм (0,13 кПа):  5-50%: a=D40(0.439), b=D41(0.241), M=D42(2.9)
                     60-95%: a=E40(0.439), b=E41(0.241), M=E42(3.0)
   10 мм (1,3 кПа):  5-50%: a=D44(0.240), b=D45(0.350), M=D46(2.8)
                     60-95%: a=E44(0.240), b=E45(0.350), M=E46(2.9)

ФР-05 (битые ссылки L): L12..L15 использовали Лист1!D47..D51 (несуществующие),
  L11 использовал E40/F41/E42 (смесь 1 мм). Исправлено.

ДОВОДКА: L5..L15 всегда использовали константы 10 мм (D42/D40/D41 для 5-50% —
  это на самом деле 1-мм константы!, и E46/E44/E45 для 60-95% — 10 мм).
  Теперь L5..L15 обёрнуты в IF(E1=N2, ветка 1 мм, ветка 10 мм):
   L5..L10 (5-50%):  1 мм -> D40/D41/D42; 10 мм -> D44/D45/D46
   L11..L15 (60-95%): 1 мм -> E40/E41/E42; 10 мм -> E44/E45/E46
  (Условие и структура — как в M/N: IF('АСТМ Д 1160'!$E$1=Лист1!$N$2, ...))

ФР-06 (смещённый MATCH в M): M11..M15 использовали MATCH(Лист2!G10..G14)
  вместо G11..G15. Выровнено.
"""
import os
import sys
import openpyxl
import re

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fix_common import D1160, norm_formula

# Строки 60-95%: row -> ожидаемый MATCH row
MATCH_ROWS = {11: 11, 12: 12, 13: 13, 14: 14, 15: 15}

COND = "'АСТМ Д 1160'!$E$1=Лист1!$N$2"


def formula_L_switch(row):
    """
    L-формула с переключением по давлению (как M/N).
    row 5..10 (5-50%): 1мм -> D40/D41/D42; 10мм -> D44/D45/D46.
    row 11..15 (60-95%): 1мм -> E40/E41/E42; 10мм -> E44/E45/E46.
    """
    if 5 <= row <= 10:
        mm1 = 'D$40', 'D$41', 'D$42'
        mm10 = 'D$44', 'D$45', 'D$46'
    else:  # 11..15
        mm1 = 'E$40', 'E$41', 'E$42'
        mm10 = 'E$44', 'E$45', 'E$46'
    a1, b1, m1 = mm1
    a10, b10, m10 = mm10
    return (
        "=IF(%s,"
        "Лист1!%s*(2.718281828*EXP((Лист1!%s+Лист1!%s*LN((1.8*Лист2!G%d)))))/1.8,"
        "Лист1!%s*(2.718281828*EXP((Лист1!%s+Лист1!%s*LN((1.8*Лист2!G%d)))))/1.8)"
        % (COND, m1, a1, b1, row, m10, a10, b10, row)
    )


def fix_MATCH(formula, row):
    """Заменить MATCH(Лист2!G{k}) на MATCH(Лист2!G{row}) в формуле M/N."""
    return re.sub(r'MATCH\(Лист2!G(\d+)', 'MATCH(Лист2!G%d' % row, formula)


def fix_file(path):
    print('\n==== D1160: %s ====' % os.path.basename(path))
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb['Лист2']
    changes = []

    # ---------- ДОВОДКА: L5..L15 переключение по давлению ----------
    for row in range(5, 16):
        old_l = ws['L%d' % row].value
        new_l = formula_L_switch(row)
        if norm_formula(old_l) != norm_formula(new_l):
            ws['L%d' % row] = new_l
            changes.append(('ДОВОДКА-L', 'Лист2!L%d' % row, old_l, new_l))
        else:
            print('  L%d уже корректна (пропуск)' % row)

    # ---------- ФР-06: M11..M15 MATCH выровнен ----------
    for row, mrow in MATCH_ROWS.items():
        old_m = ws['M%d' % row].value
        new_m = fix_MATCH(old_m, mrow)
        if norm_formula(old_m) != norm_formula(new_m):
            ws['M%d' % row] = new_m
            changes.append(('ФР-06', 'Лист2!M%d' % row, old_m, new_m))
        else:
            print('  ФР-06: M%d уже корректна (пропуск)' % row)

    # ---------- Контроль N ----------
    for row, mrow in MATCH_ROWS.items():
        old_n = ws['N%d' % row].value
        mg = re.findall(r'MATCH\(Лист2!G(\d+)', old_n or '')
        if mg and all(int(x) == row for x in mg):
            print('  N%d: MATCH корректен (%s)' % (row, mg))
        else:
            new_n = fix_MATCH(old_n, row)
            if norm_formula(old_n) != norm_formula(new_n):
                ws['N%d' % row] = new_n
                changes.append(('ФР-06', 'Лист2!N%d' % row, old_n, new_n))

    wb.save(path)
    print('  SAVED:', path)
    wb.close()
    return changes


def main():
    changes = fix_file(D1160)
    print('\nИТОГО изменений:', len(changes))
    for c in changes:
        print('  %s | %s | было=%r | стало=%r' % c[:4])
    return 0


if __name__ == '__main__':
    sys.exit(main())