# -*- coding: utf-8 -*-
"""
ФР-13 — ГОСТ 2177: потери по таблице 4 (A·L+B), не по ISO.

ФАЙЛ: фракционка ГОСТ_2177.xlsx
ФАКТЫ (инспекция 2026-10-01):
- Лист1!A1 = '='Ввод данных'!D3'  (= 760 мм рт.ст. — давление в мм)
- Лист1!E1 = '=A1*0.133322'        (кПа — использовалось ISO-формулой)
- Лист1!C28 = '=Вывод!E30'         (потери L, сейчас 0.5 %)
- Лист1!H28 = '=(F26/G26)+0.5'     (ISO: Lc = 0.5+(L-0.5)/(1+(101.3-p)/8))  -> ЗАМЕНИТЬ
- Лист1!H29 = '=C26+(C28-H28)'     (скорректированный отгон — пересчитается сам)
- Лист1!K28 = '=MROUND(ROUNDUP(H28,1),0.5)', K29 = '=MROUND(ROUNDUP(H29,1),0.5)'

ПРАВКА:
(а) Новый лист «Табл4» с таблицей 4 ГОСТ 2177-99: давление кПа, мм рт.ст., A, B (21 строка).
(б) Лист1!H28 = V_K = A·L + B, где A/B — линейная интерполяция по давлению A1 (мм рт.ст.)
    между строками таблицы 4 (560..760). Если A1 за пределами — экстраполяция крайними.
(в) H29/K29/K28 пересчитаются автоматически (формулы ссылаются на H28).
(г) НЕ трогать: номограмму, r/R, K ГСО, 0,84R, остальные колонки.

ПРОВЕРКА: P=760 (A=1,B=0) -> V_K = L.  L=2.5, P=740 (A=0.750,B=0.125) -> V_K=2.0.
"""
import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fix_common import norm_formula, backup_if_needed

GOST = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                    '..', 'фракционный состав', 'фракционка ГОСТ_2177.xlsx')
GOST = os.path.normpath(GOST)

# Таблица 4 ГОСТ 2177-99: (кПа, мм рт.ст., A, B)
TAB4 = [
    (74.7, 560, 0.231, 0.384),
    (76.0, 570, 0.240, 0.380),
    (77.3, 580, 0.250, 0.375),
    (78.7, 590, 0.261, 0.369),
    (80.0, 600, 0.273, 0.363),
    (81.3, 610, 0.286, 0.357),
    (82.6, 620, 0.300, 0.350),
    (84.0, 630, 0.316, 0.342),
    (85.3, 640, 0.333, 0.333),
    (86.6, 650, 0.353, 0.323),
    (88.0, 660, 0.375, 0.312),
    (89.3, 670, 0.400, 0.300),
    (90.6, 680, 0.428, 0.286),
    (92.0, 690, 0.461, 0.269),
    (93.3, 700, 0.500, 0.250),
    (94.6, 710, 0.545, 0.227),
    (96.0, 720, 0.600, 0.200),
    (97.3, 730, 0.667, 0.166),
    (98.6, 740, 0.750, 0.125),
    (100.0, 750, 0.857, 0.071),
    (101.3, 760, 1.000, 0.000),
]


def add_tab4_sheet(wb):
    """Добавить лист «Табл4» с таблицей 4 (идемпотентно: не пересоздавать с данными)."""
    if 'Табл4' in wb.sheetnames:
        ws = wb['Табл4']
        # если уже заполнен — пропустить
        if ws['A1'].value is not None:
            return False
    else:
        ws = wb.create_sheet('Табл4')
    ws['A1'] = 'Таблица 4 ГОСТ 2177-99 — константы для расчёта потерь V_K = A·L + B'
    ws['A2'] = 'Давление, 10³ Па'
    ws['B2'] = 'мм рт.ст.'
    ws['C2'] = 'A'
    ws['D2'] = 'B'
    for i, (kpa, mm, a, b) in enumerate(TAB4):
        r = 3 + i
        ws.cell(row=r, column=1, value=kpa)
        ws.cell(row=r, column=2, value=mm)
        ws.cell(row=r, column=3, value=a)
        ws.cell(row=r, column=4, value=b)
    return True


def formula_VK():
    """
    V_K = A·L + B с КУСОЧНОЙ линейной интерполяцией A/B по давлению A1 (мм рт.ст.).
    Табл4: B3..B23 = мм рт.ст., C3..C23 = A, D3..D23 = B (21 точка, 560..760).

    Идея (без VBA):
      idx = MATCH(A1, B3:B23, 1)          — индекс нижней границы (последняя B <= A1)
      если A1 <= B3  -> берём первую точку (C3)
      если A1 >= B23 -> берём последнюю (C23)
      иначе линейно между idx и idx+1.

    Реализация через IF + INDEX/MATCH (все версии Excel):
      A = IF(A1<=$B$3, $C$3,
           IF(A1>=$B$23, $C$23,
              INDEX($C$3:$C$23,MATCH(A1,$B$3:$B$23,1))
              + (INDEX($C$4:$C$24,MATCH(A1,$B$3:$B$23,1))
                 - INDEX($C$3:$C$23,MATCH(A1,$B$3:$B$23,1)))
                * (A1-INDEX($B$3:$B$23,MATCH(A1,$B$3:$B$23,1)))
                / (INDEX($B$4:$B$24,MATCH(A1,$B$3:$B$23,1))
                   - INDEX($B$3:$B$23,MATCH(A1,$B$3:$B$23,1)))))
    Аналогично B (колонка D).
    V_K = A*C28 + B
    """
    def interp(col):
        # col: 'C' (для A) или 'D' (для B)
        return (
            "IF(A1<=Табл4!$B$3,Табл4!$%s$3,"
            "IF(A1>=Табл4!$B$23,Табл4!$%s$23,"
            "INDEX(Табл4!$%s$3:$%s$23,MATCH(A1,Табл4!$B$3:$B$23,1))"
            "+(INDEX(Табл4!$%s$4:$%s$24,MATCH(A1,Табл4!$B$3:$B$23,1))"
            "-INDEX(Табл4!$%s$3:$%s$23,MATCH(A1,Табл4!$B$3:$B$23,1)))"
            "*(A1-INDEX(Табл4!$B$3:$B$23,MATCH(A1,Табл4!$B$3:$B$23,1)))"
            "/(INDEX(Табл4!$B$4:$B$24,MATCH(A1,Табл4!$B$3:$B$23,1))"
            "-INDEX(Табл4!$B$3:$B$23,MATCH(A1,Табл4!$B$3:$B$23,1)))))"
        ) % (col, col, col, col, col, col, col, col)
    return "=%s*C28+%s" % (interp('C'), interp('D'))


def fix_file(path):
    print('\n==== ГОСТ_2177 (ФР-13): %s ====' % os.path.basename(path))
    backup_if_needed()  # страховка перед модификацией (идемпотентно, не перезаписывает)
    wb = openpyxl.load_workbook(path, data_only=False)
    changes = []

    # (а) лист «Табл4»
    added = add_tab4_sheet(wb)
    print('  Табл4: %s' % ('добавлена/заполнена' if added else 'уже существует'))

    # (б) H28 = V_K
    ws = wb['Лист1']
    old_h28 = ws['H28'].value
    new_h28 = formula_VK()
    if norm_formula(old_h28) != norm_formula(new_h28):
        ws['H28'] = new_h28
        changes.append(('ФР-13', 'Лист1!H28', old_h28, new_h28))
        print('  H28: V_K = A*L+B (FORECAST-интерполяция)')
    else:
        print('  H28 уже корректна (пропуск)')

    # (в) контроль связанных (пересчитаются автоматически)
    print('  H29 (отгон) =', repr(ws['H29'].value), '— ссылается на H28, пересчитается')
    print('  K28 =', repr(ws['K28'].value), '| K29 =', repr(ws['K29'].value))

    wb.save(path)
    print('  SAVED:', path)
    wb.close()
    return changes


def main():
    changes = fix_file(GOST)
    print('\nИТОГО изменений:', len(changes))
    for c in changes:
        print('  %s | %s | было=%r | стало=%r' % c[:4])
    return 0


if __name__ == '__main__':
    sys.exit(main())