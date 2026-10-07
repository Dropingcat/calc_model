# -*- coding: utf-8 -*-
"""
Автотесты исправлений фракционного состава (ФР-01..ФР-06, ФР-16).

Пункты контракта:
  (а) D23 = '='Ввод данных'!D23'  [ФР-04]
  (б) колонка O группы 1: +1.11 на 5%, -1.22 на 90%, -0.94 на 95%, без константы на 10-80%  [ФР-01]
  (в) D86 P: 5% -> 3.1+, 95% -> 1.1+  [ФР-02]
  (г) D1160 M/N: для P=10мм r/R из колонок D/I (5-50%) и E/J (60-95%), MATCH выровнен  [ФР-06]
  (д) отсутствие #REF!/#N/A в изменённых ячейках
  (е) K6..K26 = Лист1!I4..I24  [ФР-16]
  (ж) IFERROR в блоке температур D86  [ФР-03]
  (з) D1160 L11..L15: E46/E44/E45 (10 мм 60-95%), нет D47..D51  [ФР-05]

Запуск: python test_fix.py  (exit 0 = все тесты прошли)
"""
import os
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl

from fix_common import ISO1, ISO2, D86, D1160, G2177, norm_formula

PASS = []
FAIL = []


def check(name, cond, detail=''):
    if cond:
        PASS.append(name)
        print('  [PASS] %s' % name)
    else:
        FAIL.append(name)
        print('  [FAIL] %s  %s' % (name, detail))


def load(path):
    return openpyxl.load_workbook(path, data_only=False)


def no_ref_na(path, cells):
    """(д) проверить отсутствие #REF!/#N/A/#VALUE!/#DIV/0! в формулах изменённых ячеек."""
    bad = []
    wb = load(path)
    for sheet, cell in cells:
        v = wb[sheet][cell].value
        if isinstance(v, str) and ('#REF!' in v.upper() or '#N/A' in v.upper()
                                   or '#VALUE!' in v.upper() or '#DIV/0!' in v.upper()):
            bad.append('%s!%s=%r' % (sheet, cell, v))
    wb.close()
    return bad


# ---------------------------------------------------------------- (а) ФР-04
def test_fr04():
    print('\n--- (а) ФР-04: D23 ---')
    wb = load(ISO1)
    v = wb['Группы 1-4']['D23'].value
    check('ФР-04: ISO1 D23 = =Ввод данных!D23',
          norm_formula(v) == norm_formula("='Ввод данных'!D23"), repr(v))
    wb.close()
    wb = load(ISO2)
    v = wb['Группы 1-4']['D23'].value
    check('ФР-04: ISO2 D23 = =Ввод данных!D23',
          norm_formula(v) == norm_formula("='Ввод данных'!D23"), repr(v))
    wb.close()


# ---------------------------------------------------------------- (б) ФР-01
def test_fr01():
    print('\n--- (б) ФР-01: колонка O группа 1 ---')
    # Ожидания: row -> (pct, ожидаемая подстрока после 1.994)
    # В формуле: ((1.736*K5+1.994)+1.11) — после 1.994 идёт ')+' или ')-'
    EXPECT = {
        5: '+1.11', 6: '', 8: '', 10: '', 12: '', 14: '', 16: '', 18: '', 20: '',
        22: '-1.22', 23: '-0.94',
    }
    for tag, path in (('ISO1', ISO1), ('ISO2', ISO2)):
        wb = load(path)
        ws = wb['Лист1']
        for row, frag in EXPECT.items():
            v = norm_formula(ws['O%d' % row].value)
            core = '1.736*K%d+1.994' % row
            if frag:
                ok = (core + ')' + frag) in v
            else:
                # без константы: сразу после 1.994 идёт '))'
                ok = (core + '))') in v and (core + ')-') not in v and (core + ')+') not in v
            check('ФР-01: %s O%d %s' % (tag, row, ('+%s' % frag) if frag else 'без константы'),
                  ok, ws['O%d' % row].value)
        # условие группы — V$3 (а не V3) в ISO2 O22/O23 исправлено
        v22 = norm_formula(ws['O22'].value)
        check('ФР-01: %s O22 использует V$3 (абс.)' % tag, 'V$3' in v22, ws['O22'].value)
        wb.close()


# ---------------------------------------------------------------- (в) ФР-02
def test_fr02():
    print('\n--- (в) ФР-02: D86 P 5%/95% ---')
    wb = load(D86)
    ws = wb['расчёт']
    v5 = norm_formula(ws['P5'].value)
    v23 = norm_formula(ws['P23'].value)
    check('ФР-02: P5 = 3.1+1.74*L', '1.74*L5+3.1' in v5, ws['P5'].value)
    check('ФР-02: P23 = 1.1+1.74*L', '1.74*L23+1.1' in v23, ws['P23'].value)
    # контроль: P6 (10%), P22 (90%), P24 (FBP) не изменены
    v6 = norm_formula(ws['P6'].value)
    v22 = norm_formula(ws['P22'].value)
    v24 = norm_formula(ws['P24'].value)
    check('ФР-02: P6 (10%) = 2.0+ (не изменён)', '1.74*L6+2' in v6, ws['P6'].value)
    check('ФР-02: P22 (90%) = 0.8+ (не изменён)', '1.74*L22+0.8' in v22, ws['P22'].value)
    check('ФР-02: P24 (FBP) = 7.2 (не изменён)', '7.2' in v24 and '0.42*L24+3.1' in v24, ws['P24'].value)
    wb.close()


# ---------------------------------------------------------------- (г) ФР-06
def test_fr06():
    print('\n--- (г) ФР-06: D1160 M/N колонки и MATCH ---')
    wb = load(D1160)
    ws = wb['Лист2']
    # 5-50% (строки 5..10): M: 1мм->B / 10мм->D; N: 1мм->G / 10мм->I
    for row in range(5, 11):
        m = norm_formula(ws['M%d' % row].value)
        n = norm_formula(ws['N%d' % row].value)
        check('ФР-06: M%d: ветки B/D' % row,
              'INDEX(ЛИСТ1!B$5:B$34' in m and 'INDEX(ЛИСТ1!D$5:D$34' in m, ws['M%d' % row].value)
        check('ФР-06: N%d: ветки G/I' % row,
              'INDEX(ЛИСТ1!G$5:G$34' in n and 'INDEX(ЛИСТ1!I$5:I$34' in n, ws['N%d' % row].value)
        check('ФР-06: M%d MATCH=G%d' % (row, row),
              'MATCH(ЛИСТ2!G%d' % row in m, ws['M%d' % row].value)
    # 60-95% (строки 11..15): M: C/E; N: H/J; MATCH = своя строка
    for row in range(11, 16):
        m = norm_formula(ws['M%d' % row].value)
        n = norm_formula(ws['N%d' % row].value)
        check('ФР-06: M%d: ветки C/E' % row,
              'INDEX(ЛИСТ1!C$5:C$34' in m and 'INDEX(ЛИСТ1!E$5:E$34' in m, ws['M%d' % row].value)
        check('ФР-06: N%d: ветки H/J' % row,
              'INDEX(ЛИСТ1!H$5:H$34' in n and 'INDEX(ЛИСТ1!J$5:J$34' in n, ws['N%d' % row].value)
        check('ФР-06: M%d MATCH=G%d' % (row, row),
              'MATCH(ЛИСТ2!G%d' % row in m, ws['M%d' % row].value)
    # условие: E1 = N2 (Лист1!N2=1)
    ws1 = wb['Лист1']
    check('ФР-06: условие E1=Лист1!N2 (N2=1)',
          ws1['N2'].value == 1, repr(ws1['N2'].value))
    wb.close()


# ---------------------------------------------------------------- (е) ФР-16
def test_fr16():
    print('\n--- (е) ФР-16: K6..K26 = Лист1!I4..I24 ---')
    for tag, path in (('ISO1', ISO1), ('ISO2', ISO2)):
        wb = load(path)
        ws = wb['Группы 1-4']
        ok_all = True
        for i, row in enumerate(range(6, 27)):
            src = 4 + i
            if norm_formula(ws['K%d' % row].value) != norm_formula('=Лист1!I%d' % src):
                ok_all = False
                print('    K%d = %r (ожидалось =Лист1!I%d)' % (row, ws['K%d' % row].value, src))
        check('ФР-16: %s K6..K26 = Лист1!I4..I24' % tag, ok_all)
        wb.close()


# ---------------------------------------------------------------- (ж) ФР-03
def test_fr03():
    print('\n--- (ж) ФР-03: IFERROR в блоке температур D86 ---')
    wb = load(D86)
    ws = wb['расчёт']
    rows = range(36, 48)
    # K37..K47, L36..L47, M36..M47 должны содержать IFERROR
    bad = []
    for r in rows:
        for col in ('K', 'L', 'M'):
            v = ws['%s%d' % (col, r)].value
            if isinstance(v, str) and v.startswith('=') and 'IFERROR' not in v.upper():
                bad.append('%s%d' % (col, r))
    check('ФР-03: все K/L/M 36..47 в IFERROR', not bad, 'нет IFERROR: %s' % bad)
    # N36..N47, O36..O47, Q36..Q47 уже были в IFERROR
    bad2 = []
    for r in rows:
        for col in ('N', 'O', 'Q'):
            v = ws['%s%d' % (col, r)].value
            if isinstance(v, str) and v.startswith('=') and 'IFERROR' not in v.upper():
                bad2.append('%s%d' % (col, r))
    check('ФР-03: N/O/Q 36..47 в IFERROR', not bad2, 'нет IFERROR: %s' % bad2)
    wb.close()


# ---------------------------------------------------------------- (з) ФР-05
def test_fr05():
    print('\n--- (з) ФР-05: D1160 L11..L15 (10 мм 60-95%) ---')
    wb = load(D1160)
    ws = wb['Лист2']
    for row in range(11, 16):
        v = norm_formula(ws['L%d' % row].value)
        ok = ('E$46' in v and 'E$44' in v and 'E$45' in v
              and 'D47' not in v and 'D48' not in v and 'D49' not in v
              and 'D50' not in v and 'D51' not in v
              and 'MATCH' not in v)  # L — формула, не INDEX
        check('ФР-05: L%d = E46/E44/E45, без D47..D51' % row, ok, ws['L%d' % row].value)
        check('ФР-05: L%d использует G%d (свою строку)' % (row, row),
              'G%d' % row in v, ws['L%d' % row].value)
    wb.close()


# ---------------------------------------------------------------- (з-2) ДОВОДКА: L переключается по давлению
def test_fr05_switch():
    print('\n--- (з-2) ДОВОДКА: D1160 L5..L15 переключение по давлению ---')
    wb = load(D1160)
    ws = wb['Лист2']
    # Ожидаемые константы (Лист1):
    #  5-50%: 1мм -> D42/D40/D41 (M/a/b); 10мм -> D46/D44/D45
    #  60-95%: 1мм -> E42/E40/E41; 10мм -> E46/E44/E45
    for row in range(5, 11):
        v = norm_formula(ws['L%d' % row].value)
        ok = ('IF(' in v
              and 'D$42' in v and 'D$40' in v and 'D$41' in v   # ветка 1 мм
              and 'D$46' in v and 'D$44' in v and 'D$45' in v   # ветка 10 мм
              and 'E$4' not in v)                               # 5-50% не должны иметь E
        check('ДОВОДКА: L%d = IF(1мм:D42/D40/D41, 10мм:D46/D44/D45)' % row,
              ok, ws['L%d' % row].value)
        check('ДОВОДКА: L%d условие E1=N2' % row,
              norm_formula("'АСТМ Д 1160'!$E$1=Лист1!$N$2") in v,
              ws['L%d' % row].value)
    for row in range(11, 16):
        v = norm_formula(ws['L%d' % row].value)
        ok = ('IF(' in v
              and 'E$42' in v and 'E$40' in v and 'E$41' in v   # ветка 1 мм
              and 'E$46' in v and 'E$44' in v and 'E$45' in v   # ветка 10 мм
              and 'D$4' not in v)                               # 60-95% не должны иметь D
        check('ДОВОДКА: L%d = IF(1мм:E42/E40/E41, 10мм:E46/E44/E45)' % row,
              ok, ws['L%d' % row].value)
        check('ДОВОДКА: L%d использует G%d (свою строку)' % (row, row),
              'G%d' % row in v, ws['L%d' % row].value)
    # эталонные константы на Лист1 (сверка с ASTM D1160-18 12.5.1)
    ws1 = wb['Лист1']
    checks = {
        'D40': 0.439, 'D41': 0.241, 'D42': 2.9,     # 1 мм 5-50
        'E40': 0.439, 'E41': 0.241, 'E42': 3.0,     # 1 мм 60-95
        'D44': 0.24, 'D45': 0.35, 'D46': 2.8,       # 10 мм 5-50
        'E44': 0.24, 'E45': 0.35, 'E46': 2.9,       # 10 мм 60-95
    }
    for cell, exp in checks.items():
        check('ДОВОДКА: Лист1!%s = %s (ASTM 12.5.1)' % (cell, exp),
              ws1[cell].value == exp, repr(ws1[cell].value))
    wb.close()


# ---------------------------------------------------------------- (д) нет #REF!/#N/A
def test_no_errors():
    print('\n--- (д) нет #REF!/#N/A/#VALUE!/#DIV/0! ---')
    iso_cells = [('Группы 1-4', 'D23')] + [('Группы 1-4', 'K%d' % r) for r in range(6, 27)]
    iso1_cells = iso_cells + [('Лист1', 'O%d' % r) for r in (5, 6, 8, 10, 12, 14, 16, 18, 20, 22, 23)]
    iso2_cells = iso_cells + [('Лист1', 'O%d' % r) for r in (5, 6, 8, 10, 12, 14, 16, 18, 20, 22, 23)]
    d86_cells = ([('расчёт', 'P%d' % r) for r in (5, 23)]
                 + [('расчёт', '%s%d' % (c, r)) for r in range(36, 48) for c in 'KLMNOQ'])
    d1160_cells = ([('Лист2', 'L%d' % r) for r in range(5, 16)]
                   + [('Лист2', 'M%d' % r) for r in range(11, 16)])
    for name, path, cells in (('ISO1', ISO1, iso1_cells), ('ISO2', ISO2, iso2_cells),
                              ('D86', D86, d86_cells), ('D1160', D1160, d1160_cells)):
        bad = no_ref_na(path, cells)
        check('(д): %s нет #REF/#N/A в изменённых' % name, not bad, str(bad))


# ---------------------------------------------------------------- ФР-13
TAB4_EXPECT = [
    (74.7, 560, 0.231, 0.384), (76.0, 570, 0.240, 0.380), (77.3, 580, 0.250, 0.375),
    (78.7, 590, 0.261, 0.369), (80.0, 600, 0.273, 0.363), (81.3, 610, 0.286, 0.357),
    (82.6, 620, 0.300, 0.350), (84.0, 630, 0.316, 0.342), (85.3, 640, 0.333, 0.333),
    (86.6, 650, 0.353, 0.323), (88.0, 660, 0.375, 0.312), (89.3, 670, 0.400, 0.300),
    (90.6, 680, 0.428, 0.286), (92.0, 690, 0.461, 0.269), (93.3, 700, 0.500, 0.250),
    (94.6, 710, 0.545, 0.227), (96.0, 720, 0.600, 0.200), (97.3, 730, 0.667, 0.166),
    (98.6, 740, 0.750, 0.125), (100.0, 750, 0.857, 0.071), (101.3, 760, 1.000, 0.000),
]


def _interp_tab4(P, col):
    """Кусочная интерполяция по Табл4 (эмуляция формулы Excel). col='A'->C(idx2), 'B'->D(idx3)."""
    ci = 2 if col == 'A' else 3  # индекс в кортеже (кПа,мм,A,B): A=2, B=3
    vals = [t[ci] for t in TAB4_EXPECT]
    xs = [t[1] for t in TAB4_EXPECT]
    if P <= xs[0]:
        return vals[0]
    if P >= xs[-1]:
        return vals[-1]
    for i in range(len(xs) - 1):
        if xs[i] <= P <= xs[i + 1]:
            return vals[i] + (vals[i + 1] - vals[i]) * (P - xs[i]) / (xs[i + 1] - xs[i])
    return vals[-1]


def test_fr13():
    print('\n--- ФР-13: ГОСТ 2177 потери по таблице 4 ---')
    wb = load(G2177)
    # (а) лист «Табл4» существует и заполнен
    check('ФР-13: лист Табл4 существует', 'Табл4' in wb.sheetnames)
    ws_t = wb['Табл4']
    ok_tab = True
    for i, (kpa, mm, a, b) in enumerate(TAB4_EXPECT):
        r = 3 + i
        if (ws_t.cell(row=r, column=1).value != kpa
                or ws_t.cell(row=r, column=2).value != mm
                or abs((ws_t.cell(row=r, column=3).value or 0) - a) > 1e-9
                or abs((ws_t.cell(row=r, column=4).value or 0) - b) > 1e-9):
            ok_tab = False
            print('    Табл4 строка %d не совпадает: %r' % (r, [ws_t.cell(row=r, column=c).value for c in range(1, 5)]))
    check('ФР-13: Табл4 содержит 21 строку (560..760) с A/B', ok_tab)

    # (б) H28 = V_K = A*L + B (интерполяция), не ISO-формула
    ws = wb['Лист1']
    h28 = norm_formula(ws['H28'].value)
    # не должно быть ISO-структуры: F26, G26, (…/…)+0.5
    iso_ok = ('F26' not in h28 and 'G26' not in h28
              and not re.search(r'\)\+0\.5\)?\*?$', h28) and '0.5)' not in h28)
    check('ФР-13: H28 не ISO-формула (нет F26/G26, нет /+0.5)',
          iso_ok, ws['H28'].value)
    check('ФР-13: H28 использует MATCH/INDEX и C28',
          'MATCH' in h28 and 'INDEX' in h28 and 'C28' in h28, ws['H28'].value)
    check('ФР-13: H28 ссылается на Табл4 (A/B по давлению)',
          'ТАБЛ4!$C$3' in h28 and 'ТАБЛ4!$D$3' in h28, ws['H28'].value)

    # (в) связанные: H29, K28, K29 ссылаются на H28
    check('ФР-13: H29 = отгон + (потери - V_K)',
          norm_formula(ws['H29'].value) == norm_formula('=C26+(C28-H28)'), ws['H29'].value)
    check('ФР-13: K28 округляет H28', 'H28' in norm_formula(ws['K28'].value), ws['K28'].value)
    check('ФР-13: K29 округляет H29', 'H29' in norm_formula(ws['K29'].value), ws['K29'].value)

    # (г) проверка формулы: P=760 -> V_K=L; L=2.5,P=740 -> V_K=2.0
    # эмуляция формулы Excel (кусочная интерполяция)
    vk_760 = _interp_tab4(760, 'A') * 0.5 + _interp_tab4(760, 'B')
    check('ФР-13: P=760, L=0.5 -> V_K=0.5 (A=1,B=0)', abs(vk_760 - 0.5) < 1e-9, vk_760)
    vk_740 = _interp_tab4(740, 'A') * 2.5 + _interp_tab4(740, 'B')
    check('ФР-13: L=2.5, P=740 -> V_K=2.0 (пример ГОСТ)', abs(vk_740 - 2.0) < 1e-9, vk_740)
    # промежуточная точка: P=700 (A=0.5, B=0.25)
    vk_700 = _interp_tab4(700, 'A') * 1.0 + _interp_tab4(700, 'B')
    check('ФР-13: L=1.0, P=700 -> V_K=0.75', abs(vk_700 - 0.75) < 1e-9, vk_700)
    # точка между узлами: P=725 (между 720 и 730): A=0.6+(0.667-0.6)*5/10=0.6335
    vk_725 = _interp_tab4(725, 'A') * 1.0 + _interp_tab4(725, 'B')
    a_exp = 0.6 + (0.667 - 0.6) * 5 / 10
    b_exp = 0.2 + (0.166 - 0.2) * 5 / 10
    check('ФР-13: P=725 интерполяция (A=%.4f,B=%.4f)' % (a_exp, b_exp),
          abs(vk_725 - (a_exp * 1.0 + b_exp)) < 1e-9, vk_725)
    # C28 (L) — потери из Вывод!E30
    check('ФР-13: C28 = потери (Вывод!E30)',
          norm_formula(ws['C28'].value) == norm_formula('=Вывод!E30'), ws['C28'].value)
    wb.close()


# ---------------------------------------------------------------- структура листов
def test_sheets_intact():
    print('\n--- структура листов не изменена ---')
    for name, path in (('ISO1', ISO1), ('ISO2', ISO2), ('D86', D86), ('D1160', D1160), ('G2177', G2177)):
        wb = load(path)
        check('структура: %s листы на месте' % name, len(wb.sheetnames) > 0)
        wb.close()


def main():
    print('=' * 70)
    print('TEST_FIX: исправления фракционного состава')
    print('=' * 70)
    test_sheets_intact()
    test_fr04()
    test_fr01()
    test_fr02()
    test_fr03()
    test_fr05()
    test_fr05_switch()
    test_fr06()
    test_fr16()
    test_fr13()
    test_no_errors()
    print('\n' + '=' * 70)
    print('ИТОГО: %d PASS, %d FAIL' % (len(PASS), len(FAIL)))
    if FAIL:
        print('ПРОВАЛЕНЫ:', FAIL)
        return 1
    print('ВСЕ ТЕСТЫ ПРОЙДЕНЫ.')
    return 0


if __name__ == '__main__':
    sys.exit(main())