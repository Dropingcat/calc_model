# -*- coding: utf-8 -*-
"""
Автотесты НД-аудита книг фракционного состава из репозитория
(new_versions_v4 — v3.11, new_versions_v4.1 — v3.12 нормативно аудированные).

Контракт ТЗ (аудит «НД -> процедура -> модель -> расчёт -> критерий -> отчёт»):
  ASTM D86-23:
    N02  Lc = 0.5 + (L-0.5)/(1+(760-P)/60) применяется для ВСЕГО диапазона P,
         включая P > 760 (в v3.11 при P>760 поправка отключалась IF(P>760,L)).
    N03  Slope в точке 95% использует соседние равноотстоящие точки 90 и 100%.
    N08  Контроль группы 1-4 присутствует; manual-режим помечен.
    Аудит-слой: листы «Аудит_НД» и «Техдолг» с ID D86-Nxx / TD-D86-xxx.
  ГОСТ 2177-99 с Изм. №1 (5.6.3):
    S90 = (F[95]-F[85])/(95-85); S95 = (F[100]-F[90])/(100-90)
         (в v3.11 было 85-90 и 90-95 — исправлено в v3.12).
    Таблица 4: диапазон 560..760 мм рт.ст. контролируется явно,
         выход за диапазон не превращается молча в ближайшую строку.
    Формула поправки kPa-форма 0.0009*(101.3-p_kPa)*(273+t) — эквивалент
         нормативной 0.0000009*(101300-Pb)*(273+t) — НЕ должна быть «исправлена».
    Аудит-слой: листы «Аудит_НД» и «Техдолг» с ID.

Запуск: python test_books.py   (exit 0 = все тесты прошли)
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl

from fix_common import V4, V41

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


def formulas_on_sheet(wb, sheet):
    out = {}
    ws = wb[sheet]
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                out[c.coordinate] = ''.join(c.value.split()).upper()
    return out


def norm(s):
    return ''.join(str(s).split()).upper() if s else ''


# ------------------------------------------------------------------ ASTM D86
def test_d86_lc_full_range():
    print('\n--- D86-N02: Lc для всего диапазона давлений (без отключения при P>760) ---')
    # v3.11: дефект подтверждён (IF(P>760 -> наблюдаемая потеря))
    wb = load(V4['D86_MANUAL'])
    f = formulas_on_sheet(wb, 'Расчёт')
    g2 = f.get('G2', '')
    check('v3.11: дефект на месте (IF(B1>760,...) в G2)', 'B$1>760' in g2 or 'B1>760' in g2, g2[:120])
    wb.close()
    # v3.12: единая формула для всех P
    wb = load(V41['D86'])
    f = formulas_on_sheet(wb, 'Расчёт')
    bad_guard = [k for k, v in f.items() if re.match(r'^G\d+$', k) and '>760' in v]
    check('v3.12: ни одна G-формула не содержит guard P>760', not bad_guard, str(bad_guard[:5]))
    ok_lc = all('(760-' in v and '/60)' in v and '0.5+' in v for k, v in f.items()
                if re.match(r'^G\d+$', k))
    check('v3.12: все G-формулы = 0.5+(L-0.5)/(1+(760-P)/60)', ok_lc,
          str([f[k][:80] for k in sorted(f)[:2]]))
    wb.close()


def test_d86_slope_95():
    print('\n--- D86-N03: slope 95% через 90 и 100% ---')
    # найти столбец крутизны: формула вида ABS(F{i+?}-F{i-?})/(A... ) вокруг строк 20-24
    def slope_cells(wb, sheet='Расчёт'):
        ws = wb[sheet]
        found = {}
        for col in range(1, 30):
            letter = openpyxl.utils.get_column_letter(col)
            v = ws['%s23' % letter].value
            if isinstance(v, str) and v.startswith('=') and 'ABS(' in v.upper() and '/' in v \
                    and ('F2' in v or 'E2' in v):
                found[letter] = {r: ws['%s%d' % (letter, r)].value for r in range(21, 25)}
        return found

    for tag, path in (('v3.11', V4['D86_MANUAL']), ('v3.12', V41['D86'])):
        wb = load(path)
        sc = slope_cells(wb)
        if not sc:
            check('D86 slope (%s): столбец крутизны найден' % tag, False, 'не найден')
            wb.close()
            continue
        letter = sorted(sc)[0]
        cells = sc[letter]
        v95 = norm(cells.get('T' and list(cells)[3] and 23)) if False else norm(cells.get(23))
        # slope в строке 23 (95%) должен ссылаться на F22 (90%) и F24 (100%)
        ok = ('F22' in v95 and 'F24' in v95)
        check('D86 slope (%s): 95%% = (F100-F90)/(100-90) [%s23]' % (tag, letter), ok, v95[:120])
        wb.close()


def test_d86_audit_layer():
    print('\n--- D86: аудит-слой (листы Аудит_НД / Техдолг, редакции НД) ---')
    wb = load(V41['D86'])
    check('лист Аудит_НД присутствует', 'Аудит_НД' in wb.sheetnames, str(wb.sheetnames))
    check('лист Техдолг присутствует', 'Техдолг' in wb.sheetnames)
    if 'Аудит_НД' in wb.sheetnames:
        txt = '\n'.join(str(c.value) for row in wb['Аудит_НД'].iter_rows() for c in row if c.value)
        check('Аудит_НД: указана редакция D86-23', 'D86-23' in txt)
        check('Аудит_НД: есть пункт про полный диапазон Lc (N02)', 'D86-N02' in txt)
        check('Аудит_НД: есть пункт про slope 90/100 (N03)', 'D86-N03' in txt)
    if 'Техдолг' in wb.sheetnames:
        txt = '\n'.join(str(c.value) for row in wb['Техдолг'].iter_rows() for c in row if c.value)
        for tdid in ['TD-D86-00%d' % i for i in range(1, 9)]:
            check('Техдолг: %s зарегистрирован' % tdid, tdid in txt)
    # контроль группы 1-4
    allf = {}
    for sh in wb.sheetnames:
        allf.update(formulas_on_sheet(wb, sh))
    grp = [v for v in allf.values() if 'ГРУПП' in v or re.search(r'\b(1|2|3|4)\b.*GROUP', v)]
    check('контроль группы присутствует в формулах', bool(grp), '')
    wb.close()


# ------------------------------------------------------------------ ГОСТ 2177
def test_g2177_slope_563():
    print('\n--- ГОСТ 2177 5.6.3: S90 через 85-95, S95 через 90-100 ---')
    # строки уровней: A18=85, A19=90(?), ... уточняем по значениям столбца A листа Расчёт
    for tag, path, expect_ok in (('v3.11', V4['G2177'], False), ('v3.12', V41['G2177'], True)):
        wb = load(path)
        ws = wb['Расчёт']
        # находим строки со значениями % 85,90,95,100 в столбце A
        pct_row = {}
        for r in range(2, 40):
            a = ws['A%d' % r].value
            if isinstance(a, (int, float)) and int(a) in (85, 90, 95, 100):
                pct_row.setdefault(int(a), r)
        if len(pct_row) < 4:
            # значения могут быть формулами — ищем константы 85/90/95/100 где угодно в A
            for r in range(2, 40):
                v = ws['A%d' % r].value
                if v is None:
                    continue
            check('ГОСТ slope (%s): строки 85/90/95/100 определены' % tag, False, str(pct_row))
            wb.close()
            continue
        r85, r90, r95, r100 = (pct_row[85], pct_row[90], pct_row[95], pct_row[100])
        s90 = norm(ws['S%d' % r90].value)
        s95 = norm(ws['S%d' % r95].value)
        exp90 = ('F%d' % r95) in s90 and ('F%d' % r85) in s90
        exp95 = ('F%d' % r100) in s95 and ('F%d' % r90) in s95
        check('ГОСТ slope (%s): S90 = (F95-F85)/(95-85)' % tag, exp90 == expect_ok, s90[:110])
        check('ГОСТ slope (%s): S95 = (F100-F90)/(100-90)' % tag, exp95 == expect_ok, s95[:110])
        wb.close()


def test_g2177_tab4_control():
    print('\n--- ГОСТ 2177: таблица 4, контроль диапазона 560..760 ---')
    wb = load(V41['G2177'])
    ws = wb['Табл4']
    # 21 строка узлов 560..760 (мм рт.ст.)
    mm_vals = [ws.cell(row=r, column=2).value for r in range(3, 24)]
    ok_nodes = len(mm_vals) == 21 and mm_vals[0] == 560 and mm_vals[-1] == 760
    check('Табл4: 21 узел 560..760', ok_nodes, str(mm_vals[:3]) + '...' + str(mm_vals[-2:]))
    # явный контроль выхода за диапазон: ищем формулы с <560 / >760 guard
    guards = []
    for sh in wb.sheetnames:
        for coord, fv in formulas_on_sheet(wb, sh).items():
            if ('560' in fv or '760' in fv) and ('<' in fv or '>' in fv) and 'MATCH' in fv.upper():
                guards.append((sh, coord))
    check('v3.12: явный guard диапазона таблицы 4', len(guards) > 0, str(guards[:3]))
    wb.close()


def test_g2177_pressure_formula_preserved():
    print('\n--- ГОСТ 2177: kPa-форма поправки сохранена (не «исправлять»!) ---')
    # 0.0009*(101.3-p*0.133322)*(273+t) == нормативной 0.0000009*(101300-Pb)*(273+t)
    wb = load(V41['G2177'])
    f = formulas_on_sheet(wb, 'Расчёт')
    b30 = f.get('B30', '')
    ok = '0.0009*' in b30 and '101.3-' in b30 and '0.133322' in b30
    check('сохранена корректная kPa-форма 0.0009*(101.3-P*0.133322)', ok, b30[:120])
    # округление до 0.5 после корректировки (MROUND/ROUND(...,0.5 via ROUND(x*2)/2 или MROUND)
    wb.close()


def test_g2177_audit_layer():
    print('\n--- ГОСТ 2177: аудит-слой ---')
    wb = load(V41['G2177'])
    check('лист Аудит_НД присутствует', 'Аудит_НД' in wb.sheetnames, str(wb.sheetnames))
    check('лист Техдолг присутствует', 'Техдолг' in wb.sheetnames)
    if 'Техдолг' in wb.sheetnames:
        txt = '\n'.join(str(c.value) for row in wb['Техдолг'].iter_rows() for c in row if c.value)
        for tdid in ['TD-2177-00%d' % i for i in range(1, 9)]:
            check('Техдолг: %s зарегистрирован' % tdid, tdid in txt)
    if 'Аудит_НД' in wb.sheetnames:
        txt = '\n'.join(str(c.value) for row in wb['Аудит_НД'].iter_rows() for c in row if c.value)
        check('Аудит_НД: упомянута редакция ГОСТ 2177-99 с Изм. №1',
              ('2177-99' in txt and ('Изм' in txt or 'измен' in txt.lower())), txt[:80])
    wb.close()


# ------------------------------------------------------------------ целостность
def test_books_openable():
    print('\n--- целостность: все книги открываются, структура листов на месте ---')
    REQUIRED_D86 = ['О книге', 'Ввод', 'Расчёт', 'Итог', 'График']
    REQUIRED_G = ['О книге', 'Ввод', 'Табл4', 'Номограмма', 'Расчёт', 'Итог', 'График']
    cases = [(V4['ISO3405'], None), (V4['D86_MANUAL'], REQUIRED_D86), (V4['D86_AUTO'], REQUIRED_D86),
             (V4['D1160'], REQUIRED_D86), (V4['G2177'], REQUIRED_G), (V4['G2177_APPA'], REQUIRED_G),
             (V41['D86'], REQUIRED_D86 + ['Аудит_НД', 'Техдолг']),
             (V41['G2177'], REQUIRED_G + ['Аудит_НД', 'Техдолг'])]
    for p, req in cases:
        name = os.path.basename(p)
        try:
            wb = load(p)
            ok = (req is None) or all(s in wb.sheetnames for s in req)
            check('книга открывается, листы на месте: %s' % name[:50], ok, str(wb.sheetnames))
            wb.close()
        except Exception as e:
            check('книга открывается: %s' % name[:50], False, repr(e))


def main():
    print('=' * 70)
    print('TEST_BOOKS: НД-аудит книг new_versions_v4 / v4.1')
    print('=' * 70)
    test_books_openable()
    test_d86_lc_full_range()
    test_d86_slope_95()
    test_d86_audit_layer()
    test_g2177_slope_563()
    test_g2177_tab4_control()
    test_g2177_pressure_formula_preserved()
    test_g2177_audit_layer()
    print('\n' + '=' * 70)
    print('ИТОГО: %d PASS, %d FAIL' % (len(PASS), len(FAIL)))
    if FAIL:
        print('ПРОВАЛЕНЫ:', FAIL)
        return 1
    print('ВСЕ ТЕСТЫ ПРОЙДЕНЫ.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
