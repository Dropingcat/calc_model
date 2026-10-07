# -*- coding: utf-8 -*-
"""Проверка идемпотентности: повторный прогон fix-скриптов не должен менять формулы."""
import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fix_common import ISO1, ISO2, D86, D1160

# 1. Снимок формул до
def snapshot(path, sheets):
    wb = openpyxl.load_workbook(path, data_only=False)
    out = {}
    for sn in sheets:
        ws = wb[sn]
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    out['%s!%s' % (sn, c.coordinate)] = c.value
    wb.close()
    return out

s1 = {p: snapshot(p, ['Группы 1-4', 'Лист1']) for p in (ISO1, ISO2)}
s2 = {p: snapshot(p, ['расчёт']) for p in (D86,)}
s3 = {p: snapshot(p, ['Лист2']) for p in (D1160,)}

# 2. Повторный прогон
import fix_iso3405, fix_d86, fix_d1160
print('--- повторный прогон fix-скриптов (ожидается: ничего не менять) ---')
fix_iso3405.main()
fix_d86.main()
fix_d1160.main()

# 3. Снимок после
t1 = {p: snapshot(p, ['Группы 1-4', 'Лист1']) for p in (ISO1, ISO2)}
t2 = {p: snapshot(p, ['расчёт']) for p in (D86,)}
t3 = {p: snapshot(p, ['Лист2']) for p in (D1160,)}

# 4. Сравнение
diffs = 0
for p in (ISO1, ISO2):
    for k in set(s1[p]) | set(t1[p]):
        if s1[p].get(k) != t1[p].get(k):
            print('ИЗМЕНИЛОСЬ при повторе:', p, k)
            diffs += 1
for p in (D86,):
    for k in set(s2[p]) | set(t2[p]):
        if s2[p].get(k) != t2[p].get(k):
            print('ИЗМЕНИЛОСЬ при повторе:', p, k)
            diffs += 1
for p in (D1160,):
    for k in set(s3[p]) | set(t3[p]):
        if s3[p].get(k) != t3[p].get(k):
            print('ИЗМЕНИЛОСЬ при повторе:', p, k)
            diffs += 1

print('\nИдемпотентность: %s различий' % diffs)
sys.exit(1 if diffs else 0)