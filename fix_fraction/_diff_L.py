# -*- coding: utf-8 -*-
# NOTE(Windows): архивная утилита починки Windows-копий книг (ФР-*).
# Путь по умолчанию переопределяется через FRAC_WIN_DIR.
"""Diff D1160: было (бэкап) -> стало для L5..L15 и M11..M15."""
import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')
backup = os.environ.get('FRAC_WIN_DIR', r'C:\Users\Arhys\Desktop\каллибровка\фракционный состав\backup_20261001\астм д 1160.xlsx')
frac = os.environ.get('FRAC_WIN_DIR', r'C:\Users\Arhys\Desktop\каллибровка\фракционный состав\астм д 1160.xlsx')


def get(path, sheet, cell):
    wb = openpyxl.load_workbook(path, data_only=False)
    v = wb[sheet][cell].value
    wb.close()
    return v


print('--- L5..L15: было -> стало ---')
for r in range(5, 16):
    old = get(backup, 'Лист2', 'L%d' % r)
    new = get(frac, 'Лист2', 'L%d' % r)
    if old != new:
        print('L%d:' % r)
        print('  OLD:', repr(old)[:220])
        print('  NEW:', repr(new)[:220])
    else:
        print('L%d: без изменений' % r)