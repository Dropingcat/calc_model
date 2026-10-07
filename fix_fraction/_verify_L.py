# -*- coding: utf-8 -*-
# NOTE(Windows): архивная утилита починки Windows-копий книг (ФР-*).
# Путь по умолчанию переопределяется через FRAC_WIN_DIR.
"""Верификация L-формул D1160 после доводки."""
import os
import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')
p = os.environ.get('FRAC_WIN_DIR', r'C:\Users\Arhys\Desktop\каллибровка\фракционный состав\астм д 1160.xlsx')
wb = openpyxl.load_workbook(p, data_only=False)
ws = wb['Лист2']
print('--- L5..L15 после доводки ---')
for r in range(5, 16):
    print('L%d:' % r, ws['L%d' % r].value)
wb.close()