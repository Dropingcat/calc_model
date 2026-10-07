# -*- coding: utf-8 -*-
"""Проверка корректности сохранённых формул openpyxl (кавычки в XML)."""
import zipfile
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')
p = r'C:\Users\Arhys\Desktop\каллибровка\фракционный состав\АСТМ Д 86-1-4.xlsx'
z = zipfile.ZipFile(p)
xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
# найти L36 в XML
m = re.search(r'<c r="L36"[^>]*>(.*?)</c>', xml)
if m:
    f = re.search(r'<f>(.*?)</f>', m.group(1))
    print('RAW XML <f> L36:', f.group(1) if f else 'no <f>')
# проверить как openpyxl читает обратно
import openpyxl
wb = openpyxl.load_workbook(p, data_only=False)
ws = wb['расчёт']
print('openpyxl readback L36:', repr(ws['L36'].value))
wb.close()