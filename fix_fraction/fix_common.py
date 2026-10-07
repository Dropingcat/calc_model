# -*- coding: utf-8 -*-
"""
Общие пути и утилиты для fix-скриптов фракционного состава.
"""
import os
import sys

BASE = r'C:\Users\Arhys\Desktop\каллибровка'
FRAC_DIR = os.path.join(BASE, 'фракционный состав')
BACKUP_DIR = os.path.join(FRAC_DIR, 'backup_20261001')
FIX_DIR = os.path.dirname(os.path.abspath(__file__))

# Точные имена файлов (кириллица, как в папке)
ISO1 = os.path.join(FRAC_DIR, 'ГОСТ ISO 3405 расчет для 1_2_3_4 групп.xlsx')
ISO2 = os.path.join(FRAC_DIR, 'ГОСТ ИСО 3405 расчет для 1_2_3_4 групп .xlsx')
D86 = os.path.join(FRAC_DIR, 'АСТМ Д 86-1-4.xlsx')
D1160 = os.path.join(FRAC_DIR, 'астм д 1160.xlsx')
G2177 = os.path.join(FRAC_DIR, 'фракционка ГОСТ_2177.xlsx')


def backup_if_needed():
    """Копирует оригиналы в backup_20261001, если там ещё нет файлов (не перезаписывает)."""
    import shutil
    os.makedirs(BACKUP_DIR, exist_ok=True)
    made = []
    for f in (ISO1, ISO2, D86, D1160, G2177):
        name = os.path.basename(f)
        dst = os.path.join(BACKUP_DIR, name)
        if not os.path.exists(dst):
            shutil.copy2(f, dst)
            made.append(name)
        else:
            print('  backup уже есть:', name)
    return made


def norm_formula(s):
    """Нормализация формулы для сравнения: убрать пробелы, привести к верхнему регистру."""
    if s is None:
        return ''
    return ''.join(str(s).split()).upper()


if __name__ == '__main__':
    print('BASE   =', BASE)
    print('FRAC   =', FRAC_DIR)
    for f in (ISO1, ISO2, D86, D1160):
        print('exists:', os.path.exists(f), '->', os.path.basename(f))