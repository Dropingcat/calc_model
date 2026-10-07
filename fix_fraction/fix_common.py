# -*- coding: utf-8 -*-
"""
Общие пути и утилиты для fix-скриптов фракционного состава.
"""
import os
import sys

FIX_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_DIR = os.path.dirname(FIX_DIR)

# Пути: по умолчанию — репозиторий (кроссплатформенно), переопределение через env
# FRAC_BASE — корень с папками new_versions_v4 / new_versions_v4.1 (или старая папка «фракционный состав»)
BASE = os.environ.get('FRAC_BASE', REPO_DIR)
FRAC_DIR = os.path.join(BASE, 'фракционный состав')          # старый layout (Windows у заказчика)
NEW_V4 = os.path.join(BASE, 'new_versions_v4')               # книги v3.11 (дефекты ФР-* НЕ исправлены)
NEW_V41 = os.path.join(BASE, 'new_versions_v4.1')            # книги v3.12 (нормативно аудированные)
BACKUP_DIR = os.environ.get('FRAC_BACKUP', os.path.join(FRAC_DIR, 'backup_20261001'))


def _pick(*candidates):
    """Возвращает первый существующий путь; если нет — первый кандидат."""
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return candidates[0] if candidates else None


def _v4(name):
    return os.path.join(NEW_V4, name)


def _v41(prefix):
    """Файл v3.12: ищем по префиксу имени (полное имя длинное, с суффиксом '_нормативно_аудированный')."""
    if not os.path.isdir(NEW_V41):
        return None
    for f in sorted(os.listdir(NEW_V41)):
        if f.startswith(prefix) and f.endswith('.xlsx'):
            return os.path.join(NEW_V41, f)
    return None


# --- Исправленные книги на Windows-машине заказчика (оригинальный контракт test_fix.py) ---
ISO1 = _pick(os.path.join(FRAC_DIR, 'ГОСТ ISO 3405 расчет для 1_2_3_4 групп.xlsx'),
             os.environ.get('FRAC_ISO1'))
ISO2 = _pick(os.path.join(FRAC_DIR, 'ГОСТ ИСО 3405 расчет для 1_2_3_4 групп .xlsx'),
             os.environ.get('FRAC_ISO2'))
D86 = _pick(os.path.join(FRAC_DIR, 'АСТМ Д 86-1-4.xlsx'),
            os.environ.get('FRAC_D86'))
D1160 = _pick(os.path.join(FRAC_DIR, 'астм д 1160.xlsx'),
              os.environ.get('FRAC_D1160'))
G2177 = _pick(os.path.join(FRAC_DIR, 'фракционка ГОСТ_2177.xlsx'),
              os.environ.get('FRAC_G2177'))

# --- Книги из репозитория (для работы здесь) ---
V4 = {
    'ISO3405': _v4('ГОСТ ISO 3405 расчет (v3.11).xlsx'),
    'D86_MANUAL': _v4('АСТМ Д 86 ручной метод (v3.11).xlsx'),
    'D86_AUTO': _v4('АСТМ Д 86 автомат. метод (v3.11).xlsx'),
    'D1160': _v4('астм д 1160 (v3.11).xlsx'),
    'G2177': _v4('фракционка ГОСТ_2177 (v3.11).xlsx'),
    'G2177_APPA': _v4('фракционка ГОСТ_2177 автомат. прил.А (v3.11).xlsx'),
}
V41 = {
    'D86': _v41('ASTM_D86_ручной_метод_v3.12'),
    'G2177': _v41('ГОСТ_2177_метод_А_v3.12'),
}


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