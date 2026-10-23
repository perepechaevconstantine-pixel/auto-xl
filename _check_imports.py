# -*- coding: utf-8 -*-
"""Проверка доступности библиотек, необходимых act_analysis.py.
Вызывается из install.bat. Код выхода 0 = все импорты ОК, 1 = есть проблемы."""
import sys

# Пакет ставится как python-calamine, импортируется как python_calamine.
# win32* требует pywin32 (Windows-only post-install скрипт не запускался?).
mods = ['python_calamine', 'openpyxl', 'win32com', 'win32api', 'win32con']
failed = []
for m in mods:
    try:
        __import__(m)
        print('  OK  ', m)
    except Exception as e:
        print('  FAIL', m, '-', e)
        failed.append(m)

if failed:
    print('Не удалось импортировать:', ', '.join(failed))
    sys.exit(1)

# Проверка API calamine: act_analysis.py использует CalamineWorkbook.from_path
# и get_sheet_by_name(...).to_python()
try:
    from python_calamine import CalamineWorkbook
    assert hasattr(CalamineWorkbook, 'from_path'), 'нет CalamineWorkbook.from_path'
    print('  OK   python_calamine.CalamineWorkbook (API совместим)')
except Exception as e:
    print('  FAIL python_calamine API -', e)
    sys.exit(1)

print('Все библиотеки доступны.')
