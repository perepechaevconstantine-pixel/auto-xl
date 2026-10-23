# -*- coding: utf-8 -*-
"""Проверка доступности библиотек, необходимых act_analysis.py.
Вызывается из install.bat. Код выхода 0 = все импорты ОК, 1 = есть проблемы."""
import sys

mods = ['calamine', 'openpyxl', 'win32com', 'win32api', 'win32con']
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

print('Все библиотеки доступны.')
