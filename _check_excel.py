# -*- coding: utf-8 -*-
"""Информационная проверка: запущен ли Excel (COM). Не влияет на код выхода."""
try:
    import win32com.client
    xl = win32com.client.GetActiveObject('Excel.Application')
    print('  OK  Excel запущен, версия:', xl.Version)
except ImportError:
    print('  INFO pywin32 не установлен - см. проверку библиотек выше')
except Exception:
    print('  INFO Excel не запущен (запустится при первом вызове скрипта)')
