# -*- coding: utf-8 -*-
"""Тест импортов проекта «Анализ актов».

Проверяет:
  1. Все сторонние библиотеки (python_calamine, win32* на Windows).
  2. API python_calamine, который использует act_analysis.py
     (CalamineWorkbook.from_path / get_sheet_by_name / to_python).
  3. Импорт модулей проекта: config, act_analysis (без запуска main()).

Запуск:  python _test_import.py
Код выхода 0 = «Все импорты OK», 1 = есть проблемы.
"""
import sys
import importlib

IS_WIN = sys.platform.startswith("win")
errors = []


def check(name, fn):
    try:
        fn()
        print("  OK  ", name)
    except Exception as e:
        print("  FAIL", name, "-", type(e).__name__, e)
        errors.append(name)


# ------------------ сторонние библиотеки ------------------
check("python_calamine", lambda: importlib.import_module("python_calamine"))


if IS_WIN:
    for m in ("win32com", "win32api", "win32con", "win32com.client"):
        check(m, lambda m=m: importlib.import_module(m))
else:
    print("  SKIP pywin32 (только Windows)")

# ------------------ API calamine, используемый в коде ------------------
def check_calamine_api():
    from python_calamine import CalamineWorkbook, CalamineSheet  # noqa: F401
    assert hasattr(CalamineWorkbook, "from_path"), "нет CalamineWorkbook.from_path"
    assert hasattr(CalamineWorkbook, "get_sheet_by_name"), "нет CalamineWorkbook.get_sheet_by_name"
    assert hasattr(CalamineSheet, "to_python"), "нет CalamineSheet.to_python"
check("calamine API (CalamineWorkbook.from_path)", check_calamine_api)

# ------------------ модули проекта ------------------
check("config", lambda: importlib.import_module("config"))


def check_act_analysis():
    if not IS_WIN:
        raise RuntimeError("пропускается не на Windows: act_analysis требует pywin32")
    importlib.import_module("act_analysis")

if IS_WIN:
    check("act_analysis (импорт без запуска)", check_act_analysis)
else:
    print("  SKIP act_analysis (требует pywin32/Windows; проверен py_compile)")

# ------------------ итог ------------------
print()
if errors:
    print("Провалившие импорты:", ", ".join(errors))
    sys.exit(1)
print("Все импорты OK")
sys.exit(0)
