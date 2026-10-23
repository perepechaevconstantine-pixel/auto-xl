@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

rem ========================================================================
rem install.bat — установка окружения для act_analysis (v3.0.1)
rem ========================================================================

echo.
echo ========================================================================
echo  Установка окружения для Анализ актов v3.0.1
echo ========================================================================
echo.

rem ----- определение папки скрипта -----
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo Текущая папка: %CD%
echo.

rem ----- проверка Python -----
where python >nul 2>&1
if errorlevel 1 (
    where py >nul 2>&1
    if errorlevel 1 (
        echo [ОШИБКА] Python не найден в PATH.
        echo Установите Python 3.10+ с https://www.python.org/downloads/
        echo и добавьте его в PATH при установке (галочка "Add Python to PATH").
        echo.
        pause
        exit /b 1
    ) else (
        set "PY=py"
    )
) else (
    set "PY=python"
)

echo Найден Python:
%PY% --version
echo.

rem ----- проверка venv -----
if not exist ".venv\Scripts\activate.bat" (
    echo Создание виртуального окружения .venv ...
    %PY% -m venv .venv
    if errorlevel 1 (
        echo [ОШИБКА] Не удалось создать виртуальное окружение.
        echo.
        pause
        exit /b 1
    )
    echo Готово.
) else (
    echo Виртуальное окружение .venv уже существует.
)
echo.

rem ----- установка зависимостей -----
echo Установка зависимостей из requirements.txt ...
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip >nul 2>&1
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ОШИБКА] Не удалось установить зависимости.
    echo.
    pause
    exit /b 1
)
echo Готово.
echo.

rem ----- проверка библиотек -----
echo Проверка импорта библиотек ...
(
echo import sys
echo mods = ['calamine', 'openpyxl', 'win32com', 'win32api', 'win32con']
echo failed = []
echo for m in mods:
echo     try:
echo         __import^(m^)
echo         print('  OK  ', m)
echo     except Exception as e:
echo         print('  FAIL', m, '-', e)
echo         failed.append(m)
echo if failed:
echo     print('Не удалось импортировать:', ', '.join(failed))
echo     sys.exit(1)
echo print('Все библиотеки доступны.')
) > _check_imports.py
python _check_imports.py
set "CHECK_ERR=!errorlevel!"
del /q _check_imports.py 2>nul
if not "!CHECK_ERR!"=="0" (
    echo [ОШИБКА] Некоторые библиотеки не импортируются.
    echo.
    pause
    exit /b 1
)

rem ----- проверка Excel (опционально, информация) -----
echo Проверка доступа к Excel ...
(
echo import win32com.client
echo try:
echo     xl = win32com.client.GetActiveObject('Excel.Application')
echo     ver = xl.Version
echo     print('  OK  Excel запущен, версия:', ver)
echo except Exception:
echo     print('  INFO Excel не запущен ^(запустится при первом вызове скрипта^)')
) > _check_excel.py
python _check_excel.py
del /q _check_excel.py 2>nul
echo.

rem ----- итог -----
echo ========================================================================
echo  Установка завершена успешно.
echo ========================================================================
echo.
echo Для запуска анализа:
echo   cd /d "%SCRIPT_DIR%"
echo   call .venv\Scripts\activate.bat
echo   python act_analysis.py
echo   или
echo   python act_analysis.py collect
echo.
echo Подробнее читайте в файле INSTALL.md
echo.
pause