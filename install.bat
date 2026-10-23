@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

rem ========================================================================
rem install.bat — установка окружения для act_analysis_branch_DP
rem ========================================================================

echo.
echo ========================================================================
echo  Установка окружения для act_analysis_branch_DP
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
python -c "
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
    print()
    print('Не удалось импортировать:', ', '.join(failed))
    sys.exit(1)
print()
print('Все библиотеки доступны.')
"
if errorlevel 1 (
    echo [ОШИБКА] Некоторые библиотеки не импортируются.
    echo.
    pause
    exit /b 1
)

rem ----- проверка Excel (опционально, информация) -----
echo Проверка доступа к Excel ...
python -c "
import win32com.client
try:
    xl = win32com.client.GetActiveObject('Excel.Application')
    print('  OK  Excel запущен (версия:', xl.Version, ')')
except Exception:
    print('  INFO Excel не запущен (запустится при первом вызове скрипта)')
"
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