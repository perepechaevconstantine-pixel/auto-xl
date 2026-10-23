@echo off
chcp 866 >nul
setlocal enabledelayedexpansion

rem ========================================================================
rem install.bat - установка окружения для act_analysis (v3.1.0)
rem ========================================================================

echo.
echo ========================================================================
echo  Установка окружения для Анализ актов v3.1.0
echo ========================================================================
echo.

rem ----- определение папки скрипта -----
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo Текущая папка: %CD%
echo.

rem ----- проверка Python -----
set "PY="
where python >nul 2>&1
if not errorlevel 1 set "PY=python"
if defined PY goto :py_ok
where py >nul 2>&1
if not errorlevel 1 set "PY=py"
if defined PY goto :py_ok
echo [ОШИБКА] Python не найден в PATH.
echo Установите Python 3.10+ с https://www.python.org/downloads/
echo и добавьте его в PATH при установке (галочка "Add Python to PATH").
echo.
pause
exit /b 1
:py_ok

echo Найден Python:
%PY% --version
echo.

rem ----- проверка venv -----
if exist ".venv\Scripts\activate.bat" goto :venv_skip
echo Создание виртуального окружения .venv ...
%PY% -m venv .venv
if errorlevel 1 goto :err_venv
echo Готово.
goto :venv_done
:venv_skip
echo Виртуальное окружение .venv уже существует.
:venv_done
echo.

rem ----- установка зависимостей -----
echo Установка зависимостей из requirements.txt ...
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip >nul 2>&1
python -m pip install -r requirements.txt
if errorlevel 1 goto :err_pip
echo Готово.
echo.

rem ----- проверка библиотек -----
echo Проверка импорта библиотек ...
copy /y "%SCRIPT_DIR%_check_imports.py" _check_imports.py >nul
if not exist _check_imports.py (
    echo [ОШИБКА] Не найден файл _check_imports.py рядом с install.bat.
    goto :fail_end
)
python _check_imports.py
if errorlevel 1 goto :err_libs
del /q _check_imports.py 2>nul

rem ----- проверка Excel (опционально, информация) -----
echo Проверка доступа к Excel ...
copy /y "%SCRIPT_DIR%_check_excel.py" _check_excel.py >nul
python _check_excel.py
del /q _check_excel.py 2>nul
echo.

rem ----- создание папки логов -----
if not exist "logs" mkdir "logs"
if exist "logs" (echo Папка logs/ создана.) else (echo [ВНИМАНИЕ] Не удалось создать папку logs/.)
echo.

goto :ok_end

rem ----- обработчики ошибок -----
:err_venv
echo [ОШИБКА] Не удалось создать виртуальное окружение.
goto :fail_end

:err_pip
echo [ОШИБКА] Не удалось установить зависимости.
goto :fail_end

:err_libs
echo [ОШИБКА] Некоторые библиотеки не импортируются.
goto :fail_end

:fail_end
echo.
pause
exit /b 1

rem ----- итог -----
:ok_end
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