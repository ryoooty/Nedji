@echo off
:: Запуск Python-скрипта от имени администратора
set SCRIPT_PATH=%~dp0NeDja.py
set PYTHON_PATH=python

:: Проверка наличия прав администратора
net session >nul 2>&1
if %errorLevel% == 0 (
    %PYTHON_PATH% "%SCRIPT_PATH%"
) else (
    echo Требуются права администратора. Запуск...
    powershell -Command "Start-Process '%PYTHON_PATH%' -ArgumentList '%SCRIPT_PATH%' -Verb RunAs"
)
pause