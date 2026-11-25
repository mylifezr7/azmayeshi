@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=%~dp0;%PYTHONPATH%

echo ========================================================
echo  اجرای دستیار اکسل در حالت وب‌سرویس و پیش‌نمایش مرورگر
echo  آدرس در مرورگر شما: http://localhost:8080
echo ========================================================

REM Detect Python
set PYTHON_CMD=
if exist "%~dp0.venv\Scripts\python.exe" (
    set PYTHON_CMD="%~dp0.venv\Scripts\python.exe"
) else if exist "%~dp0venv\Scripts\python.exe" (
    set PYTHON_CMD="%~dp0venv\Scripts\python.exe"
) else (
    where python >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        set PYTHON_CMD=python
    ) else (
        where py >nul 2>nul
        if %ERRORLEVEL% EQU 0 (
            set PYTHON_CMD=py -3
        )
    )
)

if "%PYTHON_CMD%"=="" (
    echo ❌ پایتون در سیستم شما پیدا نشد!
    pause
    exit /b 1
)

start http://localhost:8080
%PYTHON_CMD% -m excel_assistant.main --web --port 8080
pause
