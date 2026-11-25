@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=%~dp0;%PYTHONPATH%

echo ========================================================
echo  اجرای دستیار هوشمند Microsoft Excel 2024 (PySide6 GUI)
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

%PYTHON_CMD% -m excel_assistant.main --gui %*
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ⚠️ برنامه گرافیکی متوقف شد. پیام خطای بالا را بررسی کنید.
    echo.
    echo آیا مایلید برنامه در مرورگر وب (Web UI) اجرا شود؟ (Y/N)
    set /p CHOICE="انتخاب شما: "
    if /i "%CHOICE%"=="Y" (
        echo در حال باز کردن نسخه وب...
        start http://localhost:8080
        %PYTHON_CMD% -m excel_assistant.main --web --port 8080
    )
)
echo.
pause
