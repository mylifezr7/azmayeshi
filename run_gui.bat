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
set RUN_ERR=%ERRORLEVEL%

if %RUN_ERR% EQU 0 goto :DONE

echo.
echo ========================================================
echo ⚠️ برنامه گرافیکی متوقف شد (کد خروج: %RUN_ERR%).
echo در صورت بروز خطای عدم وجود کتابخانه‌ها، دستور زیر را اجرا کنید:
echo     %PYTHON_CMD% -m pip install -r requirements.txt
echo.
set CHOICE=
set /p CHOICE="آیا مایلید برنامه در مرورگر وب (Web UI) اجرا شود؟ (Y/N): "
if /i "%CHOICE%"=="Y" (
    echo در حال باز کردن نسخه وب...
    start http://localhost:8080
    %PYTHON_CMD% -m excel_assistant.main --web --port 8080
)

:DONE
echo.
pause
