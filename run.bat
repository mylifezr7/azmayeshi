@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=%~dp0;%PYTHONPATH%

echo ========================================================
echo   دستیار هوشمند اکسل MIYUM - Excel AI Assistant
echo ========================================================
echo.
echo در حال اجرای برنامه...
python -m excel_assistant.main %*

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ⚠️ در صورت بروز خطا، کتابخانه‌ها را با دستور زیر نصب کنید:
    echo pip install -r requirements.txt
    echo.
    pause
)
