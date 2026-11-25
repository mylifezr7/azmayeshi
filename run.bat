@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=%~dp0;%PYTHONPATH%

echo ========================================================
echo   دستیار هوشمند مالی و حسابداری MUYUM - Excel AI Assistant
echo ========================================================
echo.

where git >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo 🔄 در حال دریافت آخرین آپدیت MUYUM از گیت‌هاب...
    git fetch origin arena/814f030e-azmayeshi >nul 2>nul
    git checkout -f arena/814f030e-azmayeshi >nul 2>nul
    git reset --hard origin/arena/814f030e-azmayeshi >nul 2>nul
    git pull origin arena/814f030e-azmayeshi >nul 2>nul
)

echo در حال اجرای برنامه MUYUM...
python -m excel_assistant.main %*

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ⚠️ در صورت بروز خطا، کتابخانه‌ها را با دستور زیر نصب کنید:
    echo pip install -r requirements.txt
    echo.
    pause
)
