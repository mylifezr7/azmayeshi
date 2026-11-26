@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=%~dp0;%PYTHONPATH%

echo ========================================================
echo   محاسبه بهای تمام شده کالای فروش رفته (COGS) و انبارگردانی
echo ========================================================
echo.

where git >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo 🔄 در حال بررسی و دریافت آخرین نسخه از گیت‌هاب...
    git fetch origin arena/814f030e-azmayeshi >nul 2>nul
    git checkout -f arena/814f030e-azmayeshi >nul 2>nul
    git reset --hard origin/arena/814f030e-azmayeshi >nul 2>nul
    git pull origin arena/814f030e-azmayeshi >nul 2>nul
)

if "%~1"=="" (
    if exist "Challenge_5_Accounting.xlsx" (
        python calculate_cogs_standalone.py "Challenge_5_Accounting.xlsx"
    ) else if exist "Challenge_4_Accounting.xlsx" (
        python calculate_cogs_standalone.py "Challenge_4_Accounting.xlsx"
    ) else if exist "Challenge_3_Accounting.xlsx" (
        python calculate_cogs_standalone.py "Challenge_3_Accounting.xlsx"
    ) else if exist "Worksheet.xlsx" (
        python calculate_cogs_standalone.py "Worksheet.xlsx"
    ) else (
        python calculate_cogs_standalone.py
    )
) else (
    python calculate_cogs_standalone.py "%~1"
)

echo.
pause
