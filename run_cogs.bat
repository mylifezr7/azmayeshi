@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================================
echo   محاسبه بهای تمام شده کالای فروش رفته (COGS) و انبارگردانی
echo ========================================================
echo.

if "%~1"=="" (
    if exist "Challenge_3_Accounting.xlsx" (
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
