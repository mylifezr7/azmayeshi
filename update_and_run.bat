@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================================
echo  دریافت آخرین به‌روزرسانی‌ها از گیت‌هاب...
echo ========================================================
git pull origin arena/814f030e-azmayeshi
echo.
echo ========================================================
echo  اجرای دستیار هوشمند Microsoft Excel (PySide6 GUI)
echo ========================================================
set PYTHONPATH=%~dp0;%PYTHONPATH%
python -m excel_assistant.main --gui %*
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ⚠️ برنامه متوقف شد. پیام خطای بالا را بررسی کنید.
)
echo.
pause
