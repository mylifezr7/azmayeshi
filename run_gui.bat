@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=%~dp0;%PYTHONPATH%
echo ========================================================
echo  اجرای دستیار هوشمند Microsoft Excel 2024 (PySide6 GUI)
echo ========================================================
python -m excel_assistant.main --gui %*
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ⚠️ برنامه متوقف شد. پیام خطای بالا را بررسی کنید.
)
echo.
pause
