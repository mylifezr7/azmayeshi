@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=%~dp0;%PYTHONPATH%

echo ========================================================
echo   🟢 در حال اجرای دستیار هوشمند Microsoft Excel
echo ========================================================
echo.

REM Try launching PySide6 GUI first
python -c "import PySide6" >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo ✅ در حال باز کردن پنجره گرافیکی دستیار اکسل...
    python run.py --gui
) else (
    echo 🌐 پکیج PySide6 یافت نشد؛ در حال باز کردن پنل هوشمند در مرورگر...
    start http://localhost:8080
    python run.py --web --port 8080
)
