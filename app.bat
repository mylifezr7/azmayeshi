@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=%~dp0;%PYTHONPATH%

echo ========================================================
echo   🖥️ در حال باز کردن پنجره دسکتاپ دستیار هوشمند اکسل...
echo ========================================================
echo.

REM بررسی و نصب خودکار PySide6 در صورت نیاز
python -c "import PySide6" >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo ⏳ در حال نصب پکیج گرافیکی PySide6...
    python -m pip install PySide6
)

python run.py --gui
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ⚠️ پنجره دسکتاپ اجرا نشد. لطفاً دستور زیر را در CMD اجرا کنید:
    echo python -m pip install PySide6
    pause
)
