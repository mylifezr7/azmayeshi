@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=%~dp0;%PYTHONPATH%
echo ========================================================
echo  اجرای دستیار اکسل در حالت وب‌سرویس و پیش‌نمایش مرورگر
echo  آدرس در مرورگر شما: http://localhost:8080
echo ========================================================
python -m excel_assistant.main --web --port 8080
pause
