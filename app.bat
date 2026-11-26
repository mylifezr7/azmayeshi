@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=%~dp0;%PYTHONPATH%

echo ========================================================
echo   🟢 در حال باز کردن دستیار هوشمند MUYUM در مرورگر وب...
echo   👉 http://localhost:8080
echo ========================================================

start http://localhost:8080
python run.py --web --port 8080
pause
