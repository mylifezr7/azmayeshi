@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=%~dp0;%PYTHONPATH%
echo ========================================================
echo  اجرای آزمون‌های خودکار دستیار هوشمند اکسل
echo ========================================================
python -m pytest tests/ -v
pause
