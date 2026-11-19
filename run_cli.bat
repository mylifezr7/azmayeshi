@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=%~dp0;%PYTHONPATH%
echo ========================================================
echo  اجرای دستیار اکسل در حالت ترمینال خط فرمان (CLI)
echo ========================================================
python -m excel_assistant.main --cli
pause
