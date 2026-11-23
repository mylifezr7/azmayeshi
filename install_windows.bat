@echo off
chcp 65001 >nul
echo ========================================================
echo  نصب وابستگی‌های دستیار هوشمند Microsoft Excel 2024
echo ========================================================
echo.
cd /d "%~dp0"
echo در حال ارتقا و نصب بسته‌های مورد نیاز پایتون...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install -e .
echo.
echo ========================================================
echo  ✅ نصب با موفقیت انجام شد!
echo  اکنون می‌توانید فایل run_gui.bat یا run_web.bat را اجرا کنید.
echo ========================================================
pause
