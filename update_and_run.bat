@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================================
echo   دستیار هوشمند اکسل - به‌روزرسانی و اجرا
echo ========================================================
echo.

REM 1. Check Git and Pull
where git >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo [1/3] در حال دریافت آخرین تغییرات از گیت‌هاب...
    git pull origin arena/814f030e-azmayeshi
    if %ERRORLEVEL% NEQ 0 (
        echo ⚠️ دریافت گیت با خطا مواجه شد (ممکن است آفلاین باشید یا تغییرات محلی تداخل داشته باشد).
        echo در حال تلاش برای اجرای برنامه با فایل‌های موجود...
    ) else (
        echo ✅ آخرین نسخه دریافت شد.
    )
) else (
    echo ℹ️ گیت یافت نشد، اجرای برنامه با فایل‌های محلی...
)
echo.

REM 2. Detect Python
echo [2/3] در حال شناسایی مفسر پایتون...
set PYTHON_CMD=

if exist "%~dp0.venv\Scripts\python.exe" (
    set PYTHON_CMD="%~dp0.venv\Scripts\python.exe"
    echo استفاده از محیط مجازی: .venv
) else if exist "%~dp0venv\Scripts\python.exe" (
    set PYTHON_CMD="%~dp0venv\Scripts\python.exe"
    echo استفاده از محیط مجازی: venv
) else (
    where python >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        set PYTHON_CMD=python
    ) else (
        where py >nul 2>nul
        if %ERRORLEVEL% EQU 0 (
            set PYTHON_CMD=py -3
        )
    )
)

if "%PYTHON_CMD%"=="" (
    echo.
    echo ❌ خطای مهم: پایتون (Python) در سیستم شما پیدا نشد!
    echo لطفاً مطمئن شوید پایتون نصب است و گزینه "Add Python to PATH" فعال است.
    echo یا فایل install_windows.bat را اجرا کنید.
    echo.
    pause
    exit /b 1
)

echo پایتون فعال: %PYTHON_CMD%
echo.

REM 3. Run GUI Application
echo [3/3] اجرای برنامه گرافیکی (GUI)...
echo ========================================================
set PYTHONPATH=%~dp0;%PYTHONPATH%

%PYTHON_CMD% -m excel_assistant.main --gui %*
set RUN_ERR=%ERRORLEVEL%

if %RUN_ERR% EQU 0 goto :DONE

echo.
echo ========================================================
echo ⚠️ برنامه گرافیکی متوقف شد (کد خروج: %RUN_ERR%).
echo در صورت بروز خطای عدم وجود کتابخانه‌ها، دستور زیر را اجرا کنید:
echo     %PYTHON_CMD% -m pip install -r requirements.txt
echo.
set CHOICE=
set /p CHOICE="آیا مایلید برنامه در مرورگر وب (Web UI) اجرا شود؟ (Y/N): "
if /i "%CHOICE%"=="Y" (
    echo در حال باز کردن نسخه وب...
    echo آدرس در مرورگر: http://localhost:8080
    start http://localhost:8080
    %PYTHON_CMD% -m excel_assistant.main --web --port 8080
)

:DONE
echo.
pause
