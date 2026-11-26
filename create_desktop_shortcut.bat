@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================================
echo  ایجاد آیکون میانبر دستیار اکسل روی دسکتاپ ویندوز
echo ========================================================
powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $d = [System.Environment]::GetFolderPath('Desktop'); $s = $ws.CreateShortcut([System.IO.Path]::Combine($d, 'Excel AI Assistant.lnk')); $s.TargetPath = '%~dp0run_gui.bat'; $s.WorkingDirectory = '%~dp0'; $s.WindowStyle = 1; $s.Description = 'Excel AI Assistant'; $s.Save()"
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ✅ آیکون میانبر «Excel AI Assistant» با موفقیت روی دسکتاپ شما ایجاد شد!
    echo زین‌پس نیازی به باز کردن CMD ندارید؛ فقط روی آیکون دسکتاپ دابل‌کلیک کنید.
) else (
    echo ⚠️ خطا در ایجاد میانبر دسکتاپ.
)
echo.
pause
