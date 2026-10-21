@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ============================================
echo   تنظیم ربات اینستاگرام - فقط یک بار
echo ============================================
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo [X] Node.js نصب نیست.
  echo     برو به https://nodejs.org و نسخه LTS را نصب کن، بعد این فایل را دوباره اجرا کن.
  pause
  exit /b 1
)

if not exist .env (
  if exist .env.example (
    copy .env.example .env >nul
    echo [OK] فایل .env ساخته شد.
  ) else (
    echo [X] فایل .env.example پیدا نشد. مطمئن شو کل پوشه instagram-uploader را داری.
    pause
    exit /b 1
  )
)

echo.
echo --- قدم 1: شناسهٔ اکانت اینستاگرام ---
echo (یک عدد بلند است، نه اسم پیج)
set /p IG_ID=شناسه را اینجا بچسبان و Enter بزن: 
if "!IG_ID!"=="" (
  echo [X] خالی بود.
  pause
  exit /b 1
)

echo.
echo --- قدم 2: توکن دسترسی ---
echo (رشتهٔ بلندی که با IGQ شروع می‌شود)
set /p IG_TOKEN=توکن را اینجا بچسبان و Enter بزن: 
if "!IG_TOKEN!"=="" (
  echo [X] خالی بود.
  pause
  exit /b 1
)

echo.
echo --- قدم 3: آدرس عمومی ---
echo اگر تونل را بالا آورده‌ای، آدرسش را بچسبان (مثل https://xxxx.trycloudflare.com)
echo اگر هنوز نداری، فقط Enter بزن.
set /p IG_PUBLIC=آدرس: 

echo.>>.env
echo IG_USER_ID=!IG_ID!>>.env
echo IG_ACCESS_TOKEN=!IG_TOKEN!>>.env
if not "!IG_PUBLIC!"=="" echo IG_PUBLIC_BASE_URL=!IG_PUBLIC!>>.env

echo.
echo [OK] ذخیره شد در .env
echo.
echo --- بررسی تنظیمات ---
node cli.js check
echo.
echo اگر همه چیز سبز بود، برای انتشار واقعی:  post.bat
echo.
pause
