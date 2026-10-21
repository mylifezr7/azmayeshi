@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ============================================
echo    تنظیم ربات اینستاگرام  -  یک بار
echo ============================================
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo [X] Node.js نصب نیست.
  echo     برو به https://nodejs.org نسخه LTS را نصب کن، بعد این فایل را دوباره اجرا کن.
  pause
  exit /b 1
)

if not exist .env copy .env.example .env >nul 2>nul

echo --- قدم 1 از 2: توکن ---
echo توکنی که با IGQ شروع می‌شود را اینجا بچسبان (کلیک راست = چسباندن)
set /p IG_TOKEN=توکن: 
if "!IG_TOKEN!"=="" (
  echo [X] خالی بود.
  pause
  exit /b 1
)

echo.
echo --- قدم 2 از 2: آدرس عمومی ---
echo اگر پنجرهٔ cloudflared باز است، آدرس https://....trycloudflare.com را بچسبان.
echo اگر هنوز نداری فقط Enter بزن (بعداً می‌شود اضافه کرد).
set /p IG_PUBLIC=آدرس: 

echo.
echo [..] شناسهٔ اکانت را خودم پیدا می‌کنم...
node cli.js whoami --save --token=!IG_TOKEN!
if errorlevel 1 (
  echo.
  echo [X] با این توکن اکانتی پیدا نشد. توکن را دوباره از Instagram - API setup with Instagram login بگیر.
  pause
  exit /b 1
)

if not "!IG_PUBLIC!"=="" (
  echo.
  echo IG_PUBLIC_BASE_URL=!IG_PUBLIC!>>.env
  echo [OK] آدرس عمومی ذخیره شد.
)

echo.
echo [..] بررسی نهایی...
node cli.js check
echo.
echo ============================================
echo  اگر بالا پیام @اسم_پیج و HTTP 200 دیدی، آماده است.
echo  بعدش: ویدیو را در پوشهٔ media بگذار و post.bat را بزن.
echo ============================================
echo.
pause
