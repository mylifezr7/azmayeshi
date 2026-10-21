@echo off
rem اجرای یک‌بارهٔ ربات روی ویندوز
cd /d "%~dp0"
if not exist .env (
  echo فایل .env پیدا نشد. اول: copy .env.example .env و مقادیر را پر کن.
  pause
  exit /b 1
)
node cli.js post %*
if errorlevel 1 pause
