@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Sanal ortam bulunamadi.
  echo Lutfen klasor icindeki dosyalari silmeden tekrar deneyin.
  pause
  exit /b 1
)

start "" http://127.0.0.1:5000
".venv\Scripts\python.exe" app.py

if errorlevel 1 (
  echo.
  echo Uygulama acilirken bir hata olustu.
  pause
)
