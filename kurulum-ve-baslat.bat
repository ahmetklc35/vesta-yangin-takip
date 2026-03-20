@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Python ortami hazirlaniyor...
  python -m venv .venv
)

echo Gereken kutuphaneler kontrol ediliyor...
".venv\Scripts\python.exe" -m pip install -r requirements.txt

echo Uygulama aciliyor...
start "" http://127.0.0.1:5000
".venv\Scripts\python.exe" app.py

if errorlevel 1 (
  echo.
  echo Uygulama acilirken bir hata olustu.
  pause
)
