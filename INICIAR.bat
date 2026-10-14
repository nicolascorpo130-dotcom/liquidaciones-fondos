@echo off
title Liquidaciones y Fondos
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python. Instalalo desde https://www.python.org/downloads/ tildando "Add python.exe to PATH".
  pause
  exit /b
)
rem Si quedo abierta una version anterior del sistema, se cierra para arrancar la nueva
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5002 " ^| findstr LISTENING') do taskkill /PID %%a /F >nul 2>nul
python -c "import flask, openpyxl" 2>nul || python -m pip install -r requirements.txt
echo.
echo Liquidaciones y Fondos: http://127.0.0.1:5002
echo Mientras esta ventana este abierta, el sistema funciona. Cerrala para apagarlo.
echo.
python app.py
pause
