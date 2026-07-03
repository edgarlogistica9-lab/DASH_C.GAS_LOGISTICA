@echo off
setlocal
cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0actualizar_dashboard.ps1"
if errorlevel 1 (
  echo.
  echo No se pudo actualizar el dashboard. Revisa el mensaje anterior.
  pause
  exit /b 1
)

start "" "%~dp0dashboard.html"
exit /b 0
