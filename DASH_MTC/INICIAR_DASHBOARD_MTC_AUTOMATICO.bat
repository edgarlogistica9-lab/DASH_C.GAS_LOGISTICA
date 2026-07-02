@echo off
setlocal
cd /d "%~dp0"

start "Dashboard MTC Automatico" powershell -NoProfile -ExecutionPolicy Bypass -NoExit -File "%~dp0actualizar_dashboard_mtc.ps1" -Watch
timeout /t 2 >nul
start "" "%~dp0dashboard_mtc.html"
exit /b 0
