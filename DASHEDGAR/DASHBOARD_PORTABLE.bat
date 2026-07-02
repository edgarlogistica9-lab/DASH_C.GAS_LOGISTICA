@echo off
setlocal
cd /d "%~dp0"

title Dashboard Envios - modo portable
echo.
echo ============================================================
echo  DASHBOARD ENVIOS - MODO PORTABLE
echo ============================================================
echo  Carpeta: %CD%
echo.
echo  Fuente: ENVIOS ARCHIVO GENERAL.xlsm
echo  Hoja:   Enero2026
echo.
echo  Se abrira el dashboard y se vigilara el archivo Excel.
echo  Al guardar cambios en Excel, dashboard.html se regenerara.
echo  Para salir, cierra esta ventana.
echo ============================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0dashboard_portable.ps1"
if errorlevel 1 (
  echo.
  echo No se pudo iniciar el modo portable. Revisa el mensaje anterior.
  pause
  exit /b 1
)

exit /b 0
