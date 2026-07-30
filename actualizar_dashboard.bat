@echo off
title KUKUXUMUSU - Actualizar Dashboard
cd /d "%~dp0"
color 0A
echo ========================================
echo    KUKUXUMUSU - Dashboard Ejecutivo
echo ========================================
echo.
echo Generando dashboard con datos actualizados...
python3 "C:\Users\hiato\Downloads\generar_dashboard_v2.py"
echo.
if %errorlevel% equ 0 (
    echo Dashboard generado correctamente.
    start "" "dashboard_ejecutivo.html"
) else (
    echo ERROR: No se pudo generar el dashboard.
    pause
)
echo.
echo Para actualizar de nuevo, solo ejecuta este archivo.
timeout /t 5
