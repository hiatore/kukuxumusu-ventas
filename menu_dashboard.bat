@echo off
title KUKUXUMUSU - Dashboard
cd /d "%~dp0"
cls
echo.
echo  =============================================
echo      KUKUXUMUSU - SISTEMA DE DASHBOARD
echo  =============================================
echo.
echo  1. Generar dashboard (datos estaticos)
echo  2. Iniciar servidor web local (auto-refresh)
echo  3. Abrir carpeta de datos
echo.
choice /c 123 /n /m "Selecciona opcion (1/2/3): "
echo.

if errorlevel 3 goto carpeta
if errorlevel 2 goto servidor
if errorlevel 1 goto generar

:generar
echo Generando dashboard con datos actualizados...
python3 "C:\Users\hiato\Downloads\generar_dashboard_v3.py"
if %errorlevel% equ 0 (
    echo.
    echo Dashboard generado correctamente.
    start "" "dashboard_ejecutivo.html"
) else (
    echo ERROR al generar el dashboard.
    pause
)
goto fin

:servidor
echo Iniciando servidor web en http://localhost:8080
echo Abre el navegador y ve a esa direccion.
echo Pulsa Ctrl+C en la ventana para detenerlo.
echo.
start http://localhost:8080
python3 "C:\Users\hiato\Downloads\servidor_dashboard.py"
pause
goto fin

:carpeta
start "" "."
goto fin

:fin
echo.
echo Listo.
timeout /t 3 >nul
