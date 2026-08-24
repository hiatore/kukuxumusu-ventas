@echo off
title KUKUXUMUSU - Dashboard
cd /d "%~dp0"
cls
echo.
echo  =============================================
echo      KUKUXUMUSU - Dashboard Ejecutivo
echo  =============================================
echo.
echo  1. Generar dashboard (datos locales)
echo  2. Iniciar servidor web local
echo  3. Publicar en GitHub (commit + push)
echo  4. Abrir carpeta
echo.
choice /c 1234 /n /m "Selecciona opcion (1/2/3/4): "
echo.

if errorlevel 4 goto carpeta
if errorlevel 3 goto github
if errorlevel 2 goto servidor
if errorlevel 1 goto generar

:generar
echo Generando dashboard...
python3 "generar_dashboard.py"
if %errorlevel% equ 0 (
    echo Dashboard generado.
    start "" "dashboard_ejecutivo.html"
) else ( echo ERROR & pause )
goto fin

:servidor
echo Iniciando servidor en http://localhost:8080
start http://localhost:8080
python3 "C:\Users\hiato\Downloads\servidor_dashboard.py"
pause
goto fin

:github
echo Haciendo commit y push a GitHub...
set PATH=%PATH%;C:\Program Files\Git\bin
git add -A
git commit -m "Actualizacion dashboard %date% %time%"
git push
if %errorlevel% equ 0 ( echo Publicado correctamente ) else ( echo ERROR: Configura GitHub primero )
pause
goto fin

:carpeta
start "" "."
goto fin

:fin
echo.
