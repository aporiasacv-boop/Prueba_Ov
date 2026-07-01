@echo off
title ngrok - Solo Ordenes de venta (8080)
cd /d "%~dp0"
echo.
echo === NGROK ORDENES (puerto 8080) ===
echo.
echo El plan gratis de ngrok NO permite 2 URLs distintas bien.
echo Olnatura (8011) usa otro tunel en Excel_Restringido-\2_TUNEL_OLNATURA.bat
echo.
if not exist "%~dp00_TOKEN_NGROK.bat" goto config
echo Si nunca configuro token: 0_TOKEN_NGROK.bat
echo.
:config
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0_configurar_ngrok.ps1"
if errorlevel 1 (
    echo ERROR: Falta token. Doble clic en 0_TOKEN_NGROK.bat
    pause
    exit /b 1
)
echo.
echo URL fija: https://unexpired-joyfully-exfoliate.ngrok-free.dev
echo Pegar en Excel Ordenes, hoja Resultado, celda B1
echo.
ngrok start dynamics
pause
