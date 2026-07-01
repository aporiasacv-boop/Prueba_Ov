@echo off
title ngrok - Ordenes 8080 + Olnatura 8011
cd /d "%~dp0"
echo.
echo === 3/3 NGROK (2 tuneles) ===
echo.
echo Configurando ngrok...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0_configurar_ngrok.ps1"
if errorlevel 1 (
    echo.
    echo ERROR: Falta authtoken de ngrok.
    echo Ejecute UNA VEZ en esta ventana:
    echo   ngrok config add-authtoken SU_TOKEN
    echo.
    echo Luego vuelva a dar doble clic en 2_NGROK.bat
    pause
    exit /b 1
)
echo.
echo Arrancando tuneles...
echo   dynamics  -^> localhost:8080  (Ordenes de venta, celda Resultado!B1)
echo   olnatura  -^> localhost:8011  (Excel NIKZON login)
echo.
echo Copie las 2 URLs https que aparezcan abajo.
echo Deje esta ventana abierta.
echo.
ngrok start dynamics olnatura
pause
