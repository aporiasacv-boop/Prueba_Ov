@echo off
title ngrok - Ordenes 8080 + Olnatura 8011
cd /d "%~dp0"
echo.
echo === 3/3 NGROK (2 tuneles, 2 URLs distintas) ===
echo.
echo Si nunca configuro el token: doble clic en 0_TOKEN_NGROK.bat primero.
echo.
echo Configurando ngrok...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0_configurar_ngrok.ps1"
if errorlevel 1 (
    echo.
    echo ERROR: Falta token. Doble clic en 0_TOKEN_NGROK.bat
    pause
    exit /b 1
)
echo.
echo Arrancando...
echo   dynamics  -^> 8080  (URL fija: unexpired-joyfully-exfoliate...)
echo   olnatura  -^> 8011  (URL nueva, distinta a la de arriba)
echo.
echo Si las 2 lineas muestran la MISMA url, cierre ngrok y vuelva a abrir 2_NGROK.bat
echo Para ver URLs claras: VER_URLS_NGROK.bat o http://127.0.0.1:4040
echo.
ngrok start dynamics olnatura
pause
