@echo off
title Configurar token ngrok (solo UNA vez)
cd /d "%~dp0"
echo.
echo === TOKEN NGROK (solo la primera vez) ===
echo.
echo 1. Abra: https://dashboard.ngrok.com/get-started/your-authtoken
echo 2. Copie el token
echo 3. Peguelo aqui y pulse Enter
echo.
set /p TOKEN=Token: 
if "%TOKEN%"=="" (
    echo Cancelado.
    pause
    exit /b 1
)
ngrok config add-authtoken %TOKEN%
if errorlevel 1 (
    echo Error al guardar token.
    pause
    exit /b 1
)
echo.
echo Listo. El token quedo guardado en su PC.
echo Ya NO necesita volver a pegarlo.
echo Siguiente: doble clic en 2_NGROK.bat
echo.
pause
