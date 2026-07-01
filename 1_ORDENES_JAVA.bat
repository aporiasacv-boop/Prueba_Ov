@echo off
title Ordenes de venta - Java puerto 8080
cd /d "%~dp0dynamics-integration"
echo.
echo === 1/3 ORDENES DE VENTA (Java :8080) ===
echo Deje esta ventana abierta.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0dynamics-integration\run.ps1"
pause
