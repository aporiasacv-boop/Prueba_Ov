@echo off
echo === ngrok: Ordenes de venta (8080) + Olnatura Excel (8011) ===
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0instalar-ngrok-coexistencia.ps1"
if errorlevel 1 pause & exit /b 1
echo.
echo Deje esta ventana abierta. Copie las 2 URLs https que aparezcan abajo.
echo   dynamics  -^> Excel Ordenes (Resultado!B1)
echo   olnatura  -^> Excel NIKZON (modAppConstants / actualizar_url_ngrok.ps1)
echo.
ngrok start dynamics olnatura
