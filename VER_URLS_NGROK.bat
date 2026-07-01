@echo off
title Ver URLs de ngrok
echo.
echo Abra en el navegador: http://127.0.0.1:4040
echo.
echo Deben verse 2 URLs DIFERENTES:
echo   una para puerto 8080 (Ordenes)
echo   otra para puerto 8011 (Olnatura)
echo.
powershell -NoProfile -Command "try { $t = (Invoke-RestMethod http://127.0.0.1:4040/api/tunnels).tunnels; foreach ($x in $t) { Write-Host ('  ' + $x.public_url + ' -> ' + $x.config.addr) } } catch { Write-Host '  ngrok no esta corriendo. Arranque 2_NGROK.bat primero.' }"
echo.
pause
