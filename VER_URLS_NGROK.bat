@echo off
title Ver URL ngrok Ordenes
echo.
powershell -NoProfile -Command "try { $t = (Invoke-RestMethod http://127.0.0.1:4040/api/tunnels).tunnels; foreach ($x in $t) { Write-Host ('  ' + $x.public_url + ' -> ' + $x.config.addr) } } catch { Write-Host '  ngrok no esta corriendo.' }"
echo.
pause
