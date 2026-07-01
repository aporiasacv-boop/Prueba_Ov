@echo off
title Liberar puertos 8080 y 8011
echo.
echo === CERRAR procesos viejos en puertos 8080 y 8011 ===
echo (Si abrio los .bat varias veces, los puertos quedan ocupados)
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ports = 8080, 8011; foreach ($p in $ports) { $c = Get-NetTCPConnection -LocalPort $p -State Listen -ErrorAction SilentlyContinue; if (-not $c) { Write-Host ('Puerto ' + $p + ': libre'); continue }; foreach ($x in $c) { $pid = $x.OwningProcess; $n = (Get-Process -Id $pid -ErrorAction SilentlyContinue).ProcessName; Write-Host ('Puerto ' + $p + ': cerrando ' + $n + ' (PID ' + $pid + ')'); Stop-Process -Id $pid -Force -ErrorAction SilentlyContinue } }"
echo.
echo Listo. Ahora en orden:
echo   1. Prueba_Ov\1_ORDENES_JAVA.bat
echo   2. Excel\1_API_OLNATURA.bat
echo   3. Prueba_Ov\2_NGROK.bat
echo   4. Excel\2_TUNEL_OLNATURA.bat
echo.
pause
