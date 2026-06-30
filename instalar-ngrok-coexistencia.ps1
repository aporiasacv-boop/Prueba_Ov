# Fusiona túneles dynamics (8080) + olnatura (8011) en %LOCALAPPDATA%\ngrok\ngrok.yml
# Conserva el authtoken que ya tenga ngrok instalado en la PC.

$ErrorActionPreference = "Stop"
$ngrokDir = Join-Path $env:LOCALAPPDATA "ngrok"
$target = Join-Path $ngrokDir "ngrok.yml"
New-Item -ItemType Directory -Force -Path $ngrokDir | Out-Null

$token = $null
if (Test-Path $target) {
    $existing = Get-Content $target -Raw
    if ($existing -match '(?m)^\s*authtoken:\s*(\S+)') {
        $token = $Matches[1]
    }
}

if (-not $token) {
    Write-Host ""
    Write-Host "No se encontro authtoken en $target" -ForegroundColor Yellow
    Write-Host "Ejecute una vez: ngrok config add-authtoken SU_TOKEN" -ForegroundColor Yellow
    Write-Host "Luego vuelva a correr este script." -ForegroundColor Yellow
    exit 1
}

$content = @"
version: 3

agent:
  authtoken: $token

tunnels:
  dynamics:
    addr: 8080
    proto: http
  olnatura:
    addr: 8011
    proto: http
"@

Set-Content -Path $target -Value $content.TrimEnd() -Encoding UTF8

Write-Host ""
Write-Host "Listo: $target" -ForegroundColor Green
Write-Host "Túneles: dynamics (8080) + olnatura (8011)" -ForegroundColor Green
Write-Host "Arrancar: ngrok start dynamics olnatura" -ForegroundColor Cyan
