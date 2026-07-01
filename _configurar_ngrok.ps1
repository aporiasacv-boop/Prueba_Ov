$ngrokDir = Join-Path $env:LOCALAPPDATA "ngrok"
$target = Join-Path $ngrokDir "ngrok.yml"
New-Item -ItemType Directory -Force -Path $ngrokDir | Out-Null

# Dominio fijo gratis de la cuenta (Ordenes de venta -> 8080)
$dominioOrdenes = "unexpired-joyfully-exfoliate.ngrok-free.dev"

$token = $null
if (Test-Path $target) {
    $existing = Get-Content $target -Raw
    if ($existing -match '(?m)^\s*authtoken:\s*(\S+)') { $token = $Matches[1] }
}
if (-not $token) { exit 1 }

@"
version: 3

agent:
  authtoken: $token

tunnels:
  dynamics:
    addr: 8080
    proto: http
    domain: $dominioOrdenes
  olnatura:
    addr: 8011
    proto: http
"@ | Set-Content -Path $target -Encoding UTF8

Write-Host "OK: $target"
Write-Host "  dynamics -> $dominioOrdenes (8080)"
Write-Host "  olnatura -> URL nueva automatica (8011)"
