$ngrokDir = Join-Path $env:LOCALAPPDATA "ngrok"
$target = Join-Path $ngrokDir "ngrok.yml"
New-Item -ItemType Directory -Force -Path $ngrokDir | Out-Null

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
  olnatura:
    addr: 8011
    proto: http
"@ | Set-Content -Path $target -Encoding UTF8

Write-Host "OK: $target"
