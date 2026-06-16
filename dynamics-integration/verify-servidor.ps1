param(
    [string]$NgrokUrl = ""
)

$ErrorActionPreference = "Stop"
$localBase = "http://localhost:8080"

function Test-Endpoint {
    param([string]$Label, [string]$Url)
    Write-Host ""
    Write-Host "=== $Label ===" -ForegroundColor Cyan
    Write-Host "URL: $Url"
    try {
        $resp = Invoke-RestMethod -Uri $Url -Method Get -TimeoutSec 30
        if ($resp.success) {
            Write-Host "OK: $($resp.message)" -ForegroundColor Green
            return $true
        }
        Write-Host "FALLO: $($resp.message)" -ForegroundColor Red
        return $false
    } catch {
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

Write-Host "Verificacion del servidor de ordenes de venta" -ForegroundColor Yellow
Write-Host "Spring Boot debe estar corriendo (.\run.ps1 en otra ventana)."

$pingOk = Test-Endpoint "1. Spring Boot (local)" "$localBase/api/health/ping"
$dynamicsOk = Test-Endpoint "2. Azure + Dynamics (local)" "$localBase/api/health/dynamics"

if ($NgrokUrl -ne "") {
    $base = $NgrokUrl.TrimEnd("/")
    Test-Endpoint "3. ngrok -> Spring Boot" "$base/api/health/ping" | Out-Null
    Test-Endpoint "4. ngrok -> Dynamics" "$base/api/health/dynamics" | Out-Null
    Write-Host ""
    Write-Host "Pegue esta URL en Resultado!B1 de cada Excel:" -ForegroundColor Yellow
    Write-Host $base
} else {
    Write-Host ""
    Write-Host "Para probar ngrok, copie la URL https de la ventana de ngrok y ejecute:" -ForegroundColor Yellow
    Write-Host "  .\verify-servidor.ps1 -NgrokUrl 'https://xxxx.ngrok-free.dev'"
    Write-Host ""
    Write-Host "Esa misma URL va en Resultado!B1 del Excel."
}

Write-Host ""
if ($pingOk -and $dynamicsOk) {
    Write-Host "Servidor listo para Excel." -ForegroundColor Green
} else {
    Write-Host "Revise application.yml o la conexion a internet." -ForegroundColor Red
}
