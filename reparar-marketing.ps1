# Repara Excel marketing roto. Cierre Excel antes de ejecutar.
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Test-Path "Prueba_OV marketing.xlsx")) {
    Write-Host "ERROR: Falta 'Prueba_OV marketing.xlsx' en esta carpeta."
    Write-Host "Restaure el archivo original desde OneDrive (Version anterior) y vuelva a ejecutar."
    exit 1
}

python reparar_marketing_xlsx.py
python prepare_marketing_excel.py
python analyze_marketing_tables.py "Prueba_OV marketing - REPARADO.xlsx"
Write-Host ""
Write-Host "Abra: Prueba_OV marketing - REPARADO.xlsx (prueba)"
Write-Host "Luego: Prueba_OV marketing - Preparado.xlsx (con panel y Resultado)"
