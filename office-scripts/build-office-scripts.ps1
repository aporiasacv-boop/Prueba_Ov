$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$shared = Get-Content (Join-Path $root "_shared.osts.ts") -Raw -Encoding UTF8
$probarConexion = Get-Content (Join-Path $root "_probar_conexion.osts.ts") -Raw -Encoding UTF8
$outDir = Join-Path $root "dist"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

@(
    @{ File = "ProbarConexion.osts.ts"; Source = $probarConexion; Main = "await probarConexion(workbook);" },
    @{ File = "CrearPedido.osts.ts"; Source = $shared; Main = "await crearPedido(workbook);" },
    @{ File = "CrearLineas.osts.ts"; Source = $shared; Main = "await crearLineas(workbook);" },
    @{ File = "SubirLote.osts.ts"; Source = $shared; Main = "await subirLote(workbook);" },
    @{ File = "CrearCabeceraMarketing.osts.ts"; Source = $shared; Main = "await crearCabeceraMarketing(workbook);" },
    @{ File = "CrearOrden.osts.ts"; Source = $shared; Main = "await crearOrdenMarketing(workbook);" },
    @{ File = "CrearLineasMarketing.osts.ts"; Source = $shared; Main = "await crearLineasMarketing(workbook);" },
    @{ File = "SubirMarketing.osts.ts"; Source = $shared; Main = "await subirMarketing(workbook);" },
    @{ File = "CrearCabeceraComercial.osts.ts"; Source = $shared; Main = "await crearCabeceraComercial(workbook);" },
    @{ File = "CrearLineasComercial.osts.ts"; Source = $shared; Main = "await crearLineasComercial(workbook);" }
) | ForEach-Object {
    if ($_.File -eq "ProbarConexion.osts.ts") {
        $content = $_.Source.TrimEnd() + "`r`n`r`nasync function main(workbook: ExcelScript.Workbook): Promise<void> {`r`n  await pcProbarConexion(workbook);`r`n}`r`n"
    } else {
        $content = $_.Source.TrimEnd() + "`r`n`r`nasync function main(workbook: ExcelScript.Workbook): Promise<void> {`r`n  $($_.Main)`r`n}`r`n"
    }
    [System.IO.File]::WriteAllText((Join-Path $outDir $_.File), $content, [System.Text.UTF8Encoding]::new($false))
}
