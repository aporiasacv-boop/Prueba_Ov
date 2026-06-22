$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$shared = Get-Content (Join-Path $root "_shared.osts.ts") -Raw -Encoding UTF8
$outDir = Join-Path $root "dist"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

@(
    @{ File = "ProbarConexion.osts.ts"; Main = "await probarConexion(workbook);" },
    @{ File = "CrearPedido.osts.ts"; Main = "await crearPedido(workbook);" },
    @{ File = "CrearLineas.osts.ts"; Main = "await crearLineas(workbook);" },
    @{ File = "SubirLote.osts.ts"; Main = "await subirLote(workbook);" },
    @{ File = "CrearCabeceraMarketing.osts.ts"; Main = "await crearCabeceraMarketing(workbook);" },
    @{ File = "CrearOrden.osts.ts"; Main = "await crearOrdenMarketing(workbook);" },
    @{ File = "CrearLineasMarketing.osts.ts"; Main = "await crearLineasMarketing(workbook);" },
    @{ File = "SubirMarketing.osts.ts"; Main = "await subirMarketing(workbook);" },
    @{ File = "CrearCabeceraComercial.osts.ts"; Main = "await crearCabeceraComercial(workbook);" },
    @{ File = "CrearLineasComercial.osts.ts"; Main = "await crearLineasComercial(workbook);" }
) | ForEach-Object {
    $content = $shared.TrimEnd() + "`r`n`r`nasync function main(workbook: ExcelScript.Workbook): Promise<void> {`r`n  $($_.Main)`r`n}`r`n"
    [System.IO.File]::WriteAllText((Join-Path $outDir $_.File), $content, [System.Text.UTF8Encoding]::new($false))
}
