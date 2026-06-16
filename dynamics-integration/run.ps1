param(
    [switch]$ConfigurarCredenciales
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

. "$PSScriptRoot\ensure-credentials.ps1"
$configPath = Join-Path $PSScriptRoot "src\main\resources\application.yml"
if ($ConfigurarCredenciales) {
    Ensure-DynamicsCredentials -ConfigPath $configPath -ForcePrompt
} else {
    Ensure-DynamicsCredentials -ConfigPath $configPath
}

if (-not (Get-Command java -ErrorAction SilentlyContinue)) {
    Write-Host "Instale Java 17 o superior." -ForegroundColor Red
    exit 1
}

$puerto8080 = Get-NetTCPConnection -LocalPort 8080 -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
if ($puerto8080) {
    $proc = Get-Process -Id $puerto8080.OwningProcess -ErrorAction SilentlyContinue
    $nombre = if ($proc) { $proc.ProcessName } else { "desconocido" }
    if ($nombre -ne "java") {
        Write-Host "Puerto 8080 ocupado por '$nombre' (PID $($puerto8080.OwningProcess))." -ForegroundColor Red
        Write-Host "Detenga ese proceso o cambie server.port en application.yml." -ForegroundColor Yellow
        Write-Host "Ejemplo: Stop-Process -Id $($puerto8080.OwningProcess) -Force" -ForegroundColor Yellow
        exit 1
    }
}

$wrapperJar = Join-Path $PSScriptRoot ".mvn\wrapper\maven-wrapper.jar"
if (-not (Test-Path $wrapperJar)) {
    $jarDir = Split-Path $wrapperJar -Parent
    New-Item -ItemType Directory -Force -Path $jarDir | Out-Null
    $wrapperUrl = "https://repo.maven.apache.org/maven2/org/apache/maven/wrapper/maven-wrapper/3.3.2/maven-wrapper-3.3.2.jar"
    Invoke-WebRequest -Uri $wrapperUrl -OutFile $wrapperJar -UseBasicParsing
}

java -classpath $wrapperJar "-Dmaven.multiModuleProjectDirectory=$PSScriptRoot" org.apache.maven.wrapper.MavenWrapperMain spring-boot:run
