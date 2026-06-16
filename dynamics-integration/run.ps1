$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Get-Command java -ErrorAction SilentlyContinue)) {
    Write-Host "Instale Java 17 o superior." -ForegroundColor Red
    exit 1
}

$wrapperJar = Join-Path $PSScriptRoot ".mvn\wrapper\maven-wrapper.jar"
if (-not (Test-Path $wrapperJar)) {
    $jarDir = Split-Path $wrapperJar -Parent
    New-Item -ItemType Directory -Force -Path $jarDir | Out-Null
    $wrapperUrl = "https://repo.maven.apache.org/maven2/org/apache/maven/wrapper/maven-wrapper/3.3.2/maven-wrapper-3.3.2.jar"
    Invoke-WebRequest -Uri $wrapperUrl -OutFile $wrapperJar -UseBasicParsing
}

java -classpath $wrapperJar "-Dmaven.multiModuleProjectDirectory=$PSScriptRoot" org.apache.maven.wrapper.MavenWrapperMain spring-boot:run
