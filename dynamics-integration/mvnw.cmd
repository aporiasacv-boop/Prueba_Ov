@REM Maven Wrapper - ejecutar sin instalar Maven global
@REM Uso: mvnw.cmd spring-boot:run

@echo off
setlocal

set "MAVEN_PROJECTBASEDIR=%~dp0"
set "WRAPPER_JAR=%MAVEN_PROJECTBASEDIR%.mvn\wrapper\maven-wrapper.jar"
set "WRAPPER_LAUNCHER=org.apache.maven.wrapper.MavenWrapperMain"

if not exist "%WRAPPER_JAR%" (
  echo Error: no se encuentra maven-wrapper.jar en .mvn\wrapper\
  exit /b 1
)

set "JAVA_EXE=java"
where java >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
  echo Error: Java no esta instalado o no esta en el PATH.
  echo Instala Java 21: winget install Microsoft.OpenJDK.21
  exit /b 1
)

"%JAVA_EXE%" ^
  -classpath "%WRAPPER_JAR%" ^
  "-Dmaven.multiModuleProjectDirectory=%MAVEN_PROJECTBASEDIR%" ^
  %WRAPPER_LAUNCHER% %*
if ERRORLEVEL 1 exit /b %ERRORLEVEL%

endlocal
