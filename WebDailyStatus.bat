@echo off
setlocal
cd /d "%~dp0"

rem Force JDK 17 for this script session only (system JAVA_HOME is unchanged)
set "PROJECT_JAVA_HOME=D:\VJ\LEARNING\L\java17"

if not exist "%PROJECT_JAVA_HOME%\bin\java.exe" (
    echo JDK 17 not found at:
    echo   %PROJECT_JAVA_HOME%
    echo.
    echo Update PROJECT_JAVA_HOME at the top of run.bat if your JDK 17 is elsewhere.
    pause
    exit /b 1
)

set "JAVA_HOME=%PROJECT_JAVA_HOME%"
set "PATH=%JAVA_HOME%\bin;%PATH%"

echo ========================================
echo Daily Status Report Web
echo Using Java 17 for this window only
echo System JAVA_HOME is NOT modified
echo ========================================
"%JAVA_HOME%\bin\java.exe" -version 2>&1
echo.

rem Stop a previous run so Maven can replace the JAR
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr ":8081" ^| findstr "LISTENING"') do (
    echo Stopping previous instance on port 8081 ^(PID %%a^)...
    taskkill /PID %%a /F >nul 2>&1
)

echo Building...
call mvnw.cmd -q package -DskipTests
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Starting app at http://localhost:8081/
echo Close this window to stop the app.
echo.
"%JAVA_HOME%\bin\java.exe" -jar target\daily-status-report-web-1.0.0.jar
pause
endlocal
