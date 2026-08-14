@echo off
cd /d "%~dp0"

if not defined JAVA_HOME (
    if exist "D:\VJ\LEARNING\L\java17\bin\java.exe" (
        set "JAVA_HOME=D:\VJ\LEARNING\L\java17"
    ) else (
        echo Set JAVA_HOME to JDK 17 and try again.
        pause
        exit /b 1
    )
)

echo Using JAVA_HOME=%JAVA_HOME%
call mvnw.cmd clean package
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

echo Build complete: target\daily-status-report-web-1.0.0.jar
pause
