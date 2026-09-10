@echo off
setlocal
cd /d "%~dp0"
start "" "%~dp0docs\architecture-viewer.html"
echo Opening architecture diagrams in your default browser...
echo If nothing opens, double-click docs\architecture-viewer.html manually.
endlocal
