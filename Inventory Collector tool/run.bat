@echo off
cd /d %~dp0

echo Running inventory script...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0inventory.ps1"

echo.
echo Finished. Check the output folder.
pause