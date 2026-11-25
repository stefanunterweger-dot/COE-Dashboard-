@echo off
rem Doppelklick: baut die coe-cockpit.sppkg
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0build-sppkg.ps1"
echo.
pause
