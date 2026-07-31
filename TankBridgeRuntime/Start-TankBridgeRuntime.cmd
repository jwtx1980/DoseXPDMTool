@echo off
setlocal
set SCRIPT=%~dp0Start-TankBridgeRuntime.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
endlocal
