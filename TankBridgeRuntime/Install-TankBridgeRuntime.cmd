@echo off
setlocal
set SCRIPT=%~dp0Install-TankBridgeRuntime.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process powershell.exe -Verb RunAs -Wait -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%SCRIPT%""'"
endlocal
