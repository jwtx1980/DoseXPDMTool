@echo off
set SCRIPT=%~dp0Allow-TankBridgeRuntimeFirewall.ps1
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%SCRIPT%""'"
