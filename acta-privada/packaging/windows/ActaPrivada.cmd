@echo off
title Actas Privadas - no cierre esta ventana mientras usa la aplicacion
cd /d "%~dp0"
"%~dp0python\python.exe" "%~dp0launcher.py"
if errorlevel 1 pause
