@echo off
title Informe de PC - diagnostico completo
rem --- Este lanzador se auto-eleva a Administrador y abre la interfaz grafica ---
net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Minimized -File "%~dp0INFORME_PC.ps1"
if errorlevel 1 (
    echo.
    echo Ocurrio un error al ejecutar la aplicacion. Detalle arriba.
    pause
)
