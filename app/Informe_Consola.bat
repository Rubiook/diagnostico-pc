@echo off
title Informe de PC - modo consola (sin ventana)
rem --- Version sin interfaz: corre todo, imprime el resumen y lo copia al portapapeles ---
net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0INFORME_PC.ps1" -Auto -MuestreoSegundos 45 -TestDisco
echo.
echo ================================================================
echo  Listo. El resumen de arriba tambien quedo en el portapapeles.
echo ================================================================
pause
