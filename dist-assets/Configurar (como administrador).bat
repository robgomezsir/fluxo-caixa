@echo off
chcp 65001 >nul
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"\"%~dp0configurar.ps1\"\" -AllUsers'"
if exist ".configurado" del ".configurado"
echo.
echo Se concluiu sem erro, use: Abrir Fluxo de Caixa.bat
pause
