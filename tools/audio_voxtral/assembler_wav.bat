@echo off
setlocal

REM Dossier de travail = dossier où se trouve ce .bat
set "WORKDIR=%~dp0"
if "%WORKDIR:~-1%"=="\" set "WORKDIR=%WORKDIR:~0,-1%"

REM Script PowerShell portable, voisin de ce .bat
set "PSSCRIPT=%~dp0assemblage.ps1"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PSSCRIPT%" -WorkDir "%WORKDIR%"

if errorlevel 1 (
  echo.
  echo ECHEC - voir le message d'erreur ci-dessus.
  pause
  exit /b 1
)

echo.
echo OK - Assemblage termine.
pause
endlocal
