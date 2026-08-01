@echo off
setlocal
chcp 65001 >nul
title Actualisation des classeurs Excel
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Actualiser_Excel.ps1" -Dossier "%~dp0."
set EXIT_CODE=%ERRORLEVEL%
if not "%EXIT_CODE%"=="0" (
  echo.
  echo L'actualisation a rencontre une erreur. Consulte Actualisation.log.
  pause
)
exit /b %EXIT_CODE%
