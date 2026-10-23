@echo off
setlocal
title Excel macro installer

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "SCRIPT=%~dp0setup_excel.ps1"

if not exist "%PS%" (
  echo Windows PowerShell was not found.
  echo Please send a screenshot of this window.
  pause
  exit /b 1
)

if not exist "%SCRIPT%" (
  echo setup_excel.ps1 was not found.
  echo Extract the complete ZIP archive and run this file from the tools folder.
  pause
  exit /b 1
)

"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
set "RESULT=%ERRORLEVEL%"

echo.
if not "%RESULT%"=="0" (
  echo Installation failed or was interrupted.
  echo Do not change anything manually. Send a screenshot of this window.
) else (
  echo Installer process finished. Read the PowerShell message above.
)
pause
exit /b %RESULT%
