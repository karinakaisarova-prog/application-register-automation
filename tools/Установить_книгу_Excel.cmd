@echo off
setlocal
chcp 65001 >nul
title Установка макросов Excel

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "SCRIPT=%~dp0setup_excel.ps1"

if not exist "%PS%" (
  echo Не удалось найти Windows PowerShell.
  echo Пришлите скриншот этого окна.
  pause
  exit /b 1
)

if not exist "%SCRIPT%" (
  echo Не найден файл setup_excel.ps1.
  echo Убедитесь, что ZIP-архив полностью распакован и этот CMD-файл находится в папке tools.
  pause
  exit /b 1
)

"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
set "RESULT=%ERRORLEVEL%"

echo.
if not "%RESULT%"=="0" (
  echo Установка не завершена. Ничего не исправляйте вручную.
  echo Сделайте скриншот этого окна и пришлите его в чат.
) else (
  echo Работа установщика завершена.
)
pause
exit /b %RESULT%
