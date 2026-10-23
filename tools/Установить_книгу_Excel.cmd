@echo off
chcp 65001 >nul
set "SCRIPT=%~dp0setup_excel.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
if errorlevel 1 (
  echo.
  echo Установщик завершился с ошибкой. Прочитайте сообщение выше.
  pause
)
