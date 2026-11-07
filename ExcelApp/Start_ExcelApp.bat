@echo off
rem One click: starts ExcelApp, builds it first only if it is missing, and closes this window.
set "APP=%~dp0app\ExcelApp.exe"
if not exist "%APP%" (
  echo Building ExcelApp for the first time. This takes a few minutes...
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1"
  if errorlevel 1 (
    echo The build failed.
    pause
    exit /b 1
  )
)
start "" "%APP%" %*
exit
