@echo off
rem One click: starts the viewer, builds it first only if it is missing, and closes this window.
set "APP=%~dp0app\ExcelViewer.exe"
if not exist "%APP%" (
  echo Building Excel Viewer for the first time. This takes under a minute...
  dotnet publish "%~dp0ExcelViewer.csproj" -c Release -o "%~dp0app" --nologo -v:q -nodeReuse:false
  if errorlevel 1 (
    echo The build failed.
    pause
    exit /b 1
  )
)
start "" "%APP%" %*
exit
