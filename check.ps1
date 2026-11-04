# Builds the viewer and runs its tests. Use -Offline to skip the tests that start a hidden Excel.
param([switch]$Offline)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
dotnet build "$root\ExcelViewer.Tests\ExcelViewer.Tests.csproj" -c Release --nologo -v:q -nodeReuse:false -p:UseSharedCompilation=false
if ($LASTEXITCODE -ne 0) { "BUILD FAILED"; exit 1 }
$dll = Join-Path $root 'ExcelViewer.Tests\bin\Release\net8.0-windows\ExcelViewer.Tests.dll'
if ($Offline) { dotnet $dll offline } else { dotnet $dll }
exit $LASTEXITCODE
