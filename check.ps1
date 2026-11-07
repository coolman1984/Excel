# Builds the viewer and the editor and runs both test projects. Use -Offline to skip the tests that start a hidden Excel.
param([switch]$Offline)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$failed = $false
foreach ($name in 'ExcelViewer.Tests', 'ExcelApp.Tests') {
    dotnet build "$root\$name\$name.csproj" -c Release --nologo -v:q -nodeReuse:false -p:UseSharedCompilation=false
    if ($LASTEXITCODE -ne 0) { "BUILD FAILED: $name"; exit 1 }
}
foreach ($name in 'ExcelViewer.Tests', 'ExcelApp.Tests') {
    $dll = Join-Path $root "$name\bin\Release\net8.0-windows\$name.dll"
    if ($Offline) { dotnet $dll offline } else { dotnet $dll }
    if ($LASTEXITCODE -ne 0) { "TESTS FAILED: $name"; $failed = $true }
}
if ($failed) { exit 1 }
exit 0
