# Builds the viewer and the editor and runs both test projects. Use -Offline to skip the tests that start a hidden Excel.
# -Suites store,ingest runs only those ExcelApp suites (and skips the viewer tests): plan 4 lanes use it, because the
# older suites (style, host) touch the approved cache folder that lanes must never open.
# -m:1 keeps MSBuild in one process: worker nodes cannot start inside the Codex sandbox (build fails with 0 errors).
param([switch]$Offline, [string]$Suites)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$failed = $false
$names = if ($Suites) { @('ExcelApp.Tests') } else { @('ExcelViewer.Tests', 'ExcelApp.Tests') }
foreach ($name in $names) {
    dotnet build "$root\$name\$name.csproj" -c Release --nologo -v:q -m:1 -nodeReuse:false -p:UseSharedCompilation=false
    if ($LASTEXITCODE -ne 0) { "BUILD FAILED: $name"; exit 1 }
}
foreach ($name in $names) {
    $dll = Join-Path $root "$name\bin\Release\net8.0-windows\$name.dll"
    $testArgs = @()
    if ($Offline) { $testArgs += 'offline' }
    if ($Suites -and $name -eq 'ExcelApp.Tests') { $testArgs += "--suites=$Suites" }
    dotnet $dll @testArgs
    if ($LASTEXITCODE -ne 0) { "TESTS FAILED: $name"; $failed = $true }
}
if ($failed) { exit 1 }
exit 0
