# Builds ExcelApp: the interface (pnpm + vite), then the host (dotnet publish), then copies the interface next to it.
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$ui = Join-Path $here 'ui'
$app = Join-Path $here 'app'
Push-Location $ui
try {
    pnpm install --frozen-lockfile
    if ($LASTEXITCODE -ne 0) { throw 'pnpm install failed' }
    pnpm exec vite build
    if ($LASTEXITCODE -ne 0) { throw 'vite build failed' }
} finally { Pop-Location }
dotnet publish (Join-Path $here 'ExcelApp.csproj') -c Release -o $app --no-self-contained --nologo -v:q -nodeReuse:false
if ($LASTEXITCODE -ne 0) { throw 'dotnet publish failed' }
$target = Join-Path $app 'ui'
if (Test-Path $target) { Remove-Item $target -Recurse -Force }
New-Item -ItemType Directory -Force $target | Out-Null
Copy-Item (Join-Path $ui 'dist\*') $target -Recurse -Force
"Built: $app\ExcelApp.exe"
