#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$Root = Resolve-Path (Join-Path $PSScriptRoot "..")
$Out = Join-Path $Root "artifacts\CvGenerator"
Set-Location $Root

if ($env:DOTNET_ROOT) {
    $env:PATH = "$env:DOTNET_ROOT;$env:PATH"
}

dotnet restore .\CvGenerator.sln
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

dotnet test .\tests\CvGenerator.Core.Tests\CvGenerator.Core.Tests.csproj -c Release --nologo
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

dotnet test .\tests\CvGenerator.Docx.Tests\CvGenerator.Docx.Tests.csproj -c Release --nologo
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

dotnet publish .\src\CvGenerator.App\CvGenerator.App.csproj `
  -c Release `
  -r win-x64 `
  --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:EnableCompressionInSingleFile=true `
  -o $Out
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

New-Item -ItemType Directory -Force -Path (Join-Path $Out "Templates") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $Out "Output") | Out-Null

$templates = Join-Path $Root "Templates\*.docx"
if (Test-Path $templates) {
    Copy-Item $templates (Join-Path $Out "Templates") -Force
}

Write-Host "Published self-contained single-file win-x64 to $Out"
Write-Host "Copy the CvGenerator folder to other Windows PCs (no .NET install required)."
Write-Host "Requires: WebView2 Evergreen + LibreOffice for PDF."
