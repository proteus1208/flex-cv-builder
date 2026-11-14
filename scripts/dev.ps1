#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$Root = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $Root

if ($env:DOTNET_ROOT) {
    $env:PATH = "$env:DOTNET_ROOT;$env:PATH"
}

Write-Host "Restoring and launching CV Generator (Debug)..."
dotnet restore .\CvGenerator.sln
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

dotnet run --project .\src\CvGenerator.App\CvGenerator.App.csproj -c Debug
exit $LASTEXITCODE
