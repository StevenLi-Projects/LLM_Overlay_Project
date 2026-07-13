param(
    [string]$DotNetPath = "dotnet"
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$project = Join-Path $PSScriptRoot "src\LocalTextFormattingAssistant\LocalTextFormattingAssistant.csproj"
$output = Join-Path $PSScriptRoot "dist"
$nugetConfig = Join-Path $PSScriptRoot "NuGet.Config"

if (!(Test-Path -LiteralPath $project)) {
    throw "Project not found: $project"
}

$sdks = & $DotNetPath --list-sdks
if ($LASTEXITCODE -ne 0 -or !$sdks) {
    throw ".NET 8 SDK is required to rebuild the app. The published app only requires the .NET 8 Desktop Runtime."
}

$env:DOTNET_CLI_TELEMETRY_OPTOUT = "1"
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = "1"

& $DotNetPath restore $project --configfile $nugetConfig
if ($LASTEXITCODE -ne 0) { throw "dotnet restore failed." }

& $DotNetPath publish $project `
    -c Release `
    --self-contained false `
    --no-restore `
    -p:DebugType=None `
    -p:DebugSymbols=false `
    -o $output
if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed." }

Write-Host "Published: $(Join-Path $output 'LocalTextFormattingAssistant.exe')"
