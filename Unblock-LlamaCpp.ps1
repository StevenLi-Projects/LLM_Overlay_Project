param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json")
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

function Resolve-LocalPath {
    param([string]$PathValue)
    if ([string]::IsNullOrWhiteSpace($PathValue)) { return $PathValue }
    if ([System.IO.Path]::IsPathRooted($PathValue)) { return $PathValue }
    return (Join-Path $PSScriptRoot $PathValue)
}

if (!(Test-Path -LiteralPath $ConfigPath)) {
    throw "Config file not found: $ConfigPath"
}

$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$cppDir = Resolve-LocalPath $config.llama.cpp_dir
if (!(Test-Path -LiteralPath $cppDir)) {
    throw "llama.cpp directory not found: $cppDir"
}

Get-ChildItem -LiteralPath $cppDir -File | ForEach-Object {
    Unblock-File -LiteralPath $_.FullName
}

Write-Host "Removed Windows download blocking marks from files in:"
Write-Host "  $cppDir"
