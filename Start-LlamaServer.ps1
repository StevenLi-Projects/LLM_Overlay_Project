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

function Get-PropertyValue {
    param(
        [object]$Object,
        [string]$Name,
        [object]$Default
    )
    if ($Object -and ($Object.PSObject.Properties.Name -contains $Name)) {
        return $Object.$Name
    }
    return $Default
}

function Set-ObjectProperty {
    param(
        [object]$Object,
        [string]$Name,
        [object]$Value
    )
    if ($Object.PSObject.Properties.Name -contains $Name) {
        $Object.$Name = $Value
    } else {
        $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value
    }
}

function Apply-LlamaProfile {
    param([object]$Config)

    if (!$Config.llama -or !($Config.llama.PSObject.Properties.Name -contains "profiles")) { return }
    $profileName = [string](Get-PropertyValue -Object $Config.llama -Name "active_profile" -Default "normal")
    if (!($Config.llama.profiles.PSObject.Properties.Name -contains $profileName)) {
        if ($Config.llama.profiles.PSObject.Properties.Name -contains "normal") {
            $profileName = "normal"
        } else {
            $profileName = ($Config.llama.profiles.PSObject.Properties | Select-Object -First 1).Name
        }
        Set-ObjectProperty -Object $Config.llama -Name "active_profile" -Value $profileName
    }

    $profile = $Config.llama.profiles.$profileName
    foreach ($name in @("model_path", "model_name", "context_size", "gpu_layers", "server_args")) {
        if ($profile.PSObject.Properties.Name -contains $name) {
            Set-ObjectProperty -Object $Config.llama -Name $name -Value $profile.$name
        }
    }
}

function Get-LlamaServerArgs {
    param([object]$Config)

    $args = @(
        "--model", $Config.llama.model_path,
        "--host", $Config.llama.host,
        "--port", ([string]$Config.llama.port),
        "--ctx-size", ([string]$Config.llama.context_size),
        "--n-gpu-layers", ([string]$Config.llama.gpu_layers)
    )

    $extraArgs = Get-PropertyValue -Object $Config.llama -Name "server_args" -Default @()
    foreach ($arg in $extraArgs) {
        if (![string]::IsNullOrWhiteSpace([string]$arg)) {
            $args += [string]$arg
        }
    }

    return $args
}

if (!(Test-Path -LiteralPath $ConfigPath)) {
    throw "Config file not found: $ConfigPath. Copy config.example.json to config.json first."
}

$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$config.llama.cpp_dir = Resolve-LocalPath $config.llama.cpp_dir
if ($config.llama.PSObject.Properties.Name -contains "model_path") {
    $config.llama.model_path = Resolve-LocalPath $config.llama.model_path
}
if ($config.llama.PSObject.Properties.Name -contains "profiles") {
    foreach ($profile in $config.llama.profiles.PSObject.Properties) {
        if ($profile.Value.PSObject.Properties.Name -contains "model_path") {
            $profile.Value.model_path = Resolve-LocalPath $profile.Value.model_path
        }
    }
    Apply-LlamaProfile -Config $config
}

$exe = Join-Path $config.llama.cpp_dir "llama-server.exe"
if (!(Test-Path -LiteralPath $exe)) { throw "llama-server.exe not found: $exe" }
if (!(Test-Path -LiteralPath $config.llama.model_path)) { throw "Model not found: $($config.llama.model_path)" }

$zone = Get-Item -LiteralPath $exe -Stream Zone.Identifier -ErrorAction SilentlyContinue
if ($zone) {
    Write-Warning "Windows has marked llama-server.exe as downloaded from the internet. If launch is canceled, run .\Unblock-LlamaCpp.ps1 once."
}

$args = Get-LlamaServerArgs -Config $config

Write-Host "Starting llama.cpp server:"
Write-Host "  $exe $($args -join ' ')"
Write-Host ""
Write-Host "Keep this window open while using the assistant."
& $exe @args
