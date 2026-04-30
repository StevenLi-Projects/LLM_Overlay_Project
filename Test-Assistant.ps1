param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [switch]$RequireServer
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
    throw "Missing config file: $ConfigPath"
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
    $activeProfile = "normal"
    if ($config.llama.PSObject.Properties.Name -contains "active_profile") {
        $activeProfile = [string]$config.llama.active_profile
    }
    if ($config.llama.profiles.PSObject.Properties.Name -contains $activeProfile) {
        $config.llama.model_path = $config.llama.profiles.$activeProfile.model_path
    }
}

$serverExe = Join-Path $config.llama.cpp_dir "llama-server.exe"
$zone = Get-Item -LiteralPath $serverExe -Stream Zone.Identifier -ErrorAction SilentlyContinue
$modeTokenSettingsValid = $true
foreach ($mode in $config.modes.PSObject.Properties) {
    if ($mode.Value.PSObject.Properties.Name -contains "max_tokens") {
        $modeTokenSettingsValid = $modeTokenSettingsValid -and ([int]$mode.Value.max_tokens -gt 0)
    }
}
$serverArgsValid = $true
if ($config.llama.PSObject.Properties.Name -contains "server_args") {
    $serverArgsValid = ($null -ne $config.llama.server_args)
}
$profilesValid = $true
if ($config.llama.PSObject.Properties.Name -contains "profiles") {
    foreach ($profile in $config.llama.profiles.PSObject.Properties) {
        $profilesValid = $profilesValid -and
            ($profile.Value.PSObject.Properties.Name -contains "model_path") -and
            (Test-Path -LiteralPath $profile.Value.model_path)
    }
}
$checks = @(
    @{ Name = "config.json parses"; Passed = $true; Detail = $ConfigPath },
    @{ Name = "llama-server.exe exists"; Passed = (Test-Path -LiteralPath $serverExe); Detail = $serverExe },
    @{ Name = "GGUF model exists"; Passed = (Test-Path -LiteralPath $config.llama.model_path); Detail = $config.llama.model_path },
    @{ Name = "server URL is configured"; Passed = ($config.llama.server_url -match "^https?://"); Detail = $config.llama.server_url },
    @{ Name = "completion preference setting exists"; Passed = ($config.generation.PSObject.Properties.Name -contains "prefer_completion"); Detail = "generation.prefer_completion" },
    @{ Name = "preview setting exists"; Passed = ($config.ui.PSObject.Properties.Name -contains "preview_enabled"); Detail = "ui.preview_enabled" },
    @{ Name = "notification setting exists"; Passed = ($config.ui.PSObject.Properties.Name -contains "show_notifications"); Detail = "ui.show_notifications" },
    @{ Name = "timing notification setting exists"; Passed = ($config.ui.PSObject.Properties.Name -contains "show_timing_notifications"); Detail = "ui.show_timing_notifications" },
    @{ Name = "health cache setting exists"; Passed = ($config.llama.PSObject.Properties.Name -contains "health_cache_sec"); Detail = "llama.health_cache_sec" },
    @{ Name = "server args setting is valid"; Passed = $serverArgsValid; Detail = "llama.server_args" },
    @{ Name = "model profiles are valid"; Passed = $profilesValid; Detail = "llama.profiles.*.model_path" },
    @{ Name = "mode max token settings are valid"; Passed = $modeTokenSettingsValid; Detail = "modes.*.max_tokens" },
    @{ Name = "at least one mode enabled"; Passed = (($config.modes.PSObject.Properties | Where-Object { $_.Value.enabled }).Count -gt 0); Detail = "" }
)

if ($RequireServer) {
    $serverOk = $false
    try {
        Invoke-RestMethod -Method Get -Uri "$($config.llama.server_url.TrimEnd('/'))/health" -TimeoutSec 2 | Out-Null
        $serverOk = $true
    } catch {
        try {
            Invoke-RestMethod -Method Get -Uri "$($config.llama.server_url.TrimEnd('/'))/v1/models" -TimeoutSec 2 | Out-Null
            $serverOk = $true
        } catch {
            $serverOk = $false
        }
    }
    $checks += @{ Name = "llama.cpp server reachable"; Passed = $serverOk; Detail = $config.llama.server_url }
}

$failed = $false
foreach ($check in $checks) {
    if ($check.Passed) {
        Write-Host "[OK]   $($check.Name) $($check.Detail)"
    } else {
        Write-Host "[FAIL] $($check.Name) $($check.Detail)"
        $failed = $true
    }
}

if ($failed) { exit 1 }
if ($zone) {
    Write-Host ""
    Write-Host "[WARN] llama-server.exe has a Zone.Identifier download mark. If Windows cancels launch, run .\Unblock-LlamaCpp.ps1 once."
}
Write-Host ""
Write-Host "Validation completed."
