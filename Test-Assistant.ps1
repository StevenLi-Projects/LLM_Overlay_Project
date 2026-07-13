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

function Get-ConfigBool {
    param(
        [object]$Object,
        [string]$Name,
        [bool]$Default
    )
    return [bool](Get-PropertyValue -Object $Object -Name $Name -Default $Default)
}

function Resolve-SpeculativePaths {
    param([object]$Speculative)

    if ($Speculative -and ($Speculative.PSObject.Properties.Name -contains "draft_model_path")) {
        $Speculative.draft_model_path = Resolve-LocalPath $Speculative.draft_model_path
    }
}

if (!(Test-Path -LiteralPath $ConfigPath)) {
    throw "Missing config file: $ConfigPath"
}

$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$config.llama.cpp_dir = Resolve-LocalPath $config.llama.cpp_dir
if ($config.llama.PSObject.Properties.Name -contains "model_path") {
    $config.llama.model_path = Resolve-LocalPath $config.llama.model_path
}
if ($config.llama.PSObject.Properties.Name -contains "speculative") {
    Resolve-SpeculativePaths -Speculative $config.llama.speculative
}
if ($config.llama.PSObject.Properties.Name -contains "profiles") {
    foreach ($profile in $config.llama.profiles.PSObject.Properties) {
        if ($profile.Value.PSObject.Properties.Name -contains "model_path") {
            $profile.Value.model_path = Resolve-LocalPath $profile.Value.model_path
        }
        if ($profile.Value.PSObject.Properties.Name -contains "speculative") {
            Resolve-SpeculativePaths -Speculative $profile.Value.speculative
        }
    }
    $activeProfile = "normal"
    if ($config.llama.PSObject.Properties.Name -contains "active_profile") {
        $activeProfile = [string]$config.llama.active_profile
    }
    if ($config.llama.profiles.PSObject.Properties.Name -contains $activeProfile) {
        $config.llama.model_path = $config.llama.profiles.$activeProfile.model_path
        if ($config.llama.profiles.$activeProfile.PSObject.Properties.Name -contains "speculative") {
            $config.llama | Add-Member -NotePropertyName "speculative" -NotePropertyValue $config.llama.profiles.$activeProfile.speculative -Force
        }
    }
}

$serverExe = Join-Path $config.llama.cpp_dir "llama-server.exe"
$compiledApp = Join-Path $PSScriptRoot "dist\LocalTextFormattingAssistant.exe"
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
$requireGpu = $false
if ($config.llama.PSObject.Properties.Name -contains "require_gpu") {
    $requireGpu = [bool]$config.llama.require_gpu
}
$preferGpu = $false
if ($config.llama.PSObject.Properties.Name -contains "prefer_gpu") {
    $preferGpu = [bool]$config.llama.prefer_gpu
}
$gpuAvailable = $true
$gpuDetail = "GPU not required"
if ($requireGpu -or $preferGpu) {
    try {
        $previousErrorActionPreference = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        $deviceOutput = & $serverExe --list-devices 2>&1 | Out-String
        $gpuAvailable = ($deviceOutput -match "(?im)^\s*(?:(?:CUDA|Vulkan|SYCL|Metal)\d*\s*:|Device\s+\d+:.*(?:CUDA|NVIDIA|GeForce|RTX|Vulkan|SYCL|Metal))")
        $gpuDetail = "llama-server --list-devices"
    } catch {
        $gpuAvailable = $false
        $gpuDetail = $_.Exception.Message
    } finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
}
$gpuPolicyValid = (!$requireGpu -or $gpuAvailable)
$mtpRuntime = "unavailable"
if (Test-Path -LiteralPath $serverExe) {
    try {
        $previousErrorActionPreference = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        $helpOutput = & $serverExe --help 2>&1 | Out-String
        if (($helpOutput -match "draft-mtp") -and ($helpOutput -match "--spec-draft-model|--model-draft")) {
            $mtpRuntime = "mainline draft-mtp"
        } elseif ($helpOutput -match "--mtp-head") {
            $mtpRuntime = "Atomic mtp-head"
        }
    } catch {
        $mtpRuntime = "probe failed: $($_.Exception.Message)"
    } finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
}
$profilesValid = $true
$activeProfileValid = $true
$profileSwitchingValid = $true
$speculativeValid = $true
$speculativeEnabledProfiles = @()
if ($config.llama.PSObject.Properties.Name -contains "profiles") {
    $profileNames = @($config.llama.profiles.PSObject.Properties | ForEach-Object { $_.Name })
    $activeProfileValid = $profileNames -contains $activeProfile
    $profileSwitchingValid = ($profileNames -contains "normal") -and ($profileNames -contains "fast")

    foreach ($profile in $config.llama.profiles.PSObject.Properties) {
        $profilesValid = $profilesValid -and
            ($profile.Value.PSObject.Properties.Name -contains "model_path") -and
            (Test-Path -LiteralPath $profile.Value.model_path)

        $profileSwitchingValid = $profileSwitchingValid -and
            ($profile.Value.PSObject.Properties.Name -contains "model_name") -and
            ($profile.Value.PSObject.Properties.Name -contains "context_size") -and
            ([int]$profile.Value.context_size -gt 0)

        if ($profile.Value.PSObject.Properties.Name -contains "speculative") {
            $spec = $profile.Value.speculative
            if (Get-ConfigBool -Object $spec -Name "enabled" -Default $false) {
                $speculativeEnabledProfiles += $profile.Name
                $speculativeValid = $speculativeValid -and
                    ($spec.PSObject.Properties.Name -contains "draft_model_path") -and
                    (Test-Path -LiteralPath $spec.draft_model_path) -and
                    (!($spec.PSObject.Properties.Name -contains "draft_context_size") -or [int]$spec.draft_context_size -gt 0) -and
                    (!($spec.PSObject.Properties.Name -contains "draft_gpu_layers") -or [int]$spec.draft_gpu_layers -ge 0) -and
                    (!($spec.PSObject.Properties.Name -contains "draft_n_max") -or [int]$spec.draft_n_max -gt 0) -and
                    (!($spec.PSObject.Properties.Name -contains "draft_n_min") -or [int]$spec.draft_n_min -ge 0)
            }
        }
    }
}
$checks = @(
    @{ Name = "config.json parses"; Passed = $true; Detail = $ConfigPath },
    @{ Name = "llama-server.exe exists"; Passed = (Test-Path -LiteralPath $serverExe); Detail = $serverExe },
    @{ Name = "compiled tray app exists"; Passed = (Test-Path -LiteralPath $compiledApp); Detail = $compiledApp },
    @{ Name = "GGUF model exists"; Passed = (Test-Path -LiteralPath $config.llama.model_path); Detail = $config.llama.model_path },
    @{ Name = "server URL is configured"; Passed = ($config.llama.server_url -match "^https?://"); Detail = $config.llama.server_url },
    @{ Name = "completion preference setting exists"; Passed = ($config.generation.PSObject.Properties.Name -contains "prefer_completion"); Detail = "generation.prefer_completion" },
    @{ Name = "preview setting exists"; Passed = ($config.ui.PSObject.Properties.Name -contains "preview_enabled"); Detail = "ui.preview_enabled" },
    @{ Name = "system theme setting exists"; Passed = ($config.ui.PSObject.Properties.Name -contains "theme"); Detail = "ui.theme" },
    @{ Name = "progress overlay setting exists"; Passed = ($config.ui.PSObject.Properties.Name -contains "progress_overlay_enabled"); Detail = "ui.progress_overlay_enabled" },
    @{ Name = "prewarm setting exists"; Passed = ($config.llama.PSObject.Properties.Name -contains "prewarm_on_launch"); Detail = "llama.prewarm_on_launch" },
    @{ Name = "notification setting exists"; Passed = ($config.ui.PSObject.Properties.Name -contains "show_notifications"); Detail = "ui.show_notifications" },
    @{ Name = "timing notification setting exists"; Passed = ($config.ui.PSObject.Properties.Name -contains "show_timing_notifications"); Detail = "ui.show_timing_notifications" },
    @{ Name = "health cache setting exists"; Passed = ($config.llama.PSObject.Properties.Name -contains "health_cache_sec"); Detail = "llama.health_cache_sec" },
    @{ Name = "server args setting is valid"; Passed = $serverArgsValid; Detail = "llama.server_args" },
    @{ Name = "GPU preference setting exists"; Passed = ($config.llama.PSObject.Properties.Name -contains "prefer_gpu"); Detail = "llama.prefer_gpu" },
    @{ Name = "GPU policy is valid"; Passed = $gpuPolicyValid; Detail = $gpuDetail },
    @{ Name = "preferred GPU device is detected"; Passed = (!$preferGpu -or $gpuAvailable); Detail = $(if ($gpuAvailable) { "accelerator found" } else { "CPU fallback would be used" }) },
    @{ Name = "model profiles are valid"; Passed = $profilesValid; Detail = "llama.profiles.*.model_path" },
    @{ Name = "active profile is valid"; Passed = $activeProfileValid; Detail = "llama.active_profile" },
    @{ Name = "normal/fast profile switching is configured"; Passed = $profileSwitchingValid; Detail = "llama.profiles.normal + llama.profiles.fast" },
    @{ Name = "speculative draft settings are valid"; Passed = $speculativeValid; Detail = "enabled profiles: $($speculativeEnabledProfiles -join ', ')" },
    @{ Name = "Gemma MTP runtime is available"; Passed = (($speculativeEnabledProfiles.Count -eq 0) -or ($mtpRuntime -ne "unavailable")); Detail = $mtpRuntime },
    @{ Name = "mode max token settings are valid"; Passed = $modeTokenSettingsValid; Detail = "modes.*.max_tokens" },
    @{ Name = "at least one mode enabled"; Passed = (($config.modes.PSObject.Properties | Where-Object { $_.Value.enabled }).Count -gt 0); Detail = "" }
)

if (Test-Path -LiteralPath $compiledApp) {
    $validationOutput = & $compiledApp --validate --config $ConfigPath 2>&1 | Out-String
    $validationExitCode = $LASTEXITCODE
    if (![string]::IsNullOrWhiteSpace($validationOutput)) { Write-Host $validationOutput.TrimEnd() }
    $checks += @{ Name = "compiled app validation passes"; Passed = ($validationExitCode -eq 0); Detail = "--validate" }
}

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
