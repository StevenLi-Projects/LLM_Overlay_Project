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

function Resolve-SpeculativePaths {
    param([object]$Speculative)

    if ($Speculative -and ($Speculative.PSObject.Properties.Name -contains "draft_model_path")) {
        $Speculative.draft_model_path = Resolve-LocalPath $Speculative.draft_model_path
    }
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

function Get-PropertyDouble {
    param(
        [object]$Object,
        [string]$Name,
        [double]$Default = 0.0
    )
    $value = Get-PropertyValue -Object $Object -Name $Name -Default $null
    if ($null -eq $value) { return $Default }
    try {
        return [double]$value
    } catch {
        return $Default
    }
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
    foreach ($name in @("model_path", "model_name", "context_size", "gpu_layers", "server_args", "speculative")) {
        if ($profile.PSObject.Properties.Name -contains $name) {
            Set-ObjectProperty -Object $Config.llama -Name $name -Value $profile.$name
        }
    }
    if (!($profile.PSObject.Properties.Name -contains "speculative") -and
        !($Config.llama.PSObject.Properties.Name -contains "speculative")) {
        Set-ObjectProperty -Object $Config.llama -Name "speculative" -Value ([pscustomobject]@{ enabled = $false })
    }
}

function Get-ConfiguredDraftModelPath {
    param([object]$Config)

    $speculative = Get-PropertyValue -Object $Config.llama -Name "speculative" -Default $null
    if ($speculative -and (Get-ConfigBool -Object $speculative -Name "enabled" -Default $false)) {
        return [string](Get-PropertyValue -Object $speculative -Name "draft_model_path" -Default "")
    }
    return ""
}

function Test-LlamaServerHelpFlag {
    param(
        [object]$Config,
        [string]$Flag
    )

    $exe = Join-Path $Config.llama.cpp_dir "llama-server.exe"
    if (!(Test-Path -LiteralPath $exe)) { return $false }

    $previousErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $help = & $exe --help 2>&1 | Out-String
        return ($help -match [regex]::Escape($Flag))
    } catch {
        return $false
    } finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
}

function Get-LlamaServerArgs {
    param([object]$Config)

    $configuredGpuLayers = Get-ConfigInt -Object $Config.llama -Name "gpu_layers" -Default 0
    $preferGpu = Get-ConfigBool -Object $Config.llama -Name "prefer_gpu" -Default ($configuredGpuLayers -gt 0)
    $requireGpu = Get-ConfigBool -Object $Config.llama -Name "require_gpu" -Default $false
    $gpuAvailable = $false
    if ($preferGpu -or $requireGpu) {
        $gpuAvailable = Test-LlamaGpuAvailable -Config $Config
    }
    if ($requireGpu -and !$gpuAvailable) {
        throw "GPU is required, but this llama.cpp build currently reports no GPU devices. Run '.\Diagnose-LlamaGpu.ps1'. Set llama.require_gpu to false to allow CPU fallback."
    }

    $effectiveGpuLayers = 0
    if ($preferGpu -and $gpuAvailable) {
        $effectiveGpuLayers = $configuredGpuLayers
    }

    $args = @(
        "--model", $Config.llama.model_path,
        "--alias", $Config.llama.model_name,
        "--host", $Config.llama.host,
        "--port", ([string]$Config.llama.port),
        "--ctx-size", ([string]$Config.llama.context_size),
        "--n-gpu-layers", ([string]$effectiveGpuLayers)
    )

    if ($preferGpu -and $gpuAvailable -and $effectiveGpuLayers -gt 0) {
        $gpuDevice = [string](Get-PropertyValue -Object $Config.llama -Name "gpu_device" -Default "")
        if (![string]::IsNullOrWhiteSpace($gpuDevice)) {
            $args += @("--device", $gpuDevice)
        }
    } elseif ($preferGpu -and !$gpuAvailable) {
        Write-Warning "No llama.cpp GPU device detected; starting with CPU fallback."
    }

    $speculative = Get-PropertyValue -Object $Config.llama -Name "speculative" -Default $null
    if ($speculative -and (Get-ConfigBool -Object $speculative -Name "enabled" -Default $false)) {
        $draftModelPath = [string](Get-PropertyValue -Object $speculative -Name "draft_model_path" -Default "")
        if ([string]::IsNullOrWhiteSpace($draftModelPath)) {
            throw "Speculative decoding is enabled, but llama.speculative.draft_model_path is empty."
        }
        if (!(Test-Path -LiteralPath $draftModelPath)) {
            throw "Speculative draft model not found: $draftModelPath"
        }

        $specType = [string](Get-PropertyValue -Object $speculative -Name "type" -Default "gemma4_mtp")
        if ($specType -eq "gemma4_mtp") {
            $supportsAtomicMtp = Test-LlamaServerHelpFlag -Config $Config -Flag "--mtp-head"
            $supportsMainlineMtp = (Test-LlamaServerHelpFlag -Config $Config -Flag "draft-mtp") -and
                (Test-LlamaServerHelpFlag -Config $Config -Flag "--spec-draft-model")
            if (!($supportsAtomicMtp -or $supportsMainlineMtp)) {
                $fallback = Get-ConfigBool -Object $speculative -Name "fallback_without_support" -Default $true
                if ($fallback) {
                    Write-Warning "Gemma 4 MTP is configured, but this llama.cpp build exposes neither draft-mtp nor --mtp-head. Starting without speculative decoding."
                    return $args
                }
                throw "Gemma 4 MTP needs a llama.cpp build with draft-mtp or --mtp-head support."
            }

            if ($supportsMainlineMtp) {
                $args += @("--spec-draft-model", $draftModelPath, "--spec-type", "draft-mtp")
                $draftNMax = Get-ConfigInt -Object $speculative -Name "draft_n_max" -Default 4
                $draftNMin = Get-ConfigInt -Object $speculative -Name "draft_n_min" -Default 1
                $draftPMin = Get-PropertyDouble -Object $speculative -Name "draft_p_min" -Default 0.75
                $args += @("--spec-draft-n-max", ([string]$draftNMax), "--spec-draft-n-min", ([string]$draftNMin), "--spec-draft-p-min", ([string]$draftPMin))
            } else {
                $args += @("--mtp-head", $draftModelPath, "--spec-type", "mtp")
                $draftBlockSize = Get-ConfigInt -Object $speculative -Name "draft_block_size" -Default 2
                if ($draftBlockSize -gt 0) {
                    $args += @("--draft-block-size", ([string]$draftBlockSize))
                }
            }

            $draftGpuLayers = Get-ConfigInt -Object $speculative -Name "draft_gpu_layers" -Default $effectiveGpuLayers
            if (!($preferGpu -and $gpuAvailable)) {
                $draftGpuLayers = 0
            }
            $args += @("--n-gpu-layers-draft", ([string]$draftGpuLayers))

            if ($preferGpu -and $gpuAvailable -and $draftGpuLayers -gt 0) {
                $draftDevice = [string](Get-PropertyValue -Object $speculative -Name "draft_gpu_device" -Default (Get-PropertyValue -Object $Config.llama -Name "gpu_device" -Default ""))
                if (![string]::IsNullOrWhiteSpace($draftDevice)) {
                    $args += @($(if ($supportsMainlineMtp) { "--spec-draft-device" } else { "--device-draft" }), $draftDevice)
                }
            }
        } else {
            $draftGpuLayers = Get-ConfigInt -Object $speculative -Name "draft_gpu_layers" -Default $effectiveGpuLayers
            if (!($preferGpu -and $gpuAvailable)) {
                $draftGpuLayers = 0
            }

            $args += @("--model-draft", $draftModelPath)
            $args += @("--n-gpu-layers-draft", ([string]$draftGpuLayers))

            if ($preferGpu -and $gpuAvailable -and $draftGpuLayers -gt 0) {
                $draftDevice = [string](Get-PropertyValue -Object $speculative -Name "draft_gpu_device" -Default (Get-PropertyValue -Object $Config.llama -Name "gpu_device" -Default ""))
                if (![string]::IsNullOrWhiteSpace($draftDevice)) {
                    $args += @("--device-draft", $draftDevice)
                }
            }

            $draftContextSize = Get-ConfigInt -Object $speculative -Name "draft_context_size" -Default 0
            if ($draftContextSize -gt 0) {
                $args += @("--ctx-size-draft", ([string]$draftContextSize))
            }

            $draftNMax = Get-ConfigInt -Object $speculative -Name "draft_n_max" -Default 0
            if ($draftNMax -gt 0) {
                $args += @("--spec-draft-n-max", ([string]$draftNMax))
            }

            $draftNMin = Get-ConfigInt -Object $speculative -Name "draft_n_min" -Default 0
            if ($draftNMin -gt 0) {
                $args += @("--spec-draft-n-min", ([string]$draftNMin))
            }

            $draftPMin = Get-PropertyDouble -Object $speculative -Name "draft_p_min" -Default -1.0
            if ($draftPMin -ge 0.0) {
                $args += @("--spec-draft-p-min", ([string]$draftPMin))
            }
        }
    }

    $extraArgs = Get-PropertyValue -Object $Config.llama -Name "server_args" -Default @()
    foreach ($arg in $extraArgs) {
        if (![string]::IsNullOrWhiteSpace([string]$arg)) {
            $args += [string]$arg
        }
    }

    return $args
}

function Get-ConfigBool {
    param(
        [object]$Object,
        [string]$Name,
        [bool]$Default
    )
    if ($Object -and ($Object.PSObject.Properties.Name -contains $Name)) {
        return [bool]$Object.$Name
    }
    return $Default
}

function Get-ConfigInt {
    param(
        [object]$Object,
        [string]$Name,
        [int]$Default
    )
    if ($Object -and ($Object.PSObject.Properties.Name -contains $Name)) {
        return [int]$Object.$Name
    }
    return $Default
}

function Test-LlamaGpuAvailable {
    param([object]$Config)

    $exe = Join-Path $Config.llama.cpp_dir "llama-server.exe"
    if (!(Test-Path -LiteralPath $exe)) { return $false }

    $previousErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $output = & $exe --list-devices 2>&1 | Out-String
        return ($output -match "(?im)^\s*(?:(?:CUDA|Vulkan|SYCL|Metal)\d*\s*:|Device\s+\d+:.*(?:CUDA|NVIDIA|GeForce|RTX|Vulkan|SYCL|Metal))")
    } catch {
        Write-Warning "Could not query llama.cpp devices: $($_.Exception.Message)"
        return $false
    } finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
}

function Assert-LlamaGpuAvailable {
    param([object]$Config)

    $gpuLayers = Get-ConfigInt -Object $Config.llama -Name "gpu_layers" -Default 0
    $requireGpu = Get-ConfigBool -Object $Config.llama -Name "require_gpu" -Default $false
    if (!$requireGpu -or $gpuLayers -le 0) { return }

    if (!(Test-LlamaGpuAvailable -Config $Config)) {
        throw "GPU is required, but this llama.cpp build currently reports no GPU devices. Run '.\Diagnose-LlamaGpu.ps1'. Set llama.require_gpu to false to allow CPU fallback."
    }
}

if (!(Test-Path -LiteralPath $ConfigPath)) {
    throw "Config file not found: $ConfigPath. Copy config.example.json to config.json first."
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
    Apply-LlamaProfile -Config $config
}

$exe = Join-Path $config.llama.cpp_dir "llama-server.exe"
if (!(Test-Path -LiteralPath $exe)) { throw "llama-server.exe not found: $exe" }
if (!(Test-Path -LiteralPath $config.llama.model_path)) { throw "Model not found: $($config.llama.model_path)" }
$draftModelPath = Get-ConfiguredDraftModelPath -Config $config
if (![string]::IsNullOrWhiteSpace($draftModelPath) -and !(Test-Path -LiteralPath $draftModelPath)) {
    throw "Speculative draft model not found: $draftModelPath"
}
Assert-LlamaGpuAvailable -Config $config

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
