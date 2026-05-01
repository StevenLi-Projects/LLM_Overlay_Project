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
    throw "Missing config file: $ConfigPath"
}

$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$config.llama.cpp_dir = Resolve-LocalPath $config.llama.cpp_dir
$serverExe = Join-Path $config.llama.cpp_dir "llama-server.exe"

Write-Host "== NVIDIA driver visibility =="
try {
    nvidia-smi
} catch {
    Write-Host "[FAIL] nvidia-smi is not available or cannot see the GPU."
}

Write-Host ""
Write-Host "== llama.cpp device visibility =="
if (Test-Path -LiteralPath $serverExe) {
    & $serverExe --list-devices
} else {
    Write-Host "[FAIL] llama-server.exe not found: $serverExe"
}

Write-Host ""
Write-Host "== CUDA runtime DLL lookup =="
$dllNames = @(
    "cudart64_12.dll",
    "cublas64_12.dll",
    "cublasLt64_12.dll",
    "nvrtc64_120_0.dll"
)

foreach ($dll in $dllNames) {
    $localPath = Join-Path $config.llama.cpp_dir $dll
    if (Test-Path -LiteralPath $localPath) {
        Write-Host "[OK]   $dll found next to llama-server.exe"
        continue
    }

    $where = $null
    try {
        $where = & where.exe $dll 2>$null
    } catch {
        $where = $null
    }

    if ($where) {
        Write-Host "[OK]   $dll found on PATH: $($where -join '; ')"
    } else {
        Write-Host "[MISS] $dll not found next to llama-server.exe or on PATH"
    }
}

Write-Host ""
Write-Host "If nvidia-smi works but llama.cpp lists no CUDA device, install the official NVIDIA CUDA Toolkit or copy the missing CUDA 12 DLLs from its bin folder into the llama.cpp folder."
