param(
    [switch]$Desktop,
    [switch]$Startup
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

if (!$Desktop -and !$Startup) {
    $Desktop = $true
}

$launcher = Join-Path $PSScriptRoot "Launch-Assistant.vbs"
if (!(Test-Path -LiteralPath $launcher)) {
    throw "Missing launcher: $launcher"
}

$shell = New-Object -ComObject WScript.Shell

function New-AssistantShortcut {
    param(
        [string]$Folder,
        [string]$Name
    )

    if (!(Test-Path -LiteralPath $Folder)) {
        New-Item -ItemType Directory -Path $Folder -Force | Out-Null
    }

    $path = Join-Path $Folder $Name
    $shortcut = $shell.CreateShortcut($path)
    $shortcut.TargetPath = "$env:WINDIR\System32\wscript.exe"
    $shortcut.Arguments = "`"$launcher`""
    $shortcut.WorkingDirectory = $PSScriptRoot
    $shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,70"
    $shortcut.Description = "Launch the local llama.cpp text formatting tray assistant"
    $shortcut.Save()

    Write-Host "Created shortcut: $path"
}

if ($Desktop) {
    New-AssistantShortcut -Folder ([Environment]::GetFolderPath("DesktopDirectory")) -Name "Local Text Formatting Assistant.lnk"
}

if ($Startup) {
    New-AssistantShortcut -Folder ([Environment]::GetFolderPath("Startup")) -Name "Local Text Formatting Assistant.lnk"
}

Write-Host ""
Write-Host "Done. Shortcuts launch the assistant hidden; use the tray icon to exit."
