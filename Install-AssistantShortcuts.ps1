param(
    [switch]$Desktop,
    [switch]$Startup
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

if (!$Desktop -and !$Startup) {
    $Desktop = $true
}

$app = Join-Path $PSScriptRoot "dist\LocalTextFormattingAssistant.exe"
if (!(Test-Path -LiteralPath $app)) {
    throw "Compiled app not found: $app. Publish the app before installing shortcuts."
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
    $shortcut.TargetPath = $app
    $shortcut.Arguments = "--config `"$(Join-Path $PSScriptRoot 'config.json')`""
    $shortcut.WorkingDirectory = $PSScriptRoot
    $shortcut.IconLocation = "$app,0"
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
Write-Host "Done. Shortcuts launch the compiled tray assistant directly; use the tray icon to exit."
