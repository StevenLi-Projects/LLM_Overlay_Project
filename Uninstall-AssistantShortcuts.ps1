Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$shortcutNames = @("Local Text Formatting Assistant.lnk")
$folders = @(
    [Environment]::GetFolderPath("DesktopDirectory"),
    [Environment]::GetFolderPath("Startup")
)

foreach ($folder in $folders) {
    foreach ($name in $shortcutNames) {
        $path = Join-Path $folder $name
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
            Write-Host "Removed shortcut: $path"
        }
    }
}

Write-Host "Shortcut cleanup completed."
