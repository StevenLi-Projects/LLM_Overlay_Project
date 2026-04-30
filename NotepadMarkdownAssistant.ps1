param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json")
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:llamaHealthConfirmedAt = [DateTime]::MinValue
$script:serverModelPath = $null

$createdNew = $false
$appMutex = New-Object System.Threading.Mutex($true, "LocalTextFormattingAssistant.llama.cpp", [ref]$createdNew)
if (!$createdNew) {
    [System.Windows.Forms.MessageBox]::Show(
        "The Local Text Formatting Assistant is already running. Check the Windows tray, exit the existing instance, then launch it again.",
        "Assistant already running",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
    return
}

$helperSource = @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class HotkeyWindow : Form {
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    private const int WM_HOTKEY = 0x0312;
    private readonly System.Collections.Generic.List<int> ids = new System.Collections.Generic.List<int>();
    public event Action<int> HotkeyPressed;

    public HotkeyWindow() {
        this.ShowInTaskbar = false;
        this.WindowState = FormWindowState.Minimized;
        this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
    }

    protected override void SetVisibleCore(bool value) {
        base.SetVisibleCore(false);
    }

    public void AddHotkey(int id, uint modifiers, uint key) {
        if (!RegisterHotKey(this.Handle, id, modifiers, key)) {
            throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "Could not register hotkey " + id);
        }
        ids.Add(id);
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_HOTKEY && HotkeyPressed != null) {
            HotkeyPressed(m.WParam.ToInt32());
        }
        base.WndProc(ref m);
    }

    protected override void OnFormClosed(FormClosedEventArgs e) {
        foreach (int id in ids) {
            UnregisterHotKey(this.Handle, id);
        }
        base.OnFormClosed(e);
    }
}

public static class NativeFocus {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}

public class ClipboardSnapshot {
    private IDataObject data;

    public static ClipboardSnapshot Capture() {
        ClipboardSnapshot snapshot = new ClipboardSnapshot();
        for (int i = 0; i < 8; i++) {
            try {
                snapshot.data = Clipboard.GetDataObject();
                return snapshot;
            } catch {
                System.Threading.Thread.Sleep(50);
            }
        }
        return snapshot;
    }

    public void Restore() {
        if (data == null) return;
        for (int i = 0; i < 8; i++) {
            try {
                Clipboard.SetDataObject(data, true);
                return;
            } catch {
                System.Threading.Thread.Sleep(80);
            }
        }
    }
}
"@

Add-Type -ReferencedAssemblies "System.Windows.Forms.dll","System.Drawing.dll" -TypeDefinition $helperSource

function Resolve-LocalPath {
    param([string]$PathValue)
    if ([string]::IsNullOrWhiteSpace($PathValue)) { return $PathValue }
    if ([System.IO.Path]::IsPathRooted($PathValue)) { return $PathValue }
    return (Join-Path $PSScriptRoot $PathValue)
}

function Load-Config {
    param([string]$Path)
    if (!(Test-Path -LiteralPath $Path)) {
        throw "Config file not found: $Path. Copy config.example.json to config.json first."
    }

    $config = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
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
    return $config
}

function Show-Notice {
    param(
        [System.Windows.Forms.NotifyIcon]$Tray,
        [string]$Title,
        [string]$Message,
        [System.Windows.Forms.ToolTipIcon]$Icon = [System.Windows.Forms.ToolTipIcon]::Info,
        [bool]$Balloon = $true
    )
    Write-Host "[$Title] $Message"
    $notificationsEnabled = $true
    $configVariable = Get-Variable -Name "config" -Scope Script -ErrorAction SilentlyContinue
    if ($configVariable -and $script:config -and $script:config.PSObject.Properties.Name -contains "ui") {
        $notificationsEnabled = Get-ConfigBool -Object $script:config.ui -Name "show_notifications" -Default $true
    }
    if (!$notificationsEnabled) {
        $Balloon = $false
    }
    if ($Tray -and $Balloon) {
        $Tray.ShowBalloonTip(3500, $Title, $Message, $Icon)
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

function Get-ConfigInt {
    param(
        [object]$Object,
        [string]$Name,
        [int]$Default
    )
    return [int](Get-PropertyValue -Object $Object -Name $Name -Default $Default)
}

function Get-ConfigBool {
    param(
        [object]$Object,
        [string]$Name,
        [bool]$Default
    )
    return [bool](Get-PropertyValue -Object $Object -Name $Name -Default $Default)
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

function Get-ActiveProfileName {
    param([object]$Config)
    return [string](Get-PropertyValue -Object $Config.llama -Name "active_profile" -Default "normal")
}

function Get-ActiveProfile {
    param([object]$Config)
    if (!$Config.llama -or !($Config.llama.PSObject.Properties.Name -contains "profiles")) {
        return $null
    }

    $profileName = Get-ActiveProfileName -Config $Config
    if ($Config.llama.profiles.PSObject.Properties.Name -contains $profileName) {
        return $Config.llama.profiles.$profileName
    }
    if ($Config.llama.profiles.PSObject.Properties.Name -contains "normal") {
        Set-ObjectProperty -Object $Config.llama -Name "active_profile" -Value "normal"
        return $Config.llama.profiles.normal
    }
    $first = $Config.llama.profiles.PSObject.Properties | Select-Object -First 1
    if ($first) {
        Set-ObjectProperty -Object $Config.llama -Name "active_profile" -Value $first.Name
        return $first.Value
    }
    return $null
}

function Apply-LlamaProfile {
    param([object]$Config)

    $profile = Get-ActiveProfile -Config $Config
    if (!$profile) { return }

    foreach ($name in @("model_path", "model_name", "context_size", "gpu_layers", "server_args")) {
        if ($profile.PSObject.Properties.Name -contains $name) {
            Set-ObjectProperty -Object $Config.llama -Name $name -Value $profile.$name
        }
    }
}

function Get-ProfileLabel {
    param(
        [string]$Name,
        [object]$Profile
    )
    $fallback = $Name
    if ($Name -eq "normal") { $fallback = "Normal (E4B quality)" }
    if ($Name -eq "fast") { $fallback = "Fast (E2B)" }
    return [string](Get-PropertyValue -Object $Profile -Name "label" -Default $fallback)
}

function Convert-Hotkey {
    param([string]$Hotkey)

    $modifiers = 0x4000 # MOD_NOREPEAT
    $keyName = $null
    foreach ($part in ($Hotkey -split "\+")) {
        $token = $part.Trim().ToLowerInvariant()
        switch ($token) {
            "ctrl" { $modifiers = $modifiers -bor 0x0002 }
            "control" { $modifiers = $modifiers -bor 0x0002 }
            "alt" { $modifiers = $modifiers -bor 0x0001 }
            "shift" { $modifiers = $modifiers -bor 0x0004 }
            "win" { $modifiers = $modifiers -bor 0x0008 }
            default { $keyName = $part.Trim() }
        }
    }

    if (!$keyName) { throw "Invalid hotkey '$Hotkey': missing final key." }
    $key = [System.Enum]::Parse([System.Windows.Forms.Keys], $keyName, $true)
    return @{ Modifiers = [uint32]$modifiers; Key = [uint32]$key }
}

function Get-ModeInstruction {
    param([string]$Mode)

    switch ($Mode) {
        "markdown" { return "Clean Markdown; use headings, lists, code fences, tables, and emphasis only when useful." }
        "bullets" { return "Concise Markdown bullets; preserve hierarchy and key details." }
        "table" { return "Markdown table if appropriate; otherwise clean Markdown." }
        "cleanup" { return "Fix spelling, spacing, punctuation, and structure without changing meaning." }
        "summary" { return "Concise Markdown summary preserving essential meaning and decisions." }
        default { return "Clean Markdown." }
    }
}

function Get-PromptForMode {
    param(
        [string]$Mode,
        [string]$SelectedText
    )

    $instructions = Get-ModeInstruction -Mode $Mode

    return @"
Task: $instructions
Rules: preserve meaning; add no facts; return only transformed text; no commentary.
Text:
$SelectedText
Output:
"@
}

function Get-ModeMaxTokens {
    param(
        [object]$Config,
        [string]$Mode
    )

    $fallback = Get-ConfigInt -Object $Config.generation -Name "max_tokens" -Default 1024
    if ($Config.modes -and ($Config.modes.PSObject.Properties.Name -contains $Mode)) {
        return Get-ConfigInt -Object $Config.modes.$Mode -Name "max_tokens" -Default $fallback
    }
    return $fallback
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

function Clear-LlamaHealthCache {
    $script:llamaHealthConfirmedAt = [DateTime]::MinValue
}

function Get-ConfiguredModelPaths {
    param([object]$Config)

    $paths = @()
    if ($Config.llama.PSObject.Properties.Name -contains "model_path") {
        $paths += [string]$Config.llama.model_path
    }
    if ($Config.llama.PSObject.Properties.Name -contains "profiles") {
        foreach ($profile in $Config.llama.profiles.PSObject.Properties) {
            if ($profile.Value.PSObject.Properties.Name -contains "model_path") {
                $paths += [string]$profile.Value.model_path
            }
        }
    }
    return @($paths | Where-Object { ![string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Stop-ConfiguredLlamaServers {
    param([object]$Config)

    $exe = Join-Path $Config.llama.cpp_dir "llama-server.exe"
    $models = Get-ConfiguredModelPaths -Config $Config
    if ($models.Count -eq 0) { return }

    $servers = Get-CimInstance Win32_Process -Filter "name = 'llama-server.exe'" -ErrorAction SilentlyContinue
    foreach ($server in $servers) {
        $commandLine = [string]$server.CommandLine
        $matchesExe = $commandLine.IndexOf($exe, [StringComparison]::OrdinalIgnoreCase) -ge 0
        $matchesModel = $false
        foreach ($model in $models) {
            if ($commandLine.IndexOf($model, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $matchesModel = $true
                break
            }
        }
        if ($matchesExe -or $matchesModel) {
            try {
                Stop-Process -Id $server.ProcessId -Force -ErrorAction Stop
            } catch {
                Write-Warning "Could not stop llama-server process $($server.ProcessId): $($_.Exception.Message)"
            }
        }
    }
    Clear-LlamaHealthCache
    $script:serverModelPath = $null
}

function Set-ActiveLlamaProfile {
    param(
        [object]$Config,
        [string]$ProfileName,
        [System.Windows.Forms.ContextMenuStrip]$TrayMenu = $null,
        [System.Windows.Forms.NotifyIcon]$Tray = $null
    )

    if (!($Config.llama.PSObject.Properties.Name -contains "profiles")) {
        throw "No llama.profiles are configured."
    }
    if (!($Config.llama.profiles.PSObject.Properties.Name -contains $ProfileName)) {
        throw "Unknown llama profile: $ProfileName"
    }

    $previousModel = [string](Get-PropertyValue -Object $Config.llama -Name "model_path" -Default "")
    Set-ObjectProperty -Object $Config.llama -Name "active_profile" -Value $ProfileName
    Apply-LlamaProfile -Config $Config
    Clear-LlamaHealthCache

    if (![string]::IsNullOrWhiteSpace($previousModel) -and
        $previousModel -ne [string]$Config.llama.model_path) {
        Stop-ConfiguredLlamaServers -Config $Config
    }

    Update-ProfileMenuChecks -Menu $TrayMenu -Config $Config
    Update-TrayText -Tray $Tray -Config $Config
    Show-Notice $Tray "Model mode" "Using $(Get-ProfileLabel -Name $ProfileName -Profile $Config.llama.profiles.$ProfileName)." ([System.Windows.Forms.ToolTipIcon]::Info)
}

function Invoke-Llama {
    param(
        [object]$Config,
        [string]$Prompt,
        [int]$MaxTokens
    )

    $baseUrl = $Config.llama.server_url.TrimEnd("/")
    $timeoutSec = [int]$Config.generation.timeout_sec
    $preferCompletion = Get-ConfigBool -Object $Config.generation -Name "prefer_completion" -Default $true

    $chatBody = @{
        model = $Config.llama.model_name
        messages = @(
            @{ role = "user"; content = "Return only transformed text. $Prompt" }
        )
        temperature = [double]$Config.generation.temperature
        top_p = [double]$Config.generation.top_p
        max_tokens = $MaxTokens
        stream = $false
    } | ConvertTo-Json -Depth 8

    $completionBody = @{
        prompt = $Prompt
        temperature = [double]$Config.generation.temperature
        top_p = [double]$Config.generation.top_p
        n_predict = $MaxTokens
        stream = $false
    } | ConvertTo-Json -Depth 5

    $chatError = "not attempted"
    $completionError = "not attempted"

    if ($preferCompletion) {
        try {
            $response = Invoke-RestMethod -Method Post -Uri "$baseUrl/completion" -ContentType "application/json" -Body $completionBody -TimeoutSec $timeoutSec
            if ($response.content -and ![string]::IsNullOrWhiteSpace($response.content)) {
                return [pscustomobject]@{
                    Text = $response.content.Trim()
                    Endpoint = "completion"
                    TokensPredicted = [int](Get-PropertyValue -Object $response -Name "tokens_predicted" -Default 0)
                    TokensEvaluated = [int](Get-PropertyValue -Object $response -Name "tokens_evaluated" -Default 0)
                }
            }
            $completionError = "empty response"
        } catch {
            $completionError = $_.Exception.Message
        }

        try {
            $response = Invoke-RestMethod -Method Post -Uri "$baseUrl/v1/chat/completions" -ContentType "application/json" -Body $chatBody -TimeoutSec $timeoutSec
            $text = $response.choices[0].message.content
            if (![string]::IsNullOrWhiteSpace($text)) {
                $usage = Get-PropertyValue -Object $response -Name "usage" -Default $null
                return [pscustomobject]@{
                    Text = $text.Trim()
                    Endpoint = "chat"
                    TokensPredicted = [int](Get-PropertyValue -Object $usage -Name "completion_tokens" -Default 0)
                    TokensEvaluated = [int](Get-PropertyValue -Object $usage -Name "prompt_tokens" -Default 0)
                }
            }
            $chatError = "empty response"
        } catch {
            $chatError = $_.Exception.Message
        }
    } else {
        try {
            $response = Invoke-RestMethod -Method Post -Uri "$baseUrl/v1/chat/completions" -ContentType "application/json" -Body $chatBody -TimeoutSec $timeoutSec
            $text = $response.choices[0].message.content
            if (![string]::IsNullOrWhiteSpace($text)) {
                $usage = Get-PropertyValue -Object $response -Name "usage" -Default $null
                return [pscustomobject]@{
                    Text = $text.Trim()
                    Endpoint = "chat"
                    TokensPredicted = [int](Get-PropertyValue -Object $usage -Name "completion_tokens" -Default 0)
                    TokensEvaluated = [int](Get-PropertyValue -Object $usage -Name "prompt_tokens" -Default 0)
                }
            }
            $chatError = "empty response"
        } catch {
            $chatError = $_.Exception.Message
        }

        try {
            $response = Invoke-RestMethod -Method Post -Uri "$baseUrl/completion" -ContentType "application/json" -Body $completionBody -TimeoutSec $timeoutSec
            if ($response.content -and ![string]::IsNullOrWhiteSpace($response.content)) {
                return [pscustomobject]@{
                    Text = $response.content.Trim()
                    Endpoint = "completion"
                    TokensPredicted = [int](Get-PropertyValue -Object $response -Name "tokens_predicted" -Default 0)
                    TokensEvaluated = [int](Get-PropertyValue -Object $response -Name "tokens_evaluated" -Default 0)
                }
            }
            $completionError = "empty response"
        } catch {
            $completionError = $_.Exception.Message
        }
    }

    Clear-LlamaHealthCache
    throw "llama.cpp request failed. Completion endpoint error: $completionError. Chat endpoint error: $chatError"
}

function Test-LlamaServer {
    param(
        [object]$Config,
        [switch]$Force
    )

    $currentModelPath = [string](Get-PropertyValue -Object $Config.llama -Name "model_path" -Default "")
    if (!$Force -and $script:serverModelPath -and $currentModelPath -and
        $script:serverModelPath -ne $currentModelPath) {
        Clear-LlamaHealthCache
        return $false
    }

    if (!$Force) {
        $healthCacheSec = Get-ConfigInt -Object $Config.llama -Name "health_cache_sec" -Default 30
        if ($healthCacheSec -gt 0 -and $script:llamaHealthConfirmedAt -ne [DateTime]::MinValue) {
            if (((Get-Date) - $script:llamaHealthConfirmedAt).TotalSeconds -lt $healthCacheSec) {
                return $true
            }
        }
    }

    $baseUrl = $Config.llama.server_url.TrimEnd("/")
    try {
        Invoke-RestMethod -Method Get -Uri "$baseUrl/health" -TimeoutSec 2 | Out-Null
        $script:llamaHealthConfirmedAt = Get-Date
        $script:serverModelPath = $currentModelPath
        return $true
    } catch {
        try {
            Invoke-RestMethod -Method Get -Uri "$baseUrl/v1/models" -TimeoutSec 2 | Out-Null
            $script:llamaHealthConfirmedAt = Get-Date
            $script:serverModelPath = $currentModelPath
            return $true
        } catch {
            Clear-LlamaHealthCache
            return $false
        }
    }
}

function Start-LlamaServerIfNeeded {
    param([object]$Config)

    if (Test-LlamaServer $Config) { return }
    if (!$Config.llama.auto_start_server) {
        throw "llama.cpp server is not reachable at $($Config.llama.server_url). Start it with .\Start-LlamaServer.ps1 or enable llama.auto_start_server in config.json."
    }

    $exe = Join-Path $Config.llama.cpp_dir "llama-server.exe"
    if (!(Test-Path -LiteralPath $exe)) { throw "llama-server.exe not found: $exe" }
    if (!(Test-Path -LiteralPath $Config.llama.model_path)) { throw "Model not found: $($Config.llama.model_path)" }

    Stop-ConfiguredLlamaServers -Config $Config
    $args = Get-LlamaServerArgs -Config $Config

    Start-Process -FilePath $exe -ArgumentList $args -WorkingDirectory $Config.llama.cpp_dir -WindowStyle Hidden | Out-Null

    $deadline = (Get-Date).AddSeconds([int]$Config.llama.startup_wait_sec)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        if (Test-LlamaServer $Config -Force) { return }
    }

    throw "Started llama-server, but it did not become ready at $($Config.llama.server_url). Check the llama-server window for model or GPU errors."
}

function Get-ClipboardTextWithRetry {
    for ($i = 0; $i -lt 8; $i++) {
        try {
            if ([System.Windows.Forms.Clipboard]::ContainsText()) {
                return [System.Windows.Forms.Clipboard]::GetText()
            }
            return ""
        } catch {
            Start-Sleep -Milliseconds 60
        }
    }
    return ""
}

function Set-ClipboardTextWithRetry {
    param([string]$Text)
    for ($i = 0; $i -lt 8; $i++) {
        try {
            [System.Windows.Forms.Clipboard]::SetText($Text)
            return
        } catch {
            Start-Sleep -Milliseconds 80
        }
    }
    throw "Could not write to the clipboard. Another app may be locking it."
}

function Show-ReplacementPreview {
    param(
        [string]$OriginalText,
        [string]$ReplacementText,
        [string]$Mode,
        [string]$TelemetryText = ""
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Local Text Formatter - Preview"
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(760, 520)
    $form.MinimumSize = New-Object System.Drawing.Size(560, 380)
    $form.TopMost = $true
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 251)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $layout.ColumnCount = 1
    $layout.RowCount = 4
    $layout.Padding = New-Object System.Windows.Forms.Padding(12)
    $layout.BackColor = $form.BackColor
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 60)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))

    $monoFont = New-Object System.Drawing.Font("Cascadia Mono", 10)
    if ($monoFont.Name -ne "Cascadia Mono") {
        $monoFont = New-Object System.Drawing.Font("Consolas", 10)
    }

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = [System.Windows.Forms.DockStyle]::Fill
    $header.BackColor = [System.Drawing.Color]::FromArgb(31, 42, 68)
    $header.Padding = New-Object System.Windows.Forms.Padding(14, 8, 14, 8)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "Replacement Preview"
    $title.ForeColor = [System.Drawing.Color]::White
    $title.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 13)
    $title.AutoSize = $true
    $title.Location = New-Object System.Drawing.Point(14, 8)

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = "Mode: $Mode. Edit if needed, then replace."
    $subtitle.ForeColor = [System.Drawing.Color]::FromArgb(215, 222, 235)
    $subtitle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $subtitle.AutoSize = $true
    $subtitle.Location = New-Object System.Drawing.Point(16, 34)

    [void]$header.Controls.Add($title)
    [void]$header.Controls.Add($subtitle)

    $replacementBox = New-Object System.Windows.Forms.TextBox
    $replacementBox.Multiline = $true
    $replacementBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $replacementBox.WordWrap = $false
    $replacementBox.AcceptsReturn = $true
    $replacementBox.AcceptsTab = $true
    $replacementBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $replacementBox.Font = $monoFont
    $replacementBox.Text = $ReplacementText
    $replacementBox.BackColor = [System.Drawing.Color]::White
    $replacementBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $replacementBox.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 6)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(74, 85, 104)
    $statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    if ([string]::IsNullOrWhiteSpace($TelemetryText)) {
        $statusLabel.Text = "Local llama.cpp output. Nothing is replaced until you choose Replace."
    } else {
        $statusLabel.Text = $TelemetryText
    }

    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
    $buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $buttonPanel.AutoSize = $true
    $buttonPanel.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
    $buttonPanel.BackColor = $form.BackColor

    $replaceButton = New-Object System.Windows.Forms.Button
    $replaceButton.Text = "Replace"
    $replaceButton.Width = 124
    $replaceButton.Height = 34
    $replaceButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $replaceButton.BackColor = [System.Drawing.Color]::FromArgb(22, 101, 52)
    $replaceButton.ForeColor = [System.Drawing.Color]::White
    $replaceButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $replaceButton.FlatAppearance.BorderSize = 0

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Width = 112
    $cancelButton.Height = 34
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
    $cancelButton.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
    $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelButton.FlatAppearance.BorderSize = 0

    [void]$buttonPanel.Controls.Add($replaceButton)
    [void]$buttonPanel.Controls.Add($cancelButton)

    [void]$layout.Controls.Add($header, 0, 0)
    [void]$layout.Controls.Add($replacementBox, 0, 1)
    [void]$layout.Controls.Add($statusLabel, 0, 2)
    [void]$layout.Controls.Add($buttonPanel, 0, 3)

    $form.Controls.Add($layout)
    $form.AcceptButton = $replaceButton
    $form.CancelButton = $cancelButton
    $replacementBox.Select(0, 0)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $replacementBox.Text
    }

    return $null
}

function Write-TimingDiagnostics {
    param(
        [string]$Mode,
        [System.Collections.Specialized.OrderedDictionary]$Timings,
        [object]$Telemetry = $null
    )

    $parts = @()
    foreach ($key in $Timings.Keys) {
        $parts += "$key=$($Timings[$key])ms"
    }
    if ($Telemetry) {
        if ($Telemetry.Endpoint) { $parts += "endpoint=$($Telemetry.Endpoint)" }
        if ($Telemetry.TokensPredicted -gt 0) { $parts += "tokens=$($Telemetry.TokensPredicted)" }
        if ($Telemetry.Tps -gt 0) { $parts += ("tps={0:N1}" -f $Telemetry.Tps) }
        if ($Telemetry.TokensEvaluated -gt 0) { $parts += "prompt_tokens=$($Telemetry.TokensEvaluated)" }
    }
    Write-Host "[Timing][$Mode] $($parts -join ', ')"
}

function Format-TelemetryText {
    param(
        [object]$Telemetry,
        [long]$GenerationMs,
        [int]$MaxTokens
    )

    if (!$Telemetry) {
        return "Generated locally. Max output: $MaxTokens tokens."
    }

    $parts = @()
    if ($Telemetry.Endpoint) { $parts += "endpoint: $($Telemetry.Endpoint)" }
    if ($GenerationMs -gt 0) { $parts += ("generation: {0:N1}s" -f ($GenerationMs / 1000.0)) }
    if ($Telemetry.TokensPredicted -gt 0) { $parts += "output tokens: $($Telemetry.TokensPredicted)" }
    if ($Telemetry.Tps -gt 0) { $parts += ("TPS: {0:N1}" -f $Telemetry.Tps) }
    if ($Telemetry.TokensEvaluated -gt 0) { $parts += "prompt tokens: $($Telemetry.TokensEvaluated)" }
    $parts += "max output: $MaxTokens"
    return ($parts -join "   |   ")
}

function Invoke-FormatSelection {
    param(
        [string]$Mode,
        [object]$Config,
        [System.Windows.Forms.NotifyIcon]$Tray,
        [IntPtr]$TargetWindow = [IntPtr]::Zero
    )

    $totalSw = [Diagnostics.Stopwatch]::StartNew()
    $timings = [ordered]@{}
    $telemetry = $null
    try {
        $stageSw = [Diagnostics.Stopwatch]::StartNew()
        Start-LlamaServerIfNeeded $Config
        $stageSw.Stop()
        $timings.server = $stageSw.ElapsedMilliseconds

        $stageSw.Restart()
        $snapshot = [ClipboardSnapshot]::Capture()
        [System.Windows.Forms.Clipboard]::Clear()
        if ($TargetWindow -ne [IntPtr]::Zero) {
            [NativeFocus]::SetForegroundWindow($TargetWindow) | Out-Null
            Start-Sleep -Milliseconds 120
        }
        [System.Windows.Forms.SendKeys]::SendWait("^c")
        Start-Sleep -Milliseconds ([int]$Config.ui.copy_wait_ms)
        $selected = Get-ClipboardTextWithRetry
        $stageSw.Stop()
        $timings.copy = $stageSw.ElapsedMilliseconds

        if ([string]::IsNullOrWhiteSpace($selected)) {
            $snapshot.Restore()
            $totalSw.Stop()
            $timings.total = $totalSw.ElapsedMilliseconds
            Write-TimingDiagnostics -Mode $Mode -Timings $timings -Telemetry $telemetry
            Show-Notice $Tray "No selected text" "Select editable text in the target app, then press the hotkey again." ([System.Windows.Forms.ToolTipIcon]::Warning)
            return
        }

        Show-Notice $Tray "Formatting" "Sending selected text to local llama.cpp ($Mode)." ([System.Windows.Forms.ToolTipIcon]::Info)
        $prompt = Get-PromptForMode -Mode $Mode -SelectedText $selected
        $maxTokens = Get-ModeMaxTokens -Config $Config -Mode $Mode
        $stageSw.Restart()
        $result = Invoke-Llama -Config $Config -Prompt $prompt -MaxTokens $maxTokens
        $stageSw.Stop()
        $timings.generation = $stageSw.ElapsedMilliseconds
        $output = $result.Text
        $telemetry = [pscustomobject]@{
            Endpoint = $result.Endpoint
            TokensPredicted = [int]$result.TokensPredicted
            TokensEvaluated = [int]$result.TokensEvaluated
            Tps = 0.0
        }
        if ($telemetry.TokensPredicted -gt 0 -and $timings.generation -gt 0) {
            $telemetry.Tps = $telemetry.TokensPredicted / ($timings.generation / 1000.0)
        }

        if ([string]::IsNullOrWhiteSpace($output)) {
            $snapshot.Restore()
            $totalSw.Stop()
            $timings.total = $totalSw.ElapsedMilliseconds
            Write-TimingDiagnostics -Mode $Mode -Timings $timings -Telemetry $telemetry
            Show-Notice $Tray "Invalid response" "llama.cpp returned no usable text." ([System.Windows.Forms.ToolTipIcon]::Error)
            return
        }

        $previewEnabled = Get-ConfigBool -Object $Config.ui -Name "preview_enabled" -Default $true
        $telemetryText = Format-TelemetryText -Telemetry $telemetry -GenerationMs $timings.generation -MaxTokens $maxTokens
        if ($previewEnabled) {
            $stageSw.Restart()
            $previewOutput = Show-ReplacementPreview -OriginalText $selected -ReplacementText $output -Mode $Mode -TelemetryText $telemetryText
            $stageSw.Stop()
            $timings.preview = $stageSw.ElapsedMilliseconds
            if ($null -eq $previewOutput) {
                $snapshot.Restore()
                $totalSw.Stop()
                $timings.total = $totalSw.ElapsedMilliseconds
                Write-TimingDiagnostics -Mode $Mode -Timings $timings -Telemetry $telemetry
                Show-Notice $Tray "Canceled" "No text was replaced." ([System.Windows.Forms.ToolTipIcon]::Info)
                return
            }
            if ([string]::IsNullOrWhiteSpace($previewOutput)) {
                $snapshot.Restore()
                $totalSw.Stop()
                $timings.total = $totalSw.ElapsedMilliseconds
                Write-TimingDiagnostics -Mode $Mode -Timings $timings -Telemetry $telemetry
                Show-Notice $Tray "Empty replacement" "No text was replaced because the preview replacement was empty." ([System.Windows.Forms.ToolTipIcon]::Warning)
                return
            }
            $output = $previewOutput
        } else {
            $timings.preview = 0
        }

        $stageSw.Restart()
        Set-ClipboardTextWithRetry $output
        if ($TargetWindow -ne [IntPtr]::Zero) {
            [NativeFocus]::SetForegroundWindow($TargetWindow) | Out-Null
            Start-Sleep -Milliseconds 120
        }
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        Start-Sleep -Milliseconds ([int]$Config.ui.paste_wait_ms)
        $snapshot.Restore()
        $stageSw.Stop()
        $timings.paste = $stageSw.ElapsedMilliseconds
        $totalSw.Stop()
        $timings.total = $totalSw.ElapsedMilliseconds
        Write-TimingDiagnostics -Mode $Mode -Timings $timings -Telemetry $telemetry

        $showTiming = Get-ConfigBool -Object $Config.ui -Name "show_timing_notifications" -Default $true
        if ($showTiming) {
            if ($telemetry -and $telemetry.Tps -gt 0) {
                Show-Notice $Tray "Done" ("Generated in {0:N1}s at {1:N1} TPS; total {2:N1}s." -f ($timings.generation / 1000.0), $telemetry.Tps, ($timings.total / 1000.0)) ([System.Windows.Forms.ToolTipIcon]::Info)
            } else {
                Show-Notice $Tray "Done" ("Generated in {0:N1}s; total {1:N1}s." -f ($timings.generation / 1000.0), ($timings.total / 1000.0)) ([System.Windows.Forms.ToolTipIcon]::Info)
            }
        } else {
            Show-Notice $Tray "Done" "Replaced selection and restored the previous clipboard." ([System.Windows.Forms.ToolTipIcon]::Info)
        }
    } catch {
        Clear-LlamaHealthCache
        $totalSw.Stop()
        $timings.total = $totalSw.ElapsedMilliseconds
        Write-TimingDiagnostics -Mode $Mode -Timings $timings -Telemetry $telemetry
        Show-Notice $Tray "Assistant error" $_.Exception.Message ([System.Windows.Forms.ToolTipIcon]::Error)
    }
}

function Show-ModeMenu {
    param(
        [System.Windows.Forms.ContextMenuStrip]$ModeMenu,
        [IntPtr]$TargetWindow
    )
    $script:menuTargetWindow = $TargetWindow
    $ModeMenu.Show([System.Windows.Forms.Cursor]::Position)
}

function Update-ProfileMenuChecks {
    param(
        [System.Windows.Forms.ContextMenuStrip]$Menu,
        [object]$Config
    )
    if (!$Menu -or !($Config.llama.PSObject.Properties.Name -contains "profiles")) { return }
    $active = Get-ActiveProfileName -Config $Config
    foreach ($item in $Menu.Items) {
        if ($item.Tag -and ([string]$item.Tag).StartsWith("profile:")) {
            $profileName = ([string]$item.Tag).Substring(8)
            $item.Checked = ($profileName -eq $active)
        }
    }
}

function Update-TrayText {
    param(
        [System.Windows.Forms.NotifyIcon]$Tray,
        [object]$Config
    )
    if (!$Tray) { return }
    $label = "Local Text Formatting Assistant"
    if ($Config.llama.PSObject.Properties.Name -contains "profiles") {
        $active = Get-ActiveProfileName -Config $Config
        $profile = Get-ActiveProfile -Config $Config
        $label = "Text Assistant - $(Get-ProfileLabel -Name $active -Profile $profile)"
    }
    if ($label.Length -gt 63) {
        $label = $label.Substring(0, 63)
    }
    $Tray.Text = $label
}

$config = Load-Config $ConfigPath
$window = New-Object HotkeyWindow
$tray = New-Object System.Windows.Forms.NotifyIcon
$tray.Icon = [System.Drawing.SystemIcons]::Application
$tray.Visible = $true

$modeLabels = @{
    markdown = "Format as Markdown"
    bullets = "Convert to bullet points"
    table = "Convert to table when appropriate"
    cleanup = "Clean up without changing meaning"
    summary = "Summarize into concise Markdown"
}

$modeMenu = New-Object System.Windows.Forms.ContextMenuStrip
foreach ($mode in $config.modes.PSObject.Properties) {
    if (!$mode.Value.enabled) { continue }
    $defaultLabel = $mode.Name
    if ($modeLabels.ContainsKey($mode.Name)) {
        $defaultLabel = $modeLabels[$mode.Name]
    }
    $label = Get-PropertyValue -Object $mode.Value -Name "label" -Default $defaultLabel
    $modeName = $mode.Name
    $item = $modeMenu.Items.Add($label)
    $item.Tag = $modeName
    $item.Add_Click({
        param($sender, $eventArgs)
        Invoke-FormatSelection -Mode $sender.Tag -Config $script:config -Tray $script:tray -TargetWindow $script:menuTargetWindow
    })
}

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$itemTitle = $menu.Items.Add("Local Text Formatting Assistant")
$itemTitle.Enabled = $false
[void]$menu.Items.Add("-")
$itemProfileHeader = $menu.Items.Add("Model mode")
$itemProfileHeader.Enabled = $false
if ($config.llama.PSObject.Properties.Name -contains "profiles") {
    foreach ($profile in $config.llama.profiles.PSObject.Properties) {
        $profileName = $profile.Name
        $profileLabel = Get-ProfileLabel -Name $profileName -Profile $profile.Value
        $itemProfile = $menu.Items.Add($profileLabel)
        $itemProfile.Tag = "profile:$profileName"
        $itemProfile.Add_Click({
            param($sender, $eventArgs)
            $name = ([string]$sender.Tag).Substring(8)
            try {
                Set-ActiveLlamaProfile -Config $script:config -ProfileName $name -TrayMenu $script:tray.ContextMenuStrip -Tray $script:tray
            } catch {
                Show-Notice $script:tray "Model mode error" $_.Exception.Message ([System.Windows.Forms.ToolTipIcon]::Error)
            }
        })
    }
    [void]$menu.Items.Add("-")
}
$itemModeInfo = $menu.Items.Add("Use the popup hotkey over selected editable text")
$itemModeInfo.Enabled = $false
[void]$menu.Items.Add("-")
$itemStart = $menu.Items.Add("Start llama.cpp server")
$itemStart.Add_Click({
    try {
        Start-LlamaServerIfNeeded $script:config
        Show-Notice $script:tray "llama.cpp" "Server is reachable at $($script:config.llama.server_url)." ([System.Windows.Forms.ToolTipIcon]::Info)
    } catch {
        Show-Notice $script:tray "llama.cpp error" $_.Exception.Message ([System.Windows.Forms.ToolTipIcon]::Error)
    }
})
$itemExit = $menu.Items.Add("Exit")
$itemExit.Add_Click({
    $script:tray.Visible = $false
    $script:window.Close()
    [System.Windows.Forms.Application]::Exit()
})
$tray.ContextMenuStrip = $menu
Update-ProfileMenuChecks -Menu $menu -Config $config
Update-TrayText -Tray $tray -Config $config

$script:config = $config
$script:tray = $tray
$script:window = $window
$script:menuTargetWindow = [IntPtr]::Zero
$hotkeyModes = @{}
$hotkeyActions = @{}
$id = 100
foreach ($mode in $config.modes.PSObject.Properties) {
    if (!$mode.Value.enabled) { continue }
    try {
        $parsed = Convert-Hotkey $mode.Value.hotkey
        $window.AddHotkey($id, $parsed.Modifiers, $parsed.Key)
        $hotkeyModes[$id] = $mode.Name
        $hotkeyActions[$id] = "mode"
        Write-Host "Registered $($mode.Value.hotkey) for mode '$($mode.Name)'"
    } catch {
        Write-Warning "Could not register $($mode.Value.hotkey) for mode '$($mode.Name)': $($_.Exception.Message)"
    }
    $id++
}

$menuHotkey = Get-PropertyValue -Object $config.ui -Name "menu_hotkey" -Default "Ctrl+Alt+Space"
try {
    $parsed = Convert-Hotkey $menuHotkey
    $window.AddHotkey($id, $parsed.Modifiers, $parsed.Key)
    $hotkeyActions[$id] = "menu"
    Write-Host "Registered $menuHotkey for popup menu"
} catch {
    Write-Warning "Could not register popup menu hotkey '$menuHotkey': $($_.Exception.Message)"
}

$window.add_HotkeyPressed({
    param([int]$hotkeyId)
    $targetWindow = [NativeFocus]::GetForegroundWindow()
    if ($script:hotkeyActions.ContainsKey($hotkeyId) -and $script:hotkeyActions[$hotkeyId] -eq "menu") {
        Show-ModeMenu -ModeMenu $script:modeMenu -TargetWindow $targetWindow
    } elseif ($script:hotkeyModes.ContainsKey($hotkeyId)) {
        Invoke-FormatSelection -Mode $script:hotkeyModes[$hotkeyId] -Config $script:config -Tray $script:tray -TargetWindow $targetWindow
    }
})
$script:hotkeyModes = $hotkeyModes
$script:hotkeyActions = $hotkeyActions
$script:modeMenu = $modeMenu

if ($hotkeyActions.Count -eq 0) {
    Show-Notice $tray "No hotkeys registered" "Every configured hotkey is already in use. Edit config.json and restart the assistant." ([System.Windows.Forms.ToolTipIcon]::Error)
} else {
    Show-Notice $tray "Local Text Formatting Assistant" "Running. Use configured hotkeys with selected editable text." ([System.Windows.Forms.ToolTipIcon]::Info)
}
[System.Windows.Forms.Application]::Run($window)
$appMutex.ReleaseMutex()
$appMutex.Dispose()
