namespace LocalTextFormattingAssistant;

internal sealed class AssistantApplicationContext : ApplicationContext
{
    private readonly AppConfig _config;
    private readonly DiagnosticsBuffer _diagnostics = new();
    private readonly HotkeyWindow _hotkeys = new();
    private readonly LlamaClient _llamaClient;
    private readonly LlamaServerManager _server;
    private readonly NotifyIcon _tray = new();
    private readonly ContextMenuStrip _trayMenu = new();
    private readonly CommandPaletteForm _palette;
    private readonly ProgressOverlayForm _progress;
    private readonly Dictionary<int, string> _hotkeyModes = [];
    private readonly List<string> _hotkeyIssues = [];
    private ToolStripMenuItem? _profilesMenu;
    private ToolStripMenuItem? _previewItem;
    private ToolStripMenuItem? _serverStatusItem;
    private ToolStripMenuItem? _cancelItem;
    private CancellationTokenSource? _activeRequest;
    private IntPtr _paletteTargetWindow;
    private string? _lastMode;
    private bool _busy;
    private bool _shuttingDown;
    private AppTheme _theme;
    private Icon _appIcon;

    public AssistantApplicationContext(AppConfig config)
    {
        _config = config;
        _theme = ThemeManager.Resolve(config.Theme);
        _appIcon = ThemeManager.CreateAppIcon(_theme);
        _llamaClient = new LlamaClient(_diagnostics);
        _server = new LlamaServerManager(_diagnostics);
        _server.StatusChanged += OnServerStatusChanged;

        _palette = new CommandPaletteForm(config, _theme);
        _palette.Icon = _appIcon;
        _palette.ModeSelected += mode => _ = FormatSelectionAsync(mode, _paletteTargetWindow);
        _palette.ProfileSelected += profile => _ = ChangeProfileAsync(profile);
        _palette.PreviewChanged += enabled => SetPreview(enabled);
        _palette.SetStatus(_server.Status);

        _progress = new ProgressOverlayForm(_theme);
        _progress.Icon = _appIcon;
        _progress.CancelRequested += CancelOrCloseProgress;

        _hotkeys.HotkeyPressed += OnHotkeyPressed;
        BuildTrayMenu();
        RegisterHotkeys();

        _tray.Icon = _appIcon;
        _tray.Text = BuildTrayText();
        _tray.Visible = true;
        _tray.ContextMenuStrip = _trayMenu;
        _tray.DoubleClick += (_, _) => OpenPalette();
        _diagnostics.Add("Compiled tray application started.");
        if (_config.PrewarmOnLaunch) _ = PrewarmServerAsync();
    }

    private async Task PrewarmServerAsync()
    {
        try
        {
            await _server.EnsureReadyAsync(_config, CancellationToken.None);
        }
        catch (Exception ex)
        {
            _server.MarkError($"Prewarm failed: {ex.Message}");
        }
    }

    private void BuildTrayMenu()
    {
        _trayMenu.Items.Clear();
        var open = new ToolStripMenuItem("Open formatter");
        open.Click += (_, _) => OpenPalette();
        _trayMenu.Items.Add(open);
        _trayMenu.Items.Add(new ToolStripSeparator());

        _profilesMenu = new ToolStripMenuItem("Model profile");
        foreach (var profile in _config.Profiles)
        {
            var item = new ToolStripMenuItem(profile.Label) { Tag = profile.Name, Checked = profile.Name == _config.ActiveProfileName };
            item.Click += (_, _) => _ = ChangeProfileAsync((string)item.Tag!);
            _profilesMenu.DropDownItems.Add(item);
        }
        _trayMenu.Items.Add(_profilesMenu);

        _previewItem = new ToolStripMenuItem("Preview before replacing") { Checked = _config.PreviewEnabled, CheckOnClick = false };
        _previewItem.Click += (_, _) => SetPreview(!_config.PreviewEnabled);
        _trayMenu.Items.Add(_previewItem);

        _trayMenu.Items.Add(new ToolStripSeparator());
        _serverStatusItem = new ToolStripMenuItem(_server.Status.Detail) { Enabled = false };
        _trayMenu.Items.Add(_serverStatusItem);
        var restart = new ToolStripMenuItem("Start or restart server");
        restart.Click += async (_, _) => await RestartServerAsync();
        _trayMenu.Items.Add(restart);
        _cancelItem = new ToolStripMenuItem("Cancel current action") { Visible = false };
        _cancelItem.Click += (_, _) => _activeRequest?.Cancel();
        _trayMenu.Items.Add(_cancelItem);

        _trayMenu.Items.Add(new ToolStripSeparator());
        var settings = new ToolStripMenuItem("Settings");
        settings.Click += (_, _) => OpenSettings();
        _trayMenu.Items.Add(settings);
        var diagnostics = new ToolStripMenuItem("Copy diagnostics");
        diagnostics.Click += (_, _) => SelectionBridge.SetClipboardText(_diagnostics.GetText());
        _trayMenu.Items.Add(diagnostics);
        _trayMenu.Items.Add(new ToolStripSeparator());
        var exit = new ToolStripMenuItem("Exit");
        exit.Click += (_, _) => Shutdown();
        _trayMenu.Items.Add(exit);
        ApplyTrayTheme();
    }

    private void ApplyTrayTheme()
    {
        _trayMenu.BackColor = _theme.Surface;
        _trayMenu.ForeColor = _theme.Text;
        _trayMenu.Renderer = new ToolStripProfessionalRenderer(new AssistantColorTable(_theme));
    }

    private void RegisterHotkeys()
    {
        _hotkeys.Clear();
        _hotkeyModes.Clear();
        _hotkeyIssues.Clear();
        var id = 100;
        foreach (var mode in _config.Modes.Where(m => m.Enabled))
        {
            try
            {
                _hotkeys.Register(id, HotkeyParser.Parse(mode.Hotkey));
                _hotkeyModes[id] = mode.Name;
                _diagnostics.Add($"Registered {mode.Hotkey} for {mode.Name}.");
            }
            catch (Exception ex)
            {
                var issue = $"{mode.Label} ({mode.Hotkey}): {ex.Message}";
                _hotkeyIssues.Add(issue);
                _diagnostics.Add($"Hotkey conflict for {issue}");
            }
            id++;
        }
        try
        {
            _hotkeys.Register(id, HotkeyParser.Parse(_config.MenuHotkey));
            _hotkeyModes[id] = "__menu";
        }
        catch (Exception ex)
        {
            var issue = $"Open formatter ({_config.MenuHotkey}): {ex.Message}";
            _hotkeyIssues.Add(issue);
            _diagnostics.Add($"Palette hotkey conflict: {issue}");
        }
    }

    private void OnHotkeyPressed(int id)
    {
        if (!_hotkeyModes.TryGetValue(id, out var mode)) return;
        if (mode == "__menu") OpenPalette();
        else if (!_busy) _ = FormatSelectionAsync(mode, NativeMethods.GetForegroundWindow());
    }

    private void OpenPalette()
    {
        if (_busy) return;
        _paletteTargetWindow = NativeMethods.GetForegroundWindow();
        _palette.LoadConfig(_config, _lastMode);
        _palette.ShowNearCursor();
    }

    private async Task ChangeProfileAsync(string profileName)
    {
        if (profileName == _config.ActiveProfileName || _busy) return;
        _config.SetActiveProfile(profileName);
        _config.Save();
        await _server.ProfileChangedAsync();
        UpdateProfileChecks();
        _palette.LoadConfig(_config, _lastMode);
        _tray.Text = BuildTrayText();
    }

    private void SetPreview(bool enabled)
    {
        _config.PreviewEnabled = enabled;
        _config.Save();
        if (_previewItem is not null) _previewItem.Checked = enabled;
        _palette.LoadConfig(_config, _lastMode);
    }

    private async Task FormatSelectionAsync(string mode, IntPtr targetWindow)
    {
        if (_busy) return;
        _busy = true;
        _lastMode = mode;
        _activeRequest = new CancellationTokenSource();
        var token = _activeRequest.Token;
        ClipboardSnapshot? snapshot = null;
        var restored = false;
        SetBusyUi(true);

        try
        {
            ShowProgress("Reading selection");
            snapshot = ClipboardSnapshot.Capture();
            Clipboard.Clear();
            var selectedText = await SelectionBridge.CopySelectionAsync(targetWindow, _config.CopyWaitMs, token);
            if (string.IsNullOrWhiteSpace(selectedText))
                throw new InvalidOperationException("No editable text is selected.");

            SetProgressStage("Starting local model");
            await _server.EnsureReadyAsync(_config, token);
            _server.MarkFormatting(true, _config);
            SetProgressStage("Formatting selection");
            var result = await _llamaClient.FormatAsync(_config, mode, selectedText, token);
            selectedText = string.Empty;
            _server.MarkFormatting(false, _config);
            _progress.Finish();

            var output = result.Text.Replace("\r\n", "\n").Replace("\n", Environment.NewLine);
            if (string.IsNullOrWhiteSpace(output)) throw new InvalidDataException("llama.cpp returned no usable replacement text.");

            if (_config.PreviewEnabled)
            {
                var modeConfig = _config.Modes.First(m => m.Name == mode);
                var telemetry = BuildTelemetry(result);
                using var preview = new PreviewForm(output, modeConfig.Label, _config.ActiveProfile.Label, telemetry, _theme, SelectionBridge.SetClipboardText) { Icon = _appIcon };
                var previewResult = preview.ShowPreview();
                output = previewResult.Text;
                if (!previewResult.Replace) return;
                if (string.IsNullOrWhiteSpace(output)) throw new InvalidOperationException("The replacement is empty.");
            }

            SetProgressStage("Replacing selection");
            await SelectionBridge.PasteAsync(targetWindow, output, _config.PasteWaitMs, token);
            snapshot.Restore();
            restored = true;
            _diagnostics.Add($"Completed {mode}: {result.TokensPredicted} output tokens, {result.DecodeTps:0.0} decode TPS.");
        }
        catch (OperationCanceledException)
        {
            _diagnostics.Add("Formatting canceled.");
        }
        catch (Exception ex)
        {
            _diagnostics.Add($"Formatting error: {ex.Message}");
            if (_server.Status.State is ServerState.Starting or ServerState.Formatting)
                _server.MarkError(ex.Message);
            _progressOverlayIsWorking = false;
            if (_config.ProgressOverlayEnabled) _progress.ShowError(ex.Message);
            else MessageBox.Show(ex.Message, "Text Assistant", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        finally
        {
            if (!restored) snapshot?.Restore();
            if (_server.Status.State == ServerState.Formatting) _server.MarkFormatting(false, _config);
            if (_progress.Visible && _progressOverlayIsWorking) _progress.Finish();
            _activeRequest.Dispose();
            _activeRequest = null;
            _busy = false;
            SetBusyUi(false);
        }
    }

    private bool _progressOverlayIsWorking = true;

    private void ShowProgress(string stage)
    {
        _progressOverlayIsWorking = true;
        if (_config.ProgressOverlayEnabled) _progress.Begin(stage);
    }

    private void SetProgressStage(string stage)
    {
        if (_config.ProgressOverlayEnabled) _progress.SetStage(stage);
    }

    private void CancelOrCloseProgress()
    {
        if (_activeRequest is not null) _activeRequest.Cancel();
        else _progress.Finish();
    }

    private string BuildTelemetry(LlamaResult result)
    {
        var parts = new List<string> { result.Endpoint, $"{result.Elapsed.TotalSeconds:0.0}s" };
        if (result.TokensPredicted > 0) parts.Add($"{result.TokensPredicted} tokens");
        if (result.DecodeTps > 0) parts.Add($"{result.DecodeTps:0.0} TPS");
        if (result.PromptTps > 0) parts.Add($"prompt {result.PromptTps:0} TPS");
        if (result.DraftTokens > 0)
            parts.Add($"MTP {100.0 * result.AcceptedDraftTokens / result.DraftTokens:0}% ({result.AcceptedDraftTokens}/{result.DraftTokens})");
        parts.Add(result.DraftTokens > 0 || _server.MtpActive ? "MTP active" : _config.ActiveSpeculative is not null && AppConfig.GetBool(_config.ActiveSpeculative, "enabled", false) ? "standard decoding" : "local");
        return string.Join("  |  ", parts);
    }

    private async Task RestartServerAsync()
    {
        if (_busy) return;
        _activeRequest = new CancellationTokenSource();
        SetBusyUi(true);
        ShowProgress("Restarting local model");
        try
        {
            await _server.RestartAsync(_config, _activeRequest.Token);
            _progress.Finish();
        }
        catch (Exception ex) { _progress.ShowError(ex.Message); }
        finally
        {
            _activeRequest.Dispose();
            _activeRequest = null;
            _busy = false;
            SetBusyUi(false);
        }
    }

    private void OpenSettings()
    {
        if (_busy) return;
        var oldProfile = _config.ActiveProfileName;
        using var settings = new SettingsForm(_config, _diagnostics, _theme) { Icon = _appIcon };
        settings.ShowDialog();
        if (!settings.ConfigurationChanged) return;
        _theme = ThemeManager.Resolve(_config.Theme);
        var oldIcon = _appIcon;
        _appIcon = ThemeManager.CreateAppIcon(_theme);
        _tray.Icon = _appIcon;
        _palette.Icon = _appIcon;
        _progress.Icon = _appIcon;
        _palette.ApplyTheme(_theme);
        _progress.ApplyTheme(_theme);
        oldIcon.Dispose();
        BuildTrayMenu();
        RegisterHotkeys();
        _palette.LoadConfig(_config, _lastMode);
        if (_hotkeyIssues.Count > 0)
            MessageBox.Show("Some shortcuts could not be registered:" + Environment.NewLine + Environment.NewLine + string.Join(Environment.NewLine, _hotkeyIssues), "Shortcut conflicts", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        if (oldProfile != _config.ActiveProfileName) _ = _server.ProfileChangedAsync();
    }

    private void SetBusyUi(bool busy)
    {
        if (_cancelItem is not null) _cancelItem.Visible = busy;
        if (_profilesMenu is not null) _profilesMenu.Enabled = !busy;
    }

    private void OnServerStatusChanged(ServerStatus status)
    {
        if (_shuttingDown) return;
        if (_palette.InvokeRequired) { _palette.BeginInvoke(() => OnServerStatusChanged(status)); return; }
        _palette.SetStatus(status);
        if (_serverStatusItem is not null) _serverStatusItem.Text = status.Detail;
        _tray.Text = BuildTrayText();
    }

    private string BuildTrayText()
    {
        var text = $"Text Assistant - {_config.ActiveProfile.Label}";
        return text.Length <= 63 ? text : text[..63];
    }

    private void UpdateProfileChecks()
    {
        if (_profilesMenu is null) return;
        foreach (ToolStripMenuItem item in _profilesMenu.DropDownItems)
            item.Checked = (string?)item.Tag == _config.ActiveProfileName;
    }

    private void Shutdown()
    {
        if (_shuttingDown) return;
        _shuttingDown = true;
        _activeRequest?.Cancel();
        _tray.Visible = false;
        _hotkeys.Dispose();
        _palette.Dispose();
        _progress.Dispose();
        _trayMenu.Dispose();
        _tray.Dispose();
        _server.Dispose();
        _llamaClient.Dispose();
        _appIcon.Dispose();
        ExitThread();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing && !_shuttingDown) Shutdown();
        base.Dispose(disposing);
    }
}

internal sealed class AssistantColorTable(AppTheme theme) : ProfessionalColorTable
{
    public override Color ToolStripDropDownBackground => theme.Surface;
    public override Color MenuItemSelected => theme.Selection;
    public override Color MenuItemBorder => theme.Border;
    public override Color MenuBorder => theme.Border;
    public override Color ImageMarginGradientBegin => theme.Surface;
    public override Color ImageMarginGradientMiddle => theme.Surface;
    public override Color ImageMarginGradientEnd => theme.Surface;
    public override Color SeparatorDark => theme.Border;
    public override Color SeparatorLight => theme.Border;
}
