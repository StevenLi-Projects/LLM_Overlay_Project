using System.Drawing.Drawing2D;
using System.Text.Json.Nodes;
using Microsoft.Win32;

namespace LocalTextFormattingAssistant;

internal sealed record AppTheme(
    Color Window,
    Color Surface,
    Color SurfaceAlt,
    Color Text,
    Color Muted,
    Color Border,
    Color Accent,
    Color AccentHover,
    Color AccentText,
    Color Danger,
    Color Selection);

internal static class ThemeManager
{
    public static AppTheme Resolve(string preference)
    {
        var dark = preference.Equals("dark", StringComparison.OrdinalIgnoreCase) ||
                   preference.Equals("system", StringComparison.OrdinalIgnoreCase) && IsSystemDark();
        return dark
            ? new AppTheme(
                Color.FromArgb(22, 24, 28), Color.FromArgb(30, 33, 38), Color.FromArgb(39, 43, 49),
                Color.FromArgb(239, 242, 246), Color.FromArgb(156, 163, 175), Color.FromArgb(62, 68, 77),
                Color.FromArgb(52, 168, 112), Color.FromArgb(64, 184, 126), Color.White,
                Color.FromArgb(232, 91, 91), Color.FromArgb(43, 76, 64))
            : new AppTheme(
                Color.FromArgb(247, 249, 251), Color.White, Color.FromArgb(240, 243, 246),
                Color.FromArgb(25, 30, 36), Color.FromArgb(102, 112, 125), Color.FromArgb(214, 220, 227),
                Color.FromArgb(22, 132, 86), Color.FromArgb(18, 117, 76), Color.White,
                Color.FromArgb(194, 55, 55), Color.FromArgb(224, 242, 234));
    }

    private static bool IsSystemDark()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            return Convert.ToInt32(key?.GetValue("AppsUseLightTheme", 1)) == 0;
        }
        catch { return false; }
    }

    public static void StyleButton(Button button, AppTheme theme, bool primary = false, bool danger = false)
    {
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = primary ? 0 : 1;
        button.FlatAppearance.BorderColor = theme.Border;
        button.BackColor = primary ? theme.Accent : danger ? theme.Danger : theme.SurfaceAlt;
        button.ForeColor = primary || danger ? theme.AccentText : theme.Text;
        button.Font = new Font("Segoe UI Semibold", 9F);
        button.Cursor = Cursors.Hand;
    }

    public static Icon CreateAppIcon(AppTheme theme)
    {
        using var bitmap = new Bitmap(32, 32);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.Clear(Color.Transparent);
        using var background = new SolidBrush(theme.Accent);
        graphics.FillRoundedRectangle(background, new RectangleF(2, 2, 28, 28), 7);
        using var pen = new Pen(Color.White, 2.2F) { StartCap = LineCap.Round, EndCap = LineCap.Round };
        graphics.DrawLine(pen, 9, 11, 23, 11);
        graphics.DrawLine(pen, 9, 16, 20, 16);
        graphics.DrawLine(pen, 9, 21, 17, 21);
        var handle = bitmap.GetHicon();
        try { return (Icon)Icon.FromHandle(handle).Clone(); }
        finally { NativeMethods.DestroyIcon(handle); }
    }

    private static void FillRoundedRectangle(this Graphics graphics, Brush brush, RectangleF bounds, float radius)
    {
        using var path = new GraphicsPath();
        var diameter = radius * 2;
        path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
        path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
        path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        graphics.FillPath(brush, path);
    }
}

internal sealed class ModeListItem
{
    public ModeListItem(ModeConfig mode) => Mode = mode;
    public ModeConfig Mode { get; }
    public override string ToString() => Mode.Label;
}

internal sealed class CommandPaletteForm : Form
{
    private readonly ListBox _modes = new();
    private readonly ComboBox _profiles = new();
    private readonly CheckBox _preview = new();
    private readonly Label _status = new();
    private readonly Panel _statusDot = new();
    private AppTheme _theme;
    private bool _suppressProfileEvent;
    private bool _suppressPreviewEvent;

    public CommandPaletteForm(AppConfig config, AppTheme theme)
    {
        _theme = theme;
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        KeyPreview = true;
        Width = 460;
        Height = 420;
        MinimumSize = new Size(420, 390);
        Padding = new Padding(12);

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, ColumnCount = 1 };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));

        var header = new Panel { Dock = DockStyle.Fill };
        var title = new Label { Text = "Format selection", AutoSize = true, Location = new Point(2, 3), Font = new Font("Segoe UI Semibold", 13F) };
        var caption = new Label { Text = "Choose a transformation", AutoSize = true, Location = new Point(3, 29), Font = new Font("Segoe UI", 8.5F) };
        header.Controls.Add(title);
        header.Controls.Add(caption);

        _modes.Dock = DockStyle.Fill;
        _modes.BorderStyle = BorderStyle.None;
        _modes.DrawMode = DrawMode.OwnerDrawFixed;
        _modes.ItemHeight = 48;
        _modes.IntegralHeight = false;
        _modes.DrawItem += DrawModeItem;
        _modes.DoubleClick += (_, _) => ActivateSelected();

        var controls = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(0, 8, 0, 4) };
        controls.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 52));
        controls.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 48));
        _profiles.Dock = DockStyle.Fill;
        _profiles.DropDownStyle = ComboBoxStyle.DropDownList;
        _profiles.SelectedIndexChanged += (_, _) =>
        {
            if (!_suppressProfileEvent && _profiles.SelectedItem is ProfileConfig profile)
                ProfileSelected?.Invoke(profile.Name);
        };
        _preview.Text = "Preview replacement";
        _preview.AutoSize = true;
        _preview.Anchor = AnchorStyles.Right;
        _preview.CheckedChanged += (_, _) =>
        {
            if (!_suppressPreviewEvent) PreviewChanged?.Invoke(_preview.Checked);
        };
        controls.Controls.Add(_profiles, 0, 0);
        controls.Controls.Add(_preview, 1, 0);

        var statusPanel = new Panel { Dock = DockStyle.Fill };
        _statusDot.Size = new Size(8, 8);
        _statusDot.Location = new Point(3, 10);
        _statusDot.Paint += (_, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using var brush = new SolidBrush(_statusDot.BackColor);
            e.Graphics.FillEllipse(brush, 0, 0, 8, 8);
        };
        _status.AutoSize = false;
        _status.Location = new Point(17, 3);
        _status.Size = new Size(375, 24);
        _status.TextAlign = ContentAlignment.MiddleLeft;
        statusPanel.Controls.Add(_statusDot);
        statusPanel.Controls.Add(_status);

        layout.Controls.Add(header, 0, 0);
        layout.Controls.Add(_modes, 0, 1);
        layout.Controls.Add(controls, 0, 2);
        layout.Controls.Add(statusPanel, 0, 3);
        Controls.Add(layout);

        LoadConfig(config, null);
        ApplyTheme(theme);
        Paint += (_, e) =>
        {
            using var pen = new Pen(_theme.Border);
            e.Graphics.DrawRectangle(pen, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
        };
        Deactivate += (_, _) => Hide();
        KeyDown += HandleKeyDown;
    }

    public event Action<string>? ModeSelected;
    public event Action<string>? ProfileSelected;
    public event Action<bool>? PreviewChanged;

    public void LoadConfig(AppConfig config, string? lastMode)
    {
        _modes.Items.Clear();
        foreach (var mode in config.Modes.Where(m => m.Enabled)) _modes.Items.Add(new ModeListItem(mode));
        if (_modes.Items.Count > 0)
        {
            var index = Enumerable.Range(0, _modes.Items.Count).FirstOrDefault(i => ((ModeListItem)_modes.Items[i]).Mode.Name == lastMode);
            _modes.SelectedIndex = index;
        }

        _suppressProfileEvent = true;
        _profiles.Items.Clear();
        foreach (var profile in config.Profiles) _profiles.Items.Add(profile);
        _profiles.DisplayMember = nameof(ProfileConfig.Label);
        _profiles.SelectedItem = _profiles.Items.Cast<ProfileConfig>().FirstOrDefault(p => p.Name == config.ActiveProfileName);
        if (_profiles.SelectedIndex < 0 && _profiles.Items.Count > 0) _profiles.SelectedIndex = 0;
        _suppressPreviewEvent = true;
        _preview.Checked = config.PreviewEnabled;
        _suppressPreviewEvent = false;
        _suppressProfileEvent = false;
    }

    public void SetStatus(ServerStatus status)
    {
        _status.Text = status.Detail;
        _statusDot.BackColor = status.State switch
        {
            ServerState.Ready => _theme.Accent,
            ServerState.Formatting or ServerState.Starting => Color.FromArgb(212, 151, 45),
            ServerState.Error => _theme.Danger,
            _ => _theme.Muted
        };
        _statusDot.Invalidate();
    }

    public void ShowNearCursor()
    {
        var screen = Screen.FromPoint(Cursor.Position).WorkingArea;
        var x = Math.Min(Cursor.Position.X, screen.Right - Width - 8);
        var y = Math.Min(Cursor.Position.Y + 8, screen.Bottom - Height - 8);
        Location = new Point(Math.Max(screen.Left + 8, x), Math.Max(screen.Top + 8, y));
        Show();
        Activate();
        _modes.Focus();
    }

    public void ApplyTheme(AppTheme theme)
    {
        _theme = theme;
        BackColor = theme.Window;
        ForeColor = theme.Text;
        foreach (Control control in Controls) ApplyColors(control, theme);
        _modes.BackColor = theme.Surface;
        _modes.ForeColor = theme.Text;
        _profiles.BackColor = theme.Surface;
        _profiles.ForeColor = theme.Text;
        _modes.Invalidate();
    }

    private static void ApplyColors(Control parent, AppTheme theme)
    {
        parent.BackColor = theme.Window;
        parent.ForeColor = parent is Label label && label.Font.Size < 10 ? theme.Muted : theme.Text;
        foreach (Control child in parent.Controls) ApplyColors(child, theme);
    }

    private void DrawModeItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0) return;
        var item = (ModeListItem)_modes.Items[e.Index];
        var selected = (e.State & DrawItemState.Selected) != 0;
        using var backgroundBrush = new SolidBrush(selected ? _theme.Selection : _theme.Surface);
        e.Graphics.FillRectangle(backgroundBrush, e.Bounds);
        var badge = new Rectangle(e.Bounds.Left + 10, e.Bounds.Top + 10, 28, 28);
        using var badgeBrush = new SolidBrush(selected ? _theme.Accent : _theme.SurfaceAlt);
        e.Graphics.FillEllipse(badgeBrush, badge);
        using var badgeFont = new Font("Segoe UI Semibold", 9F);
        using var textFont = new Font("Segoe UI Semibold", 9.5F);
        using var hotkeyFont = new Font("Segoe UI", 8.5F);
        var badgeText = item.Mode.Name switch { "markdown" => "M", "bullets" => "B", "table" => "T", "cleanup" => "C", "summary" => "S", _ => item.Mode.Label[..1].ToUpperInvariant() };
        TextRenderer.DrawText(e.Graphics, badgeText, badgeFont, badge, selected ? _theme.AccentText : _theme.Muted, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        TextRenderer.DrawText(e.Graphics, item.Mode.Label, textFont, new Point(e.Bounds.Left + 50, e.Bounds.Top + 8), _theme.Text);
        TextRenderer.DrawText(e.Graphics, item.Mode.Hotkey, hotkeyFont, new Rectangle(e.Bounds.Right - 130, e.Bounds.Top + 14, 118, 22), _theme.Muted, TextFormatFlags.Right);
    }

    private void HandleKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Escape) { Hide(); e.Handled = true; }
        else if (e.KeyCode == Keys.Enter) { ActivateSelected(); e.Handled = true; }
    }

    private void ActivateSelected()
    {
        if (_modes.SelectedItem is not ModeListItem item) return;
        Hide();
        ModeSelected?.Invoke(item.Mode.Name);
    }
}

internal sealed class ProgressOverlayForm : Form
{
    private readonly Label _stage = new();
    private readonly Label _elapsed = new();
    private readonly Button _cancel = new();
    private readonly ProgressBar _bar = new();
    private readonly System.Windows.Forms.Timer _timer = new() { Interval = 250 };
    private DateTime _started;

    public ProgressOverlayForm(AppTheme theme)
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        Size = new Size(360, 104);
        Padding = new Padding(14);

        _stage.Location = new Point(14, 12);
        _stage.Size = new Size(260, 24);
        _stage.Font = new Font("Segoe UI Semibold", 10F);
        _elapsed.Location = new Point(278, 12);
        _elapsed.Size = new Size(66, 24);
        _elapsed.TextAlign = ContentAlignment.MiddleRight;
        _bar.Location = new Point(14, 48);
        _bar.Size = new Size(258, 8);
        _bar.Style = ProgressBarStyle.Marquee;
        _bar.MarqueeAnimationSpeed = 22;
        _cancel.Text = "Cancel";
        _cancel.Location = new Point(280, 40);
        _cancel.Size = new Size(66, 28);
        _cancel.Click += (_, _) => CancelRequested?.Invoke();
        Controls.AddRange([_stage, _elapsed, _bar, _cancel]);
        _timer.Tick += (_, _) => _elapsed.Text = $"{(DateTime.UtcNow - _started).TotalSeconds:0.0}s";
        ApplyTheme(theme);
    }

    public event Action? CancelRequested;

    public void Begin(string stage)
    {
        _started = DateTime.UtcNow;
        _stage.Text = stage;
        _cancel.Text = "Cancel";
        _cancel.Enabled = true;
        _bar.Visible = true;
        _timer.Start();
        PositionOnActiveScreen();
        Show();
        Activate();
    }

    public void SetStage(string stage) => _stage.Text = stage;

    public void ShowError(string message)
    {
        _timer.Stop();
        _stage.Text = message;
        _elapsed.Text = string.Empty;
        _bar.Visible = false;
        _cancel.Text = "Close";
        _cancel.Enabled = true;
        if (!Visible) { PositionOnActiveScreen(); Show(); }
    }

    public void Finish()
    {
        _timer.Stop();
        Hide();
    }

    public void ApplyTheme(AppTheme theme)
    {
        BackColor = theme.Surface;
        ForeColor = theme.Text;
        _stage.ForeColor = theme.Text;
        _elapsed.ForeColor = theme.Muted;
        ThemeManager.StyleButton(_cancel, theme);
    }

    private void PositionOnActiveScreen()
    {
        var area = Screen.FromPoint(Cursor.Position).WorkingArea;
        Location = new Point(area.Right - Width - 16, area.Bottom - Height - 16);
    }
}

internal sealed record PreviewResult(bool Replace, string Text);

internal sealed class PreviewForm : Form
{
    private readonly TextBox _replacement = new();
    private readonly Label _telemetry = new();
    private PreviewResult _result;

    public PreviewForm(string text, string modeLabel, string profileLabel, string telemetry, AppTheme theme, Action<string> copyAction)
    {
        Text = "Preview replacement";
        StartPosition = FormStartPosition.CenterScreen;
        Size = new Size(700, 470);
        MinimumSize = new Size(520, 340);
        TopMost = true;
        KeyPreview = true;
        Padding = new Padding(14);
        _result = new PreviewResult(false, text);

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, ColumnCount = 1 };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));

        var header = new Panel { Dock = DockStyle.Fill };
        var title = new Label { Text = "Replacement", AutoSize = true, Font = new Font("Segoe UI Semibold", 12F), Location = new Point(1, 4) };
        var meta = new Label { Text = $"{modeLabel} | {profileLabel}", AutoSize = true, Font = new Font("Segoe UI", 8.5F), Location = new Point(2, 27) };
        header.Controls.Add(title);
        header.Controls.Add(meta);

        _replacement.Multiline = true;
        _replacement.AcceptsReturn = true;
        _replacement.AcceptsTab = true;
        _replacement.ScrollBars = ScrollBars.Vertical;
        _replacement.WordWrap = true;
        _replacement.Dock = DockStyle.Fill;
        _replacement.Font = new Font("Cascadia Mono", 9.5F);
        _replacement.Text = text.Replace("\r\n", "\n").Replace("\n", Environment.NewLine);

        _telemetry.Text = telemetry;
        _telemetry.Dock = DockStyle.Fill;
        _telemetry.TextAlign = ContentAlignment.MiddleLeft;
        _telemetry.AutoEllipsis = true;
        _telemetry.Font = new Font("Segoe UI", 8.5F);

        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0, 5, 0, 0) };
        var replace = new Button { Text = "Replace", Size = new Size(104, 31) };
        var copy = new Button { Text = "Copy", Size = new Size(88, 31) };
        var cancel = new Button { Text = "Cancel", Size = new Size(88, 31) };
        ThemeManager.StyleButton(replace, theme, primary: true);
        ThemeManager.StyleButton(copy, theme);
        ThemeManager.StyleButton(cancel, theme);
        replace.Click += (_, _) => { _result = new PreviewResult(true, _replacement.Text); DialogResult = DialogResult.OK; Close(); };
        cancel.Click += (_, _) => { _result = new PreviewResult(false, _replacement.Text); DialogResult = DialogResult.Cancel; Close(); };
        copy.Click += (_, _) => { copyAction(_replacement.Text); _telemetry.Text = "Copied. The previous clipboard returns when this window closes."; };
        buttons.Controls.AddRange([replace, copy, cancel]);

        layout.Controls.Add(header, 0, 0);
        layout.Controls.Add(_replacement, 0, 1);
        layout.Controls.Add(_telemetry, 0, 2);
        layout.Controls.Add(buttons, 0, 3);
        Controls.Add(layout);
        AcceptButton = replace;
        CancelButton = cancel;
        KeyDown += (_, e) =>
        {
            if (e.Control && e.KeyCode == Keys.Enter) { replace.PerformClick(); e.Handled = true; }
            else if (e.KeyCode == Keys.Escape) { cancel.PerformClick(); e.Handled = true; }
        };
        ApplyTheme(theme, meta, layout, header, buttons);
    }

    public PreviewResult ShowPreview()
    {
        ShowDialog();
        return _result;
    }

    private void ApplyTheme(AppTheme theme, Label meta, TableLayoutPanel layout, Panel header, FlowLayoutPanel buttons)
    {
        BackColor = theme.Window;
        ForeColor = theme.Text;
        layout.BackColor = theme.Window;
        header.BackColor = theme.Window;
        buttons.BackColor = theme.Window;
        meta.ForeColor = theme.Muted;
        _replacement.BackColor = theme.Surface;
        _replacement.ForeColor = theme.Text;
        _replacement.BorderStyle = BorderStyle.FixedSingle;
        _telemetry.BackColor = theme.SurfaceAlt;
        _telemetry.ForeColor = theme.Muted;
        _telemetry.Padding = new Padding(8, 0, 8, 0);
    }
}

internal sealed class SettingsForm : Form
{
    private readonly AppConfig _config;
    private readonly DiagnosticsBuffer _diagnostics;
    private readonly ComboBox _theme = new();
    private readonly CheckBox _preview = new();
    private readonly CheckBox _progress = new();
    private readonly ComboBox _profile = new();
    private readonly TextBox _modelPath = new();
    private readonly NumericUpDown _context = new();
    private readonly TextBox _cppPath = new();
    private readonly NumericUpDown _port = new();
    private readonly Dictionary<string, TextBox> _hotkeys = [];
    private readonly NumericUpDown _maxTokens = new();
    private readonly NumericUpDown _temperature = new();
    private readonly NumericUpDown _timeout = new();
    private readonly Label _validation = new();
    private readonly AppTheme _appTheme;

    public SettingsForm(AppConfig config, DiagnosticsBuffer diagnostics, AppTheme appTheme)
    {
        _config = config;
        _diagnostics = diagnostics;
        _appTheme = appTheme;
        Text = "Text Assistant settings";
        StartPosition = FormStartPosition.CenterScreen;
        Size = new Size(640, 500);
        MinimumSize = new Size(580, 440);
        KeyPreview = true;

        var tabs = new TabControl { Dock = DockStyle.Fill, Padding = new Point(14, 6) };
        tabs.TabPages.Add(CreateGeneralTab());
        tabs.TabPages.Add(CreateModelsTab());
        tabs.TabPages.Add(CreateShortcutsTab());
        tabs.TabPages.Add(CreateAdvancedTab());

        var footer = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 48, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(8) };
        var save = new Button { Text = "Save", Size = new Size(92, 31) };
        var cancel = new Button { Text = "Cancel", Size = new Size(92, 31), DialogResult = DialogResult.Cancel };
        ThemeManager.StyleButton(save, appTheme, primary: true);
        ThemeManager.StyleButton(cancel, appTheme);
        save.Click += async (_, _) => await SaveAsync();
        footer.Controls.AddRange([save, cancel]);
        Controls.Add(tabs);
        Controls.Add(footer);
        CancelButton = cancel;
        ApplyTheme(this, appTheme);
        Shown += async (_, _) => await UpdateRuntimeStatusAsync();
    }

    public bool ConfigurationChanged { get; private set; }

    private TabPage CreateGeneralTab()
    {
        var page = NewPage("General");
        var table = NewSettingsTable();
        _theme.DropDownStyle = ComboBoxStyle.DropDownList;
        _theme.Items.AddRange(["System", "Light", "Dark"]);
        _theme.SelectedItem = char.ToUpperInvariant(_config.Theme[0]) + _config.Theme[1..];
        _preview.Checked = _config.PreviewEnabled;
        _preview.Text = "Preview before replacing";
        _progress.Checked = _config.ProgressOverlayEnabled;
        _progress.Text = "Show progress overlay";
        AddRow(table, "Theme", _theme);
        AddRow(table, "Replacement", _preview);
        AddRow(table, "Progress", _progress);
        var startup = new Label { Text = "The model starts on the first request to keep GPU memory free.", AutoSize = true };
        AddRow(table, "Startup", startup);
        page.Controls.Add(table);
        return page;
    }

    private TabPage CreateModelsTab()
    {
        var page = NewPage("Models");
        var table = NewSettingsTable();
        _profile.DropDownStyle = ComboBoxStyle.DropDownList;
        foreach (var profile in _config.Profiles) _profile.Items.Add(profile);
        _profile.DisplayMember = nameof(ProfileConfig.Label);
        _profile.SelectedItem = _profile.Items.Cast<ProfileConfig>().FirstOrDefault(p => p.Name == _config.ActiveProfileName);
        _profile.SelectedIndexChanged += (_, _) => LoadSelectedProfile();
        _modelPath.Dock = DockStyle.Fill;
        _context.Minimum = 512;
        _context.Maximum = 262144;
        _context.Increment = 512;
        _context.ThousandsSeparator = true;
        _cppPath.Text = AppConfig.GetString(_config.Llama, "cpp_dir", string.Empty);
        _cppPath.Dock = DockStyle.Fill;
        _port.Minimum = 1;
        _port.Maximum = 65535;
        _port.Value = _config.Port;
        AddRow(table, "Profile", _profile);
        AddRow(table, "Model path", _modelPath);
        AddRow(table, "Context", _context);
        AddRow(table, "llama.cpp", _cppPath);
        AddRow(table, "Port", _port);
        _validation.AutoSize = true;
        AddRow(table, "Runtime", _validation);
        LoadSelectedProfile();
        page.Controls.Add(table);
        return page;
    }

    private TabPage CreateShortcutsTab()
    {
        var page = NewPage("Shortcuts");
        var table = NewSettingsTable();
        foreach (var mode in _config.Modes.Where(m => m.Enabled))
        {
            var input = new TextBox { Text = mode.Hotkey, Dock = DockStyle.Fill };
            _hotkeys[mode.Name] = input;
            AddRow(table, mode.Label, input);
        }
        var menu = new TextBox { Text = _config.MenuHotkey, Dock = DockStyle.Fill, Tag = "menu" };
        _hotkeys["__menu"] = menu;
        AddRow(table, "Open formatter", menu);
        page.Controls.Add(table);
        return page;
    }

    private TabPage CreateAdvancedTab()
    {
        var page = NewPage("Advanced");
        var table = NewSettingsTable();
        _maxTokens.Minimum = 64;
        _maxTokens.Maximum = 32768;
        _maxTokens.Value = AppConfig.GetInt(_config.Generation, "max_tokens", 2048);
        _temperature.DecimalPlaces = 2;
        _temperature.Increment = 0.05M;
        _temperature.Minimum = 0;
        _temperature.Maximum = 2;
        _temperature.Value = (decimal)AppConfig.GetDouble(_config.Generation, "temperature", 0.2);
        _timeout.Minimum = 10;
        _timeout.Maximum = 1800;
        _timeout.Value = _config.TimeoutSeconds;
        AddRow(table, "Maximum output", _maxTokens);
        AddRow(table, "Temperature", _temperature);
        AddRow(table, "Timeout (seconds)", _timeout);
        var copyDiagnostics = new Button { Text = "Copy diagnostics", Size = new Size(140, 31) };
        ThemeManager.StyleButton(copyDiagnostics, _appTheme);
        copyDiagnostics.Click += (_, _) => SelectionBridge.SetClipboardText(_diagnostics.GetText());
        AddRow(table, "Troubleshooting", copyDiagnostics);
        page.Controls.Add(table);
        return page;
    }

    private async Task SaveAsync()
    {
        var errors = ValidateInputs();
        if (errors.Count > 0)
        {
            MessageBox.Show(string.Join(Environment.NewLine, errors), "Check settings", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _config.Theme = _theme.SelectedItem?.ToString()?.ToLowerInvariant() ?? "system";
        _config.PreviewEnabled = _preview.Checked;
        _config.ProgressOverlayEnabled = _progress.Checked;
        if (_profile.SelectedItem is ProfileConfig selected)
        {
            _config.SetActiveProfile(selected.Name);
            selected.Source["model_path"] = _config.ToRelativePath(_config.ResolvePath(_modelPath.Text.Trim()));
            selected.Source["context_size"] = (int)_context.Value;
        }
        _config.Llama["cpp_dir"] = _config.ToRelativePath(_config.ResolvePath(_cppPath.Text.Trim()));
        _config.Llama["port"] = (int)_port.Value;
        _config.Llama["server_url"] = $"http://{_config.Host}:{(int)_port.Value}";
        _config.Generation["max_tokens"] = (int)_maxTokens.Value;
        _config.Generation["temperature"] = (double)_temperature.Value;
        _config.Generation["timeout_sec"] = (int)_timeout.Value;
        foreach (var mode in _config.Modes)
            if (_hotkeys.TryGetValue(mode.Name, out var input) && _config.ModesNode[mode.Name] is JsonObject modeNode)
                modeNode["hotkey"] = input.Text.Trim();
        _config.Ui["menu_hotkey"] = _hotkeys["__menu"].Text.Trim();
        _config.Save();
        ConfigurationChanged = true;
        _validation.Text = "Saved";
        await Task.Delay(180);
        DialogResult = DialogResult.OK;
        Close();
    }

    private List<string> ValidateInputs()
    {
        var errors = new List<string>();
        if (!File.Exists(Path.Combine(_config.ResolvePath(_cppPath.Text.Trim()), "llama-server.exe"))) errors.Add("llama.cpp folder does not contain llama-server.exe.");
        if (!File.Exists(_config.ResolvePath(_modelPath.Text.Trim()))) errors.Add("The selected model file does not exist.");
        if (_profile.SelectedItem is ProfileConfig profile && profile.Source["speculative"] is JsonObject spec && AppConfig.GetBool(spec, "enabled", false))
        {
            var draft = _config.ResolvePath(AppConfig.GetString(spec, "draft_model_path", string.Empty));
            if (!File.Exists(draft)) errors.Add("The selected profile's MTP draft model does not exist.");
        }
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var pair in _hotkeys)
        {
            try { HotkeyParser.Parse(pair.Value.Text.Trim()); }
            catch (Exception ex) { errors.Add($"{pair.Key}: {ex.Message}"); continue; }
            if (!used.Add(pair.Value.Text.Trim())) errors.Add($"Duplicate hotkey: {pair.Value.Text.Trim()}");
        }
        return errors;
    }

    private async Task UpdateRuntimeStatusAsync()
    {
        try
        {
            using var manager = new LlamaServerManager(_diagnostics);
            var runtime = await manager.DetectMtpRuntimeAsync(_config, CancellationToken.None);
            _validation.Text = runtime switch
            {
                MtpRuntimeStyle.Mainline => "MTP ready (llama.cpp draft-mtp)",
                MtpRuntimeStyle.Atomic => "MTP ready (Atomic mtp-head)",
                _ => "MTP unavailable; standard decoding fallback"
            };
            _validation.ForeColor = runtime != MtpRuntimeStyle.None ? _appTheme.Accent : _appTheme.Muted;
        }
        catch (Exception ex)
        {
            _validation.Text = $"Runtime check failed: {ex.Message}";
            _validation.ForeColor = _appTheme.Danger;
        }
    }

    private void LoadSelectedProfile()
    {
        if (_profile.SelectedItem is not ProfileConfig profile) return;
        _modelPath.Text = AppConfig.GetString(profile.Source, "model_path", string.Empty);
        _context.Value = Math.Clamp(profile.ContextSize, (int)_context.Minimum, (int)_context.Maximum);
        var spec = profile.Source["speculative"] as JsonObject;
        var enabled = spec is not null && AppConfig.GetBool(spec, "enabled", false);
        _validation.Text = enabled ? "MTP configured; runtime checked when the server starts." : "Standard decoding";
    }

    private TabPage NewPage(string text) => new(text) { BackColor = _appTheme.Window, ForeColor = _appTheme.Text, Padding = new Padding(12) };

    private static TableLayoutPanel NewSettingsTable()
    {
        var table = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, Padding = new Padding(4) };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 155));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        return table;
    }

    private static void AddRow(TableLayoutPanel table, string label, Control control)
    {
        var row = table.RowCount++;
        table.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
        var caption = new Label { Text = label, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI Semibold", 9F) };
        control.Margin = new Padding(4, 9, 4, 8);
        table.Controls.Add(caption, 0, row);
        table.Controls.Add(control, 1, row);
    }

    private static void ApplyTheme(Control parent, AppTheme theme)
    {
        if (parent is not Button)
        {
            parent.BackColor = parent is TextBox or ComboBox or NumericUpDown ? theme.Surface : theme.Window;
            parent.ForeColor = theme.Text;
        }
        foreach (Control child in parent.Controls)
            ApplyTheme(child, theme);
    }
}
