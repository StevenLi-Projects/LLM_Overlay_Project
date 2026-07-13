using System.ComponentModel;
using System.Runtime.InteropServices;

namespace LocalTextFormattingAssistant;

[Flags]
internal enum HotkeyModifiers : uint
{
    Alt = 0x0001,
    Control = 0x0002,
    Shift = 0x0004,
    Win = 0x0008,
    NoRepeat = 0x4000
}

internal sealed record ParsedHotkey(HotkeyModifiers Modifiers, Keys Key);

internal static class HotkeyParser
{
    public static ParsedHotkey Parse(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new FormatException("Hotkey is empty.");

        var modifiers = HotkeyModifiers.NoRepeat;
        string? keyName = null;
        foreach (var raw in value.Split('+', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            switch (raw.ToLowerInvariant())
            {
                case "ctrl":
                case "control": modifiers |= HotkeyModifiers.Control; break;
                case "alt": modifiers |= HotkeyModifiers.Alt; break;
                case "shift": modifiers |= HotkeyModifiers.Shift; break;
                case "win": modifiers |= HotkeyModifiers.Win; break;
                default: keyName = raw; break;
            }
        }

        if (keyName is null || !Enum.TryParse<Keys>(keyName, true, out var key))
            throw new FormatException($"Invalid hotkey '{value}'.");
        return new ParsedHotkey(modifiers, key);
    }
}

internal sealed class HotkeyWindow : NativeWindow, IDisposable
{
    private const int WmHotkey = 0x0312;
    private readonly HashSet<int> _ids = [];

    public HotkeyWindow()
    {
        CreateHandle(new CreateParams { Caption = "LocalTextFormattingAssistant.Hotkeys" });
    }

    public event Action<int>? HotkeyPressed;

    public void Register(int id, ParsedHotkey hotkey)
    {
        if (!NativeMethods.RegisterHotKey(Handle, id, (uint)hotkey.Modifiers, (uint)hotkey.Key))
            throw new Win32Exception(Marshal.GetLastWin32Error(), $"Could not register hotkey {id}.");
        _ids.Add(id);
    }

    public void Clear()
    {
        foreach (var id in _ids)
            NativeMethods.UnregisterHotKey(Handle, id);
        _ids.Clear();
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WmHotkey)
            HotkeyPressed?.Invoke(m.WParam.ToInt32());
        base.WndProc(ref m);
    }

    public void Dispose()
    {
        Clear();
        DestroyHandle();
    }
}

internal static class NativeMethods
{
    [DllImport("user32.dll", SetLastError = true)]
    internal static extern bool RegisterHotKey(IntPtr hWnd, int id, uint modifiers, uint key);

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [DllImport("user32.dll")]
    internal static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    internal static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int virtualKey);

    [DllImport("user32.dll")]
    internal static extern bool DestroyIcon(IntPtr handle);

    [DllImport("kernel32.dll", SetLastError = true)]
    internal static extern bool AttachConsole(uint processId);

    internal static bool IsKeyDown(Keys key) => (GetAsyncKeyState((int)key) & 0x8000) != 0;
}

internal sealed class ClipboardSnapshot
{
    private readonly IDataObject? _data;

    private ClipboardSnapshot(IDataObject? data) => _data = data;

    public static ClipboardSnapshot Capture()
    {
        for (var attempt = 0; attempt < 8; attempt++)
        {
            try { return new ClipboardSnapshot(Clipboard.GetDataObject()); }
            catch { Thread.Sleep(40); }
        }
        return new ClipboardSnapshot(null);
    }

    public void Restore()
    {
        if (_data is null) return;
        for (var attempt = 0; attempt < 8; attempt++)
        {
            try
            {
                Clipboard.SetDataObject(_data, true);
                return;
            }
            catch { Thread.Sleep(60); }
        }
    }
}

internal static class SelectionBridge
{
    public static async Task<string?> CopySelectionAsync(IntPtr targetWindow, int waitMs, CancellationToken token)
    {
        if (targetWindow != IntPtr.Zero)
        {
            NativeMethods.SetForegroundWindow(targetWindow);
            await Task.Delay(90, token);
        }
        await WaitForModifiersAsync(token);
        SendKeys.SendWait("^c");
        await Task.Delay(waitMs, token);

        for (var attempt = 0; attempt < 8; attempt++)
        {
            try
            {
                if (Clipboard.ContainsText())
                    return Clipboard.GetText();
            }
            catch { }
            await Task.Delay(35, token);
        }
        return null;
    }

    public static async Task PasteAsync(IntPtr targetWindow, string text, int waitMs, CancellationToken token)
    {
        SetClipboardText(text);
        if (targetWindow != IntPtr.Zero)
        {
            NativeMethods.SetForegroundWindow(targetWindow);
            await Task.Delay(90, token);
        }
        await WaitForModifiersAsync(token);
        SendKeys.SendWait("^v");
        await Task.Delay(waitMs, token);
    }

    public static void SetClipboardText(string text)
    {
        for (var attempt = 0; attempt < 8; attempt++)
        {
            try
            {
                Clipboard.SetText(text);
                return;
            }
            catch { Thread.Sleep(40); }
        }
        throw new InvalidOperationException("The Windows clipboard is busy.");
    }

    private static async Task WaitForModifiersAsync(CancellationToken token)
    {
        var deadline = DateTime.UtcNow.AddSeconds(2);
        while (DateTime.UtcNow < deadline &&
               (NativeMethods.IsKeyDown(Keys.ControlKey) || NativeMethods.IsKeyDown(Keys.Menu) || NativeMethods.IsKeyDown(Keys.ShiftKey) || NativeMethods.IsKeyDown(Keys.LWin) || NativeMethods.IsKeyDown(Keys.RWin)))
            await Task.Delay(20, token);
    }
}
