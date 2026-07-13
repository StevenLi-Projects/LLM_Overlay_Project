namespace LocalTextFormattingAssistant;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        try
        {
            var configPath = FindConfigPath(args);
            var config = AppConfig.Load(configPath);
            if (args.Contains("--validate", StringComparer.OrdinalIgnoreCase))
            {
                AttachParentConsole();
                return Validate(config).GetAwaiter().GetResult();
            }

            using var mutex = new Mutex(true, "LocalTextFormattingAssistant.Compiled", out var createdNew);
            if (!createdNew)
            {
                MessageBox.Show("The Text Assistant is already running. Check the system tray.", "Text Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 2;
            }

            using var context = new AssistantApplicationContext(config);
            Application.Run(context);
            try { mutex.ReleaseMutex(); } catch { }
            return 0;
        }
        catch (Exception ex)
        {
            if (args.Contains("--validate", StringComparer.OrdinalIgnoreCase))
                Console.Error.WriteLine(ex.Message);
            else
                MessageBox.Show(ex.Message, "Text Assistant startup error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
    }

    private static async Task<int> Validate(AppConfig config)
    {
        var errors = await ConfigValidator.ValidateAsync(config, CancellationToken.None);
        Console.WriteLine($"[OK] config.json: {config.ConfigPath}");
        Console.WriteLine($"[OK] llama-server.exe: {config.ServerExecutable}");
        Console.WriteLine($"[OK] active profile: {config.ActiveProfile.Label}");
        foreach (var mode in config.Modes.Where(m => m.Enabled)) Console.WriteLine($"[OK] hotkey {mode.Hotkey}: {mode.Label}");
        if (errors.Count == 0)
        {
            Console.WriteLine("Validation completed.");
            return 0;
        }
        foreach (var error in errors) Console.Error.WriteLine($"[FAIL] {error}");
        return 1;
    }

    private static void AttachParentConsole()
    {
        const uint attachParentProcess = 0xFFFFFFFF;
        NativeMethods.AttachConsole(attachParentProcess);
        Console.SetOut(new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true });
        Console.SetError(new StreamWriter(Console.OpenStandardError()) { AutoFlush = true });
    }

    private static string FindConfigPath(string[] args)
    {
        for (var i = 0; i < args.Length - 1; i++)
            if (args[i].Equals("--config", StringComparison.OrdinalIgnoreCase))
                return Path.GetFullPath(args[i + 1]);

        var candidates = new[]
        {
            Path.Combine(AppContext.BaseDirectory, "config.json"),
            Path.Combine(AppContext.BaseDirectory, "..", "config.json"),
            Path.Combine(Environment.CurrentDirectory, "config.json")
        };
        return candidates.Select(Path.GetFullPath).FirstOrDefault(File.Exists)
            ?? throw new FileNotFoundException("config.json was not found. Launch with --config <path>.");
    }
}
