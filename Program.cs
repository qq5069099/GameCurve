using GameCurve.Ui;

namespace GameCurve;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        if (args.Length > 0 && args[0] == "--test")
        {
            Environment.Exit(TestHarness.Run(args));
        }
        Application.SetHighDpiMode(HighDpiMode.SystemAware);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        string? startup = args.Length > 0 && File.Exists(args[0]) ? args[0] : null;
        Application.Run(new MainForm(startup));
    }
}
