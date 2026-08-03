using System.IO;
using System.Windows;
using System.Windows.Threading;
using EmailManage.Services;

namespace EmailManage;

/// <summary>
/// Application entry point. Initializes diagnostics and handles startup/shutdown.
/// </summary>
public partial class App : Application
{
    protected override async void OnStartup(StartupEventArgs e)
    {
        // Global exception handlers
        DispatcherUnhandledException += App_DispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

        try
        {
            // Initialize diagnostic logger BEFORE creating any windows
            DiagnosticLogger.Initialize();
            DiagnosticLogger.Instance.Info("EmailManage starting up...");

            base.OnStartup(e);

            if (e.Args.Any(arg => string.Equals(arg, "--sync", StringComparison.OrdinalIgnoreCase)))
            {
                await RunHeadlessSyncAsync(ParseDateArgument(e.Args, "--sync-from"), ParseDateArgument(e.Args, "--sync-to"), ParseStringArgument(e.Args, "--sync-account"));
                return;
            }

            // Create and show main window manually to catch XAML errors
            var mainWindow = new SearchWindow();
            mainWindow.Show();
            DiagnosticLogger.Instance.Info("MainWindow shown successfully.");
        }
        catch (Exception ex)
        {
            LogCrash("OnStartup", ex);
            MessageBox.Show($"Startup error:\n\n{ex}", "EmailManage - Startup Error",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static async Task RunHeadlessSyncAsync(DateTime? fromUtc, DateTime? toUtc, string? account)
    {
        var code = await new HeadlessSyncRunner().RunAsync(fromUtc, toUtc, account);
        DiagnosticLogger.Instance.Info("Headless sync completed with exit code {ExitCode}", (int)code);
        Current.Shutdown((int)code);
    }

    private static DateTime? ParseDateArgument(string[] args, string name)
    {
        var argument = args.FirstOrDefault(arg => arg.StartsWith(name + "=", StringComparison.OrdinalIgnoreCase));
        var value = argument is null ? null : argument.Substring(name.Length + 1);
        return DateTime.TryParse(value, out var parsed) ? DateTime.SpecifyKind(parsed, DateTimeKind.Local).ToUniversalTime() : null;
    }

    private static string? ParseStringArgument(string[] args, string name)
    {
        var argument = args.FirstOrDefault(arg => arg.StartsWith(name + "=", StringComparison.OrdinalIgnoreCase));
        return argument is null ? null : argument.Substring(name.Length + 1);
    }

    private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        LogCrash("DispatcherUnhandled", e.Exception);
        MessageBox.Show($"Unexpected error:\n\n{e.Exception.Message}\n\nSee crash.log for details.",
            "EmailManage - Error", MessageBoxButton.OK, MessageBoxImage.Error);
        e.Handled = true;
    }

    private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
            LogCrash("AppDomainUnhandled", ex);
    }

    private static void LogCrash(string context, Exception ex)
    {
        try
        {
            var crashLog = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "EmailManage", "crash.log");
            Directory.CreateDirectory(Path.GetDirectoryName(crashLog)!);
            File.AppendAllText(crashLog,
                $"\n[{DateTime.UtcNow:O}] {context}: {ex}\n");
        }
        catch { /* last resort - nothing we can do */ }

        try
        {
            DiagnosticLogger.Instance.Fatal("CRASH [{Context}]: {Error}", context, ex.ToString());
        }
        catch { /* logger may not be initialized */ }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        try
        {
            DiagnosticLogger.Instance.Info("EmailManage shutting down.");
            DiagnosticLogger.Instance.Dispose();
        }
        catch { /* ignore */ }
        base.OnExit(e);
    }
}

