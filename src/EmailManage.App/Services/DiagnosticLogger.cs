using System.IO;
using Serilog;
using Serilog.Events;

namespace EmailManage.Services;

/// <summary>
/// Structured diagnostic logger using Serilog.
/// Writes to a rolling file under %LOCALAPPDATA%\EmailManage\diagnostic.log.
/// </summary>
public sealed class DiagnosticLogger : IDisposable
{
    private readonly ILogger _logger;
    private static DiagnosticLogger? _instance;

    public static DiagnosticLogger Instance => _instance
        ?? throw new InvalidOperationException("DiagnosticLogger not initialized. Call Initialize() first.");

    private DiagnosticLogger(ILogger logger) => _logger = logger;

    public static void Initialize(string? logPath = null)
    {
        logPath ??= Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "EmailManage",
            "diagnostic.log");

        Directory.CreateDirectory(Path.GetDirectoryName(logPath)!);

        var logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.File(
                logPath,
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 14,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
            .CreateLogger();

        _instance = new DiagnosticLogger(logger);
        _instance.Info("DiagnosticLogger initialized. Path: {LogPath}", logPath);
    }

    public void Debug(string template, params object[] args) =>
        _logger.Debug(template, args);

    public void Info(string template, params object[] args) =>
        _logger.Information(template, args);

    public void Warn(string template, params object[] args) =>
        _logger.Warning(template, args);

    public void Error(string template, params object[] args) =>
        _logger.Error(template, args);

    public void Error(Exception ex, string template, params object[] args) =>
        _logger.Error(ex, template, args);

    public void Fatal(string template, params object[] args) =>
        _logger.Fatal(template, args);

    public void Dispose()
    {
        if (_logger is IDisposable disposable)
            disposable.Dispose();
    }
}
