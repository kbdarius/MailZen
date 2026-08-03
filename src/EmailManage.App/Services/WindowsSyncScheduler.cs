using System.Diagnostics;

namespace EmailManage.Services;

public sealed class WindowsSyncScheduler : ISyncScheduler
{
    public const string TaskName = "MailZen Daily Sync";
    public async Task ConfigureDailyAsync(TimeOnly localTime, CancellationToken cancellationToken = default)
    {
        var executable = Process.GetCurrentProcess().MainModule?.FileName ?? throw new InvalidOperationException("Cannot determine MailZen executable path.");
        var arguments = $"/Create /TN \"{TaskName}\" /SC DAILY /ST {localTime:HH\\:mm} /TR \"\\\"{executable}\\\" --sync --quiet\" /RL LIMITED /F";
        await RunAsync(arguments, cancellationToken);
    }
    public async Task RemoveAsync(CancellationToken cancellationToken = default) => await RunAsync($"/Delete /TN \"{TaskName}\" /F", cancellationToken, true);
    public async Task<bool> ExistsAsync(CancellationToken cancellationToken = default)
    { var result = await RunAsync($"/Query /TN \"{TaskName}\"", cancellationToken, true); return result == 0; }

    private static async Task<int> RunAsync(string arguments, CancellationToken token, bool allowFailure = false)
    {
        using var process = Process.Start(new ProcessStartInfo("schtasks.exe", arguments) { CreateNoWindow = true, UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true }) ?? throw new InvalidOperationException("Could not start Task Scheduler.");
        await process.WaitForExitAsync(token); if (process.ExitCode != 0 && !allowFailure) throw new InvalidOperationException((await process.StandardError.ReadToEndAsync(token)).Trim()); return process.ExitCode;
    }
}
