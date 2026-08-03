namespace EmailManage.Services;

public enum SyncExitCode { Success = 0, PartialSuccess = 10, Cancelled = 20, OutlookUnavailable = 30, Fatal = 40, AlreadyRunning = 50 }

public sealed class HeadlessSyncRunner
{
    private static readonly Mutex SyncMutex = new(false, "Local\\MailZen-HeadlessSync");

    public async Task<SyncExitCode> RunAsync(DateTime? requestedFromUtc = null, DateTime? requestedToUtc = null,
        string? requestedAccount = null, CancellationToken cancellationToken = default)
    {
        if (!SyncMutex.WaitOne(TimeSpan.Zero)) return SyncExitCode.AlreadyRunning;
        try
        {
            var database = new MailZenDatabase();
            var outlook = new OutlookReadAdapter();
            DiagnosticLogger.Instance.Info("Headless sync: reading Outlook accounts.");
            var accounts = await outlook.GetAccountsAsync(cancellationToken);
            DiagnosticLogger.Instance.Info("Headless sync: found {AccountCount} Outlook accounts.", accounts.Count);
            if (accounts.Count == 0) return SyncExitCode.OutlookUnavailable;
            if (!string.IsNullOrWhiteSpace(requestedAccount))
            {
                accounts = accounts.Where(account =>
                    string.Equals(account.AccountId, requestedAccount, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(account.DisplayName, requestedAccount, StringComparison.OrdinalIgnoreCase)).ToArray();
                DiagnosticLogger.Instance.Info("Headless sync: account filter {RequestedAccount} matched {AccountCount} account(s).", requestedAccount, accounts.Count);
                if (accounts.Count == 0) return SyncExitCode.OutlookUnavailable;
            }
            var coordinator = new EmailIndexCoordinator(outlook, database);
            var accountIds = accounts.Select(a => a.AccountId).ToHashSet();
            // Use the oldest successful checkpoint so a folder that failed during a prior
            // run is included in the catch-up window instead of being skipped.
            var lastSuccessful = requestedFromUtc.HasValue ? null : await database.GetEarliestSuccessfulSyncUtcAsync(accountIds, cancellationToken);
            var sinceUtc = requestedFromUtc ?? lastSuccessful ?? DateTime.UtcNow.AddDays(-2);
            DiagnosticLogger.Instance.Info("Headless sync: indexing from {SinceUtc} through {ToUtc}.", sinceUtc, requestedToUtc?.ToString("O") ?? "now");
            await coordinator.SyncAsync(accountIds, sinceUtc, requestedToUtc, null, cancellationToken);
            DiagnosticLogger.Instance.Info("Headless sync: indexing coordinator completed.");
            var coverage = await database.GetIndexCoverageAsync(accountIds, cancellationToken);
            DiagnosticLogger.Instance.Info("Headless sync: database coverage is {MessageCount} message(s), earliest {Earliest}, latest {Latest}.", coverage.MessageCount, coverage.EarliestReceivedUtc?.ToString("O") ?? "none", coverage.LatestReceivedUtc?.ToString("O") ?? "none");
            return coverage.MessageCount > 0 ? SyncExitCode.Success : SyncExitCode.PartialSuccess;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { return SyncExitCode.Cancelled; }
        catch (System.Runtime.InteropServices.COMException) { return SyncExitCode.OutlookUnavailable; }
        catch { return SyncExitCode.Fatal; }
        finally { SyncMutex.ReleaseMutex(); }
    }
}
