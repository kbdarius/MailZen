using EmailManage.Models;

namespace EmailManage.Services;

/// <summary>
/// Coordinates resumable, account-scoped indexing. Database writes are performed one
/// message at a time through the database service so cancellation leaves committed rows valid.
/// </summary>
public sealed class EmailIndexCoordinator : IEmailIndexService
{
    private readonly IOutlookReadService _outlook;
    private readonly MailZenDatabase _database;

    public EmailIndexCoordinator(IOutlookReadService outlook, MailZenDatabase database)
    {
        _outlook = outlook;
        _database = database;
    }

    public async Task SyncAsync(IReadOnlySet<string> accountIds, DateTime? sinceUtc, DateTime? beforeUtc = null,
        IProgress<string>? progress = null, CancellationToken cancellationToken = default)
    {
        await _database.InitializeAsync(cancellationToken);
        var accounts = await _outlook.GetAccountsAsync(cancellationToken);
        var selected = accounts.Where(a => accountIds.Contains(a.AccountId)).ToList();
        DiagnosticLogger.Instance.Info("Index coordinator: selected {AccountCount} account(s).", selected.Count);

        foreach (var account in selected)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                await foreach (var folder in _outlook.EnumerateFoldersAsync(account.AccountId, cancellationToken))
                {
                    try
                    {
                        DiagnosticLogger.Instance.Info("Index coordinator: starting {Account}/{Folder}.", account.DisplayName, folder.Path);
                        progress?.Report($"Indexing {account.DisplayName}: {folder.Path}");
                        var effectiveSinceUtc = sinceUtc;
                        if (sinceUtc.HasValue && beforeUtc.HasValue)
                        {
                            effectiveSinceUtc = await _database.GetEffectiveSyncStartUtcAsync(account.AccountId, folder.FolderId,
                                sinceUtc.Value, beforeUtc.Value, TimeSpan.FromDays(2), cancellationToken);
                            DiagnosticLogger.Instance.Info("Index coordinator: requested {RequestedStart}–{RequestedEnd}; effective read starts at {EffectiveStart}.",
                                sinceUtc.Value, beforeUtc.Value, effectiveSinceUtc.Value);
                        }
                        var options = new OutlookReadOptions(effectiveSinceUtc, beforeUtc);
                        await foreach (var message in _outlook.ReadMessagesAsync(folder, options, cancellationToken))
                            await _database.UpsertMessageAsync(message, cancellationToken);
                        var completedUtc = DateTime.UtcNow;
                        await _database.MarkFolderSyncSuccessAsync(account.AccountId, folder.FolderId, completedUtc, cancellationToken);
                        if (sinceUtc.HasValue)
                            await _database.RecordSyncCoverageAsync(account.AccountId, folder.FolderId, sinceUtc.Value,
                                beforeUtc ?? completedUtc, completedUtc, cancellationToken: cancellationToken);
                        DiagnosticLogger.Instance.Info("Index coordinator: completed {Account}/{Folder}.", account.DisplayName, folder.Path);
                    }
                    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
                    catch (Exception ex)
                    {
                        DiagnosticLogger.Instance.Error(ex, "Index coordinator: failed {Account}/{Folder}.", account.DisplayName, folder.Path);
                        await _database.MarkFolderSyncFailureAsync(account.AccountId, folder.FolderId, DateTime.UtcNow, ex.Message, cancellationToken);
                        if (sinceUtc.HasValue)
                            await _database.RecordSyncCoverageAsync(account.AccountId, folder.FolderId, sinceUtc.Value,
                                beforeUtc ?? DateTime.UtcNow, DateTime.UtcNow, "failed", ex.Message, cancellationToken);
                        progress?.Report($"Unable to index {account.DisplayName}/{folder.Path}: {ex.Message}");
                    }
                }
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { throw; }
            catch (Exception ex)
            {
                // One account failure must not roll back already committed accounts.
                DiagnosticLogger.Instance.Error(ex, "Index coordinator: failed account {Account}.", account.DisplayName);
                progress?.Report($"Unable to index {account.DisplayName}: {ex.Message}");
            }
        }
    }
}
