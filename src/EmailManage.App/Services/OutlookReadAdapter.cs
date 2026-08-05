using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Threading.Channels;
using System.Globalization;
using EmailManage.Models;

namespace EmailManage.Services;

/// <summary>
/// Read-only Outlook Classic adapter for the MailZen 2.0 indexing path.
/// It deliberately has no Save, Move, Delete, Categorize, or MarkAsRead calls.
/// </summary>
public sealed class OutlookReadAdapter : IOutlookReadService
{
    private const int OlFolderInbox = 6;
    private const int OlFolderSentMail = 5;

    public async Task<IReadOnlyList<OutlookAccount>> GetAccountsAsync(CancellationToken cancellationToken = default)
    {
        var result = new List<OutlookAccount>();
        await RunStaAsync(() =>
        {
            dynamic? app = null;
            dynamic? session = null;
            dynamic? stores = null;
            try
            {
                DiagnosticLogger.Instance.Info("Outlook adapter: acquiring Outlook application.");
                app = GetOrCreateOutlookApplication();
                DiagnosticLogger.Instance.Info("Outlook adapter: acquiring MAPI namespace.");
                session = app.GetNamespace("MAPI");
                DiagnosticLogger.Instance.Info("Outlook adapter: reading Stores collection.");
                stores = session.Stores;
                DiagnosticLogger.Instance.Info("Outlook adapter: Stores count is {StoreCount}.", (int)stores.Count);
                for (var i = 1; i <= (int)stores.Count; i++)
                {
                    dynamic? store = null;
                    try
                    {
                        DiagnosticLogger.Instance.Info("Outlook adapter: reading store {StoreIndex}.", i);
                        store = stores[i];
                        var storeId = (string?)store.StoreID ?? string.Empty;
                        var displayName = (string?)store.DisplayName ?? storeId;
                        result.Add(new OutlookAccount(storeId, storeId, displayName, displayName));
                        DiagnosticLogger.Instance.Info("Outlook adapter: store {StoreIndex} is {DisplayName}.", i, displayName);
                    }
                    catch (Exception ex) { DiagnosticLogger.Instance.Warn("Outlook adapter: skipped store {StoreIndex}: {Error}", i, ex.Message); }
                    finally { Release((object?)store); }
                }
            }
            finally
            {
                DiagnosticLogger.Instance.Info("Outlook adapter: releasing Stores collection.");
                Release((object?)stores);
                DiagnosticLogger.Instance.Info("Outlook adapter: releasing MAPI namespace.");
                Release((object?)session);
                DiagnosticLogger.Instance.Info("Outlook adapter: releasing Outlook application.");
                Release((object?)app);
                DiagnosticLogger.Instance.Info("Outlook adapter: account read cleanup complete.");
            }
        }, cancellationToken);
        return result;
    }

    public async IAsyncEnumerable<OutlookFolder> EnumerateFoldersAsync(
        string accountId, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var folders = new List<OutlookFolder>();
        await RunStaAsync(() =>
        {
            dynamic? app = null; dynamic? session = null; dynamic? store = null; dynamic? inbox = null; dynamic? sent = null;
            try
            {
                app = GetOrCreateOutlookApplication();
                session = app.GetNamespace("MAPI");
                store = session.GetStoreFromID(accountId);
                inbox = store.GetDefaultFolder(OlFolderInbox);
                folders.Add(new OutlookFolder((string)inbox.EntryID, accountId, (string)inbox.EntryID, (string)inbox.FolderPath, "Inbox"));
                try
                {
                    sent = store.GetDefaultFolder(OlFolderSentMail);
                    folders.Add(new OutlookFolder((string)sent.EntryID, accountId, (string)sent.EntryID, (string)sent.FolderPath, "Sent"));
                }
                catch (Exception ex)
                {
                    DiagnosticLogger.Instance.Warn("Outlook adapter: Sent Items folder unavailable for {Account}: {Error}", accountId, ex.Message);
                }
            }
            finally { Release((object?)sent); Release((object?)inbox); Release((object?)store); Release((object?)session); Release((object?)app); }
        }, cancellationToken);
        foreach (var folder in folders)
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return folder;
        }
    }

    public async IAsyncEnumerable<IndexedMessage> ReadMessagesAsync(
        OutlookFolder folder, OutlookReadOptions options,
        [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var channel = Channel.CreateBounded<IndexedMessage>(new BoundedChannelOptions(Math.Max(10, options.BatchSize))
        { FullMode = BoundedChannelFullMode.Wait, SingleWriter = true, SingleReader = true });

        _ = Task.Run(async () =>
        {
            Exception? failure = null;
            try
            {
                await RunStaAsync(() => ReadFolder(folder, options, channel.Writer, cancellationToken), cancellationToken);
            }
            catch (Exception ex) { failure = ex; }
            finally { channel.Writer.TryComplete(failure); }
        }, cancellationToken);

        await foreach (var message in channel.Reader.ReadAllAsync(cancellationToken))
            yield return message;
    }

    private static void ReadFolder(OutlookFolder folder, OutlookReadOptions options, ChannelWriter<IndexedMessage> writer, CancellationToken token)
    {
        dynamic? app = null; dynamic? session = null; dynamic? store = null; dynamic? outlookFolder = null; dynamic? allItems = null; dynamic? items = null;
        try
        {
            app = GetOrCreateOutlookApplication();
            session = app.GetNamespace("MAPI");
            store = session.GetStoreFromID(folder.AccountId);
            // GetFolderFromID is a Namespace/MAPI method, not a Store method.
            // The previous call silently prevented every folder from being indexed.
            outlookFolder = session.GetFolderFromID(folder.EntryId, folder.AccountId);
            allItems = outlookFolder.Items;
            items = allItems;
            if (options.ReceivedAfterUtc.HasValue || options.ReceivedBeforeUtc.HasValue)
            {
                var after = options.ReceivedAfterUtc?.ToLocalTime().ToString("g", CultureInfo.CurrentCulture);
                var before = options.ReceivedBeforeUtc?.ToLocalTime().ToString("g", CultureInfo.CurrentCulture);
                var restrictions = new List<string>();
                if (after is not null) restrictions.Add($"[ReceivedTime] >= '{after}'");
                if (before is not null) restrictions.Add($"[ReceivedTime] < '{before}'");
                items = allItems.Restrict(string.Join(" AND ", restrictions));
            }
            items.Sort("[ReceivedTime]", true);
            var count = (int)items.Count;
            DiagnosticLogger.Instance.Info("Outlook adapter: reading {ItemCount} date-filtered item(s) from {Folder}.", count, folder.Path);
            for (var i = 1; i <= count; i++)
            {
                token.ThrowIfCancellationRequested();
                dynamic? item = null;
                try
                {
                    item = items[i];
                    string? messageClass;
                    try { messageClass = (string?)item.MessageClass; }
                    catch { continue; }
                    if (!IsMailMessageClass(messageClass)) continue;

                    DateTime received;
                    try { received = (DateTime?)item.ReceivedTime ?? DateTime.MinValue; }
                    catch { continue; }
                    if (options.ReceivedAfterUtc.HasValue && received.ToUniversalTime() < options.ReceivedAfterUtc.Value) continue;
                    if (options.ReceivedBeforeUtc.HasValue && received.ToUniversalTime() > options.ReceivedBeforeUtc.Value) continue;

                    var entryId = (string?)item.EntryID ?? string.Empty;
                    var senderAddress = (string?)item.SenderEmailAddress ?? string.Empty;
                    var internetMessageId = TryGetString(item, "InternetMessageID");
                    var message = new IndexedMessage(
                        MessageId: BuildMessageId(folder.AccountId, entryId, internetMessageId),
                        AccountId: folder.AccountId, FolderId: folder.FolderId, StoreId: folder.AccountId,
                        EntryId: entryId, InternetMessageId: internetMessageId,
                        Subject: (string?)item.Subject ?? string.Empty, SenderName: (string?)item.SenderName ?? string.Empty,
                        SenderAddress: senderAddress, BodyText: (string?)item.Body ?? string.Empty,
                        ReceivedUtc: received.ToUniversalTime(), IsUnread: (bool?)item.UnRead ?? false,
                        HasAttachments: ((int?)item.Attachments?.Count ?? 0) > 0,
                        AttachmentNames: ReadAttachmentNames(item), ConversationId: TryGetString(item, "ConversationID"),
                        FolderType: folder.FolderType);
                    writer.WriteAsync(message, token).AsTask().GetAwaiter().GetResult();
                }
                finally { Release((object?)item); }
            }
        }
        finally
        {
            var sameItems = ReferenceEquals((object?)items, (object?)allItems);
            Release((object?)items);
            if (!sameItems) Release((object?)allItems);
            Release((object?)outlookFolder);
            Release((object?)store);
            Release((object?)session);
            Release((object?)app);
        }
    }

    private static string ReadAttachmentNames(dynamic item)
    {
        dynamic? attachments = null;
        try
        {
            attachments = item.Attachments;
            var names = new List<string>();
            for (var i = 1; i <= (int)attachments.Count; i++)
            {
                dynamic? attachment = null;
                try { attachment = attachments[i]; names.Add((string?)attachment.FileName ?? string.Empty); }
                finally { Release((object?)attachment); }
            }
            return string.Join("; ", names);
        }
        catch { return string.Empty; }
        finally { Release((object?)attachments); }
    }

    private static string? TryGetString(dynamic item, string propertyName)
    {
        try
        {
            return propertyName switch
            {
                "InternetMessageID" => (string?)item.InternetMessageID,
                "ConversationID" => (string?)item.ConversationID,
                _ => null
            };
        }
        catch { return null; }
    }

    private static bool IsMailMessageClass(string? messageClass) =>
        messageClass is not null &&
        (string.Equals(messageClass, "IPM.Note", StringComparison.OrdinalIgnoreCase) ||
         messageClass.StartsWith("IPM.Note.", StringComparison.OrdinalIgnoreCase) ||
         messageClass.StartsWith("REPORT.IPM.Note", StringComparison.OrdinalIgnoreCase));

    private static async Task RunStaAsync(Action action, CancellationToken token)
    {
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            try
            {
                token.ThrowIfCancellationRequested();
                action();
                DiagnosticLogger.Instance.Info("Outlook adapter: STA action completed; signaling caller.");
                completion.TrySetResult();
                DiagnosticLogger.Instance.Info("Outlook adapter: STA caller signaled.");
            }
            catch (Exception ex) { completion.TrySetException(ex); }
        });
        thread.IsBackground = true;
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        await completion.Task.WaitAsync(token);
    }

    private static string BuildMessageId(string accountId, string entryId, string? internetId) =>
        !string.IsNullOrWhiteSpace(internetId) ? $"{accountId}:internet:{internetId}" : $"{accountId}:entry:{entryId}";

    private static object GetActiveComObject(string progId)
    {
        var clsid = Type.GetTypeFromProgID(progId, true)!.GUID;
        var hr = GetActiveObject(ref clsid, IntPtr.Zero, out var obj);
        if (hr < 0) Marshal.ThrowExceptionForHR(hr);
        return obj;
    }

    private static object GetOrCreateOutlookApplication()
    {
        try { return GetActiveComObject("Outlook.Application"); }
        catch (COMException)
        {
            var type = Type.GetTypeFromProgID("Outlook.Application", true)!;
            return Activator.CreateInstance(type) ?? throw new COMException("Could not start Outlook Classic.");
        }
    }

    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            try { Marshal.ReleaseComObject(value); }
            catch (InvalidComObjectException) { }
        }
    }
}
