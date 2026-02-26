using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using EmailManage.Models;

namespace EmailManage.Services;

/// <summary>
/// Manages the COM connection to Outlook Desktop and provides account enumeration.
/// Uses late-bound COM (dynamic) to avoid Office interop assembly dependencies.
/// All COM calls are marshalled to an STA thread.
/// </summary>
public sealed class OutlookConnectorService : IDisposable
{
    private dynamic? _outlookApp;
    private dynamic? _session;
    private readonly DiagnosticLogger _log;
    private bool _disposed;

    // OlDefaultFolders constants (late-bound, no interop enum needed)
    private const int olFolderInbox = 6;
    private const int olFolderDeletedItems = 3;

    // Triage Folder Names
    public const string SmartCleanupFolderName = "Smart Cleanup";
    public const string DeleteCandidatesFolderName = "Delete Candidates";
    public const string NeedsReviewFolderName = "Needs Review";
    public const string AiReviewFolderName = "AI Review";

    // OlAccountType constants
    private const int olExchange = 0;
    private const int olImap = 3;
    private const int olPop3 = 1;
    private const int olHttp = 4;

    // OlExchangeStoreType constants
    private const int olExchangePublicFolder = 2;
    private const int olPrimaryExchangeMailbox = 0;
    private const int olAdditionalExchangeMailbox = 1;

    public bool IsConnected => _outlookApp is not null && _session is not null;

    public OutlookConnectorService()
    {
        _log = DiagnosticLogger.Instance;
    }

    /// <summary>
    /// Connect to a running Outlook instance or start a new one.
    /// Returns a <see cref="ConnectionResult"/> with account list or error.
    /// </summary>
    public async Task<ConnectionResult> ConnectAsync()
    {
        var sw = Stopwatch.StartNew();

        try
        {
            _log.Info("Attempting Outlook connection...");

            // COM calls must happen on an STA thread
            var result = await Task.Run(() =>
            {
                ConnectionResult? comResult = null;
                var thread = new Thread(() =>
                {
                    try { comResult = ConnectOnStaThread(); }
                    catch (Exception ex) { comResult = ConnectionResult.Fail("E005", $"Unexpected error: {ex.Message}", sw.Elapsed); }
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join(TimeSpan.FromSeconds(30));

                if (comResult is null)
                    return ConnectionResult.Fail("E005", "Outlook connection timed out after 30 seconds.", sw.Elapsed);

                return comResult;
            });

            sw.Stop();

            if (result.Success)
                _log.Info("OutlookConnected: {AccountCount} accounts found in {ElapsedMs}ms",
                    result.Accounts.Count, sw.ElapsedMilliseconds);
            else
                _log.Error("OutlookConnectionFailed: {ErrorCode} - {ErrorMessage}", result.ErrorCode ?? "", result.ErrorMessage ?? "");

            return result;
        }
        catch (Exception ex)
        {
            sw.Stop();
            _log.Error(ex, "OutlookConnectionFailed: Unexpected exception");
            return ConnectionResult.Fail("E005",
                $"Could not read Outlook accounts. Error: {ex.Message}. Check that Outlook is responding and click Retry.",
                sw.Elapsed);
        }
    }

    private ConnectionResult ConnectOnStaThread()
    {
        var sw = Stopwatch.StartNew();
        try
        {
            // Try to get running Outlook instance first
            try
            {
                _outlookApp = GetActiveComObject("Outlook.Application");
                _log.Debug("Attached to running Outlook instance.");
            }
            catch (COMException)
            {
                // Outlook not running — check if installed
                if (!IsOutlookInstalled())
                {
                    return ConnectionResult.Fail("E001",
                        "Outlook Desktop is not installed or not detected. Please install Microsoft Outlook and restart.",
                        sw.Elapsed);
                }

                try
                {
                    var outlookType = Type.GetTypeFromProgID("Outlook.Application", true)!;
                    _outlookApp = Activator.CreateInstance(outlookType);
                    _log.Debug("Started new Outlook instance.");
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80080005))
                {
                    return ConnectionResult.Fail("E003",
                        "Outlook profile is in use by another process. Close other Outlook automation tools and click Retry.",
                        sw.Elapsed);
                }
            }

            _session = _outlookApp!.GetNamespace("MAPI");
            _session.Logon("", "", false, false);

            var accounts = EnumerateAccounts();

            if (accounts.Count == 0)
            {
                return ConnectionResult.Fail("E004",
                    "No email accounts found in Outlook. Please configure at least one account in Outlook.",
                    sw.Elapsed);
            }

            return ConnectionResult.Ok(accounts, sw.Elapsed);
        }
        catch (COMException ex)
        {
            _log.Error(ex, "COM exception during Outlook connection");

            if (!IsOutlookRunning())
            {
                return ConnectionResult.Fail("E002",
                    "Outlook is not running. Please start Outlook and click Retry.",
                    sw.Elapsed);
            }

            return ConnectionResult.Fail("E005",
                $"Could not read Outlook accounts. Error: {ex.Message}. Check that Outlook is responding and click Retry.",
                sw.Elapsed);
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Unexpected exception during Outlook connection");
            return ConnectionResult.Fail("E005",
                $"Could not read Outlook accounts. Error: {ex.Message}. Check that Outlook is responding and click Retry.",
                sw.Elapsed);
        }
    }

    private List<OutlookAccountInfo> EnumerateAccounts()
    {
        var accounts = new List<OutlookAccountInfo>();

        if (_session is null) return accounts;

        try
        {
            // Enumerate via Accounts collection (preferred)
            dynamic outlookAccounts = _session.Accounts;
            int count = (int)outlookAccounts.Count;

            for (int i = 1; i <= count; i++)
            {
                try
                {
                    dynamic acct = outlookAccounts[i];
                    string displayName = "(Unknown)";
                    string email = "(No address)";

                    try { displayName = (string)(acct.DisplayName ?? "(Unknown)"); } catch { }
                    try { email = (string)(acct.SmtpAddress ?? "(No address)"); } catch { }

                    var info = new OutlookAccountInfo
                    {
                        DisplayName = displayName,
                        EmailAddress = email,
                        StoreName = GetStoreNameForAccount(acct),
                        StoreFilePath = GetStoreFilePathForAccount(acct),
                        ProviderHint = GetProviderHint(acct),
                        IsConnected = true
                    };

                    // Get folder counts
                    try
                    {
                        dynamic? store = GetStoreForAccount(acct);
                        if (store is not null)
                        {
                            try
                            {
                                dynamic inbox = store.GetDefaultFolder(olFolderInbox);
                                info.InboxCount = (int)inbox.Items.Count;
                            }
                            catch { }

                            try
                            {
                                dynamic deleted = store.GetDefaultFolder(olFolderDeletedItems);
                                info.DeletedItemsCount = (int)deleted.Items.Count;
                            }
                            catch { }
                        }
                    }
                    catch (Exception ex)
                    {
                        _log.Warn("Could not get folder counts for {Account}: {Error}",
                            info.EmailAddress, ex.Message);
                    }

                    // Check if pattern file exists
                    info.HasPatternFile = CheckPatternFileExists(info.AccountKey);

                    accounts.Add(info);
                    _log.Debug("Found account: {Email} ({DisplayName}), Provider: {Provider}",
                        info.EmailAddress, info.DisplayName, info.ProviderHint);
                }
                catch (Exception ex)
                {
                    _log.Warn("Failed to read account at index {Index}: {Error}", i, ex.Message);
                }
            }
        }
        catch (Exception ex)
        {
            _log.Error(ex, "Failed to enumerate accounts, trying stores fallback");

            // Fallback: enumerate via Stores
            try
            {
                dynamic stores = _session.Stores;
                int storeCount = (int)stores.Count;

                for (int i = 1; i <= storeCount; i++)
                {
                    try
                    {
                        dynamic store = stores[i];

                        // Skip public folders
                        try
                        {
                            int storeType = (int)store.ExchangeStoreType;
                            if (storeType == olExchangePublicFolder) continue;
                        }
                        catch { }

                        string storeName = "(Unknown Store)";
                        string storeEmail = "(Unknown)";
                        string filePath = "";
                        bool isOpen = true;

                        try { storeName = (string)(store.DisplayName ?? "(Unknown Store)"); } catch { }
                        try { filePath = (string)(store.FilePath ?? ""); } catch { }
                        try { isOpen = (bool)store.IsOpen; } catch { }

                        storeEmail = storeName.Contains('@') ? storeName : storeName;

                        var info = new OutlookAccountInfo
                        {
                            DisplayName = storeName,
                            EmailAddress = storeEmail,
                            StoreName = storeName,
                            StoreFilePath = filePath,
                            ProviderHint = GetProviderHintFromStore(store),
                            IsConnected = isOpen
                        };

                        info.HasPatternFile = CheckPatternFileExists(info.AccountKey);
                        accounts.Add(info);
                    }
                    catch (Exception ex2)
                    {
                        _log.Warn("Failed to read store at index {Index}: {Error}", i, ex2.Message);
                    }
                }
            }
            catch (Exception ex3)
            {
                _log.Error(ex3, "Stores fallback also failed");
            }
        }

        return accounts;
    }

    public async Task<List<EmailMessageInfo>> GetEmailsAsync(string emailAddress, string storeName, OutlookFolderType folderType, int maxMonths, int maxItems = 1000, bool skipUnread = true)
    {
        return await Task.Run(() =>
        {
            var emails = new List<EmailMessageInfo>();
            var thread = new Thread(() =>
            {
                try
                {
                    dynamic outlookApp = GetActiveComObject("Outlook.Application");
                    dynamic session = outlookApp.GetNamespace("MAPI");
                    session.Logon("", "", false, false);

                    _log.Info("GetEmailsAsync: Searching stores for storeName=[{StoreName}] email=[{Email}]", storeName, emailAddress);
                    dynamic? targetStore = null;
                    dynamic stores = session.Stores;
                    int count = (int)stores.Count;
                    _log.Info("GetEmailsAsync: Found {Count} stores total", count);
                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            dynamic store = stores[i];
                            string sName = (string)store.DisplayName;
                            _log.Info("GetEmailsAsync: Store[{Index}] DisplayName=[{Name}]", i, sName);
                            if (sName == storeName || sName == emailAddress || sName.Contains(emailAddress))
                            {
                                targetStore = store;
                                break;
                            }
                        }
                        catch { }
                    }

                    if (targetStore is null)
                    {
                        _log.Warn("Could not find store with name {StoreName} or email {EmailAddress}", storeName, emailAddress);
                        return;
                    }

                    dynamic folder = targetStore.GetDefaultFolder((int)folderType);
                    dynamic items = folder.Items;
                    
                    // Sort by received time descending
                    try { items.Sort("[ReceivedTime]", true); } catch { }

                    DateTime cutoffDate = DateTime.Now.AddMonths(-maxMonths);
                    int fetched = 0;

                    // Iterate through items
                    int itemCount = (int)items.Count;
                    _log.Info("Found {Count} items in folder {FolderType} for store {StoreName}", itemCount, folderType, storeName);
                    
                    for (int i = 1; i <= itemCount && fetched < maxItems; i++)
                    {
                        try
                        {
                            dynamic item = items[i];
                            
                            // Only process MailItem (Class == 43)
                            if ((int)item.Class != 43) continue;

                            DateTime receivedTime = (DateTime)item.ReceivedTime;
                            if (receivedTime < cutoffDate) break; // Since it's sorted, we can stop

                            bool unread = (bool)item.UnRead;

                            // For inbox, we only want read emails (kept) if skipUnread is true
                            if (skipUnread && folderType == OutlookFolderType.Inbox && unread) continue;

                            string senderEmail = "";
                            try { senderEmail = (string)item.SenderEmailAddress; } catch { }
                            
                            // Try to get SMTP address if it's an Exchange user
                            try
                            {
                                if ((int)item.SenderEmailType == 0) // olExchange
                                {
                                    dynamic sender = item.Sender;
                                    if (sender != null)
                                    {
                                        dynamic pa = sender.PropertyAccessor;
                                        senderEmail = (string)pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E");
                                    }
                                }
                            }
                            catch { }

                            emails.Add(new EmailMessageInfo
                            {
                                EntryId = (string)item.EntryID,
                                Subject = (string)(item.Subject ?? ""),
                                SenderName = (string)(item.SenderName ?? ""),
                                SenderEmailAddress = senderEmail,
                                Body = (string)(item.Body ?? ""),
                                ReceivedTime = receivedTime,
                                UnRead = unread,
                                IsDeleted = folderType == OutlookFolderType.DeletedItems
                            });

                            fetched++;
                        }
                        catch { }
                    }
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error fetching emails from folder {FolderType}", folderType);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromMinutes(2)); // Allow more time for fetching

            return emails;
        });
    }

    public async Task<bool> EnsureTriageFoldersAsync(string emailAddress, string storeName)
    {
        return await Task.Run(() =>
        {
            bool success = false;
            var thread = new Thread(() =>
            {
                try
                {
                    dynamic outlookApp = GetActiveComObject("Outlook.Application");
                    dynamic session = outlookApp.GetNamespace("MAPI");
                    session.Logon("", "", false, false);

                    dynamic? targetStore = null;
                    dynamic stores = session.Stores;
                    int count = (int)stores.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            dynamic store = stores[i];
                            string sName = (string)store.DisplayName;
                            if (sName == storeName || sName == emailAddress || sName.Contains(emailAddress))
                            {
                                targetStore = store;
                                break;
                            }
                        }
                        catch { }
                    }

                    if (targetStore is null) return;

                    dynamic rootFolder = targetStore.GetRootFolder();
                    dynamic folders = rootFolder.Folders;

                    // 1. Ensure "Smart Cleanup" exists
                    dynamic? smartCleanupFolder = null;
                    try { smartCleanupFolder = folders[SmartCleanupFolderName]; } catch { }
                    
                    if (smartCleanupFolder is null)
                    {
                        smartCleanupFolder = folders.Add(SmartCleanupFolderName);
                        _log.Info("Created folder: {FolderName}", SmartCleanupFolderName);
                    }

                    // 2. Ensure "Delete Candidates" exists
                    dynamic scFolders = smartCleanupFolder.Folders;
                    dynamic? deleteCandidatesFolder = null;
                    try { deleteCandidatesFolder = scFolders[DeleteCandidatesFolderName]; } catch { }
                    
                    if (deleteCandidatesFolder is null)
                    {
                        deleteCandidatesFolder = scFolders.Add(DeleteCandidatesFolderName);
                        _log.Info("Created folder: {FolderName}", DeleteCandidatesFolderName);
                    }

                    // 3. Ensure "Needs Review" exists
                    dynamic? needsReviewFolder = null;
                    try { needsReviewFolder = scFolders[NeedsReviewFolderName]; } catch { }
                    
                    if (needsReviewFolder is null)
                    {
                        needsReviewFolder = scFolders.Add(NeedsReviewFolderName);
                        _log.Info("Created folder: {FolderName}", NeedsReviewFolderName);
                    }

                    success = true;
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error ensuring triage folders exist");
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromSeconds(30));

            return success;
        });
    }

    public async Task<bool> MoveEmailAsync(string entryId, string emailAddress, string storeName, string targetSubFolderName)
    {
        return await Task.Run(() =>
        {
            bool success = false;
            var thread = new Thread(() =>
            {
                try
                {
                    dynamic outlookApp = GetActiveComObject("Outlook.Application");
                    dynamic session = outlookApp.GetNamespace("MAPI");
                    session.Logon("", "", false, false);

                    dynamic? targetStore = null;
                    dynamic stores = session.Stores;
                    int count = (int)stores.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            dynamic store = stores[i];
                            string sName = (string)store.DisplayName;
                            if (sName == storeName || sName == emailAddress || sName.Contains(emailAddress))
                            {
                                targetStore = store;
                                break;
                            }
                        }
                        catch { }
                    }

                    if (targetStore is null) return;

                    // Get target folder
                    dynamic rootFolder = targetStore.GetRootFolder();
                    dynamic smartCleanupFolder = rootFolder.Folders[SmartCleanupFolderName];
                    dynamic targetFolder = smartCleanupFolder.Folders[targetSubFolderName];

                    // Get item and move
                    dynamic item = session.GetItemFromID(entryId, targetStore.StoreID);
                    item.Move(targetFolder);
                    
                    success = true;
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error moving email {EntryId} to {TargetFolder}", entryId, targetSubFolderName);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromSeconds(15));

            return success;
        });
    }

    /// <summary>
    /// Fetches emails from a custom subfolder under "Smart Cleanup" (e.g., "Needs Review" or "Delete Candidates").
    /// </summary>
    public async Task<List<EmailMessageInfo>> GetTriageFolderEmailsAsync(string emailAddress, string storeName, string subFolderName, int maxItems = 500)
    {
        return await Task.Run(() =>
        {
            var emails = new List<EmailMessageInfo>();
            var thread = new Thread(() =>
            {
                try
                {
                    dynamic outlookApp = GetActiveComObject("Outlook.Application");
                    dynamic session = outlookApp.GetNamespace("MAPI");
                    session.Logon("", "", false, false);

                    dynamic? targetStore = FindStore(session, emailAddress, storeName);
                    if (targetStore is null) { _log.Warn("GetTriageFolderEmails: Store not found"); return; }

                    dynamic rootFolder = targetStore.GetRootFolder();
                    dynamic smartCleanup = rootFolder.Folders[SmartCleanupFolderName];
                    dynamic targetFolder = smartCleanup.Folders[subFolderName];
                    dynamic items = targetFolder.Items;

                    try { items.Sort("[ReceivedTime]", true); } catch { }

                    int itemCount = (int)items.Count;
                    _log.Info("GetTriageFolderEmails: Found {Count} items in {Folder}", itemCount, subFolderName);
                    int fetched = 0;

                    for (int i = 1; i <= itemCount && fetched < maxItems; i++)
                    {
                        try
                        {
                            dynamic item = items[i];
                            if ((int)item.Class != 43) continue;

                            string senderEmail = "";
                            try { senderEmail = (string)item.SenderEmailAddress; } catch { }
                            try
                            {
                                if ((int)item.SenderEmailType == 0)
                                {
                                    dynamic sender = item.Sender;
                                    if (sender != null)
                                    {
                                        dynamic pa = sender.PropertyAccessor;
                                        senderEmail = (string)pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E");
                                    }
                                }
                            }
                            catch { }

                            emails.Add(new EmailMessageInfo
                            {
                                EntryId = (string)item.EntryID,
                                Subject = (string)(item.Subject ?? ""),
                                SenderName = (string)(item.SenderName ?? ""),
                                SenderEmailAddress = senderEmail,
                                Body = (string)(item.Body ?? ""),
                                ReceivedTime = (DateTime)item.ReceivedTime,
                                UnRead = (bool)item.UnRead,
                                IsDeleted = false
                            });
                            fetched++;
                        }
                        catch (Exception ex) { _log.Warn("Skipping item {Index}: {Error}", i, ex.Message); }
                    }
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error fetching emails from triage folder {Folder}", subFolderName);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromMinutes(2));
            return emails;
        });
    }

    /// <summary>
    /// Moves an email to the Deleted Items folder (permanent-ish delete via Outlook).
    /// Returns the new EntryId after move (for undo tracking).
    /// </summary>
    public async Task<string?> MoveToDeletedItemsAsync(string entryId, string emailAddress, string storeName)
    {
        return await Task.Run(() =>
        {
            string? newEntryId = null;
            var thread = new Thread(() =>
            {
                try
                {
                    dynamic outlookApp = GetActiveComObject("Outlook.Application");
                    dynamic session = outlookApp.GetNamespace("MAPI");
                    session.Logon("", "", false, false);

                    dynamic? targetStore = FindStore(session, emailAddress, storeName);
                    if (targetStore is null) return;

                    dynamic deletedFolder = targetStore.GetDefaultFolder(olFolderDeletedItems);
                    dynamic item = session.GetItemFromID(entryId, targetStore.StoreID);
                    dynamic movedItem = item.Move(deletedFolder);
                    newEntryId = (string)movedItem.EntryID;
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error moving email {EntryId} to Deleted Items", entryId);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromSeconds(15));
            return newEntryId;
        });
    }

    /// <summary>
    /// Moves an email back to the Inbox folder (for Keep actions or undo).
    /// Returns the new EntryId after move.
    /// </summary>
    public async Task<string?> MoveToInboxAsync(string entryId, string emailAddress, string storeName)
    {
        return await Task.Run(() =>
        {
            string? newEntryId = null;
            var thread = new Thread(() =>
            {
                try
                {
                    dynamic outlookApp = GetActiveComObject("Outlook.Application");
                    dynamic session = outlookApp.GetNamespace("MAPI");
                    session.Logon("", "", false, false);

                    dynamic? targetStore = FindStore(session, emailAddress, storeName);
                    if (targetStore is null) return;

                    dynamic inboxFolder = targetStore.GetDefaultFolder(olFolderInbox);
                    dynamic item = session.GetItemFromID(entryId, targetStore.StoreID);
                    dynamic movedItem = item.Move(inboxFolder);
                    newEntryId = (string)movedItem.EntryID;
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error moving email {EntryId} to Inbox", entryId);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromSeconds(15));
            return newEntryId;
        });
    }

    /// <summary>
    /// Moves an email back to a triage subfolder (for undo of delete actions).
    /// Returns the new EntryId after move.
    /// </summary>
    public async Task<string?> MoveToTriageFolderAsync(string entryId, string emailAddress, string storeName, string subFolderName)
    {
        return await Task.Run(() =>
        {
            string? newEntryId = null;
            var thread = new Thread(() =>
            {
                try
                {
                    dynamic outlookApp = GetActiveComObject("Outlook.Application");
                    dynamic session = outlookApp.GetNamespace("MAPI");
                    session.Logon("", "", false, false);

                    dynamic? targetStore = FindStore(session, emailAddress, storeName);
                    if (targetStore is null) return;

                    dynamic rootFolder = targetStore.GetRootFolder();
                    dynamic smartCleanupFolder = rootFolder.Folders[SmartCleanupFolderName];
                    dynamic targetFolder = smartCleanupFolder.Folders[subFolderName];

                    dynamic item = session.GetItemFromID(entryId, targetStore.StoreID);
                    dynamic movedItem = item.Move(targetFolder);
                    newEntryId = (string)movedItem.EntryID;
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error moving email {EntryId} to triage folder {Folder}", entryId, subFolderName);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromSeconds(15));
            return newEntryId;
        });
    }

    /// <summary>
    /// Takes a snapshot of all EntryIDs (+ sender info) in a given triage subfolder.
    /// Used for before/after comparison when user reviews in Outlook.
    /// </summary>
    public async Task<List<FolderEmailSnapshot>> GetFolderSnapshotAsync(string emailAddress, string storeName, string subFolderName, int maxItems = 1000)
    {
        return await Task.Run(() =>
        {
            var snapshot = new List<FolderEmailSnapshot>();
            var thread = new Thread(() =>
            {
                try
                {
                    dynamic outlookApp = GetActiveComObject("Outlook.Application");
                    dynamic session = outlookApp.GetNamespace("MAPI");
                    session.Logon("", "", false, false);

                    dynamic? targetStore = FindStore(session, emailAddress, storeName);
                    if (targetStore is null) { _log.Warn("GetFolderSnapshot: Store not found"); return; }

                    dynamic rootFolder = targetStore.GetRootFolder();
                    dynamic smartCleanup = rootFolder.Folders[SmartCleanupFolderName];
                    dynamic targetFolder = smartCleanup.Folders[subFolderName];
                    dynamic items = targetFolder.Items;
                    int itemCount = (int)items.Count;

                    for (int i = 1; i <= itemCount && i <= maxItems; i++)
                    {
                        try
                        {
                            dynamic item = items[i];
                            if ((int)item.Class != 43) continue;

                            string senderEmail = "";
                            try { senderEmail = (string)item.SenderEmailAddress; } catch { }
                            try
                            {
                                if ((int)item.SenderEmailType == 0)
                                {
                                    dynamic sender = item.Sender;
                                    if (sender != null)
                                    {
                                        dynamic pa = sender.PropertyAccessor;
                                        senderEmail = (string)pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E");
                                    }
                                }
                            }
                            catch { }

                            snapshot.Add(new FolderEmailSnapshot
                            {
                                EntryId = (string)item.EntryID,
                                Subject = (string)(item.Subject ?? ""),
                                SenderEmail = senderEmail?.ToLowerInvariant() ?? "",
                                Domain = ExtractDomainFromEmail(senderEmail),
                                ReceivedTime = (DateTime)item.ReceivedTime
                            });
                        }
                        catch (Exception ex) { _log.Warn("Snapshot: skip item {Index}: {Error}", i, ex.Message); }
                    }

                    _log.Info("Snapshot taken: {Count} items in {Folder}", snapshot.Count, subFolderName);
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error taking folder snapshot for {Folder}", subFolderName);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromMinutes(2));
            return snapshot;
        });
    }

    /// <summary>
    /// Gets the count of items in a triage subfolder.
    /// </summary>
    public async Task<int> GetFolderItemCountAsync(string emailAddress, string storeName, string subFolderName)
    {
        return await Task.Run(() =>
        {
            int count = 0;
            var thread = new Thread(() =>
            {
                try
                {
                    dynamic outlookApp = GetActiveComObject("Outlook.Application");
                    dynamic session = outlookApp.GetNamespace("MAPI");
                    session.Logon("", "", false, false);

                    dynamic? targetStore = FindStore(session, emailAddress, storeName);
                    if (targetStore is null) return;

                    dynamic rootFolder = targetStore.GetRootFolder();
                    dynamic smartCleanup = rootFolder.Folders[SmartCleanupFolderName];
                    dynamic targetFolder = smartCleanup.Folders[subFolderName];
                    count = (int)targetFolder.Items.Count;
                }
                catch (Exception ex)
                {
                    _log.Warn("Error getting folder count for {Folder}: {Error}", subFolderName, ex.Message);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromSeconds(15));
            return count;
        });
    }

    /// <summary>
    /// Bulk-moves a list of emails (by EntryId) into a triage subfolder.
    /// Returns count of successfully moved items.
    /// </summary>
    public async Task<int> BulkMoveToTriageFolderAsync(
        List<string> entryIds, string emailAddress, string storeName, string subFolderName,
        IProgress<string>? progress = null)
    {
        return await Task.Run(() =>
        {
            int moved = 0;
            var thread = new Thread(() =>
            {
                try
                {
                    dynamic outlookApp = GetActiveComObject("Outlook.Application");
                    dynamic session = outlookApp.GetNamespace("MAPI");
                    session.Logon("", "", false, false);

                    dynamic? targetStore = FindStore(session, emailAddress, storeName);
                    if (targetStore is null) { _log.Warn("BulkMove: Store not found"); return; }

                    // Ensure folder path exists
                    dynamic rootFolder = targetStore.GetRootFolder();
                    dynamic? smartCleanup = null;
                    try { smartCleanup = rootFolder.Folders[SmartCleanupFolderName]; } catch { }
                    if (smartCleanup is null) smartCleanup = rootFolder.Folders.Add(SmartCleanupFolderName);

                    dynamic? targetFolder = null;
                    try { targetFolder = smartCleanup.Folders[subFolderName]; } catch { }
                    if (targetFolder is null) targetFolder = smartCleanup.Folders.Add(subFolderName);

                    for (int i = 0; i < entryIds.Count; i++)
                    {
                        try
                        {
                            dynamic item = session.GetItemFromID(entryIds[i], targetStore.StoreID);
                            item.Move(targetFolder);
                            moved++;
                            if ((i + 1) % 10 == 0)
                                progress?.Report($"Moving emails... {i + 1}/{entryIds.Count}");
                        }
                        catch (Exception ex)
                        {
                            _log.Warn("BulkMove: skip item {Index}: {Error}", i, ex.Message);
                        }
                    }

                    _log.Info("BulkMove complete: {Moved}/{Total} to {Folder}", moved, entryIds.Count, subFolderName);
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error during bulk move to {Folder}", subFolderName);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromMinutes(5));
            return moved;
        });
    }

    /// <summary>
    /// Bulk-moves all remaining items from a triage subfolder back to Inbox.
    /// Returns count of successfully moved items.
    /// </summary>
    public async Task<int> BulkMoveAllToInboxAsync(
        string emailAddress, string storeName, string subFolderName,
        IProgress<string>? progress = null)
    {
        return await Task.Run(() =>
        {
            int moved = 0;
            var thread = new Thread(() =>
            {
                try
                {
                    dynamic outlookApp = GetActiveComObject("Outlook.Application");
                    dynamic session = outlookApp.GetNamespace("MAPI");
                    session.Logon("", "", false, false);

                    dynamic? targetStore = FindStore(session, emailAddress, storeName);
                    if (targetStore is null) return;

                    dynamic rootFolder = targetStore.GetRootFolder();
                    dynamic smartCleanup = rootFolder.Folders[SmartCleanupFolderName];
                    dynamic sourceFolder = smartCleanup.Folders[subFolderName];
                    dynamic inboxFolder = targetStore.GetDefaultFolder(olFolderInbox);
                    dynamic items = sourceFolder.Items;

                    // Move in reverse order (COM collections shift indices on removal)
                    int count = (int)items.Count;
                    for (int i = count; i >= 1; i--)
                    {
                        try
                        {
                            dynamic item = items[i];
                            item.Move(inboxFolder);
                            moved++;
                            if (moved % 10 == 0)
                                progress?.Report($"Moving back to Inbox... {moved}/{count}");
                        }
                        catch (Exception ex)
                        {
                            _log.Warn("BulkMoveToInbox: skip item {Index}: {Error}", i, ex.Message);
                        }
                    }

                    _log.Info("BulkMoveToInbox complete: {Moved} from {Folder}", moved, subFolderName);
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error during bulk move to inbox from {Folder}", subFolderName);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromMinutes(5));
            return moved;
        });
    }

    private static string ExtractDomainFromEmail(string? email)
    {
        if (string.IsNullOrWhiteSpace(email) || !email.Contains('@')) return "";
        return email.ToLowerInvariant().Split('@')[1];
    }

    /// <summary>
    /// Helper: finds the store matching the given email/store name across all stores.
    /// </summary>
    private dynamic? FindStore(dynamic session, string emailAddress, string storeName)
    {
        dynamic stores = session.Stores;
        int count = (int)stores.Count;
        for (int i = 1; i <= count; i++)
        {
            try
            {
                dynamic store = stores[i];
                string sName = (string)store.DisplayName;
                if (sName == storeName || sName == emailAddress || sName.Contains(emailAddress))
                    return store;
            }
            catch { }
        }
        return null;
    }

    private dynamic? GetStoreForAccount(dynamic account)
    {
        try
        {
            return account.DeliveryStore;
        }
        catch
        {
            if (_session is null) return null;
            try
            {
                string acctName = (string)account.DisplayName;
                dynamic stores = _session.Stores;
                int count = (int)stores.Count;
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic store = stores[i];
                        if ((string)store.DisplayName == acctName)
                            return store;
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }
    }

    private string GetStoreNameForAccount(dynamic account)
    {
        try { return (string)(GetStoreForAccount(account)?.DisplayName ?? ""); }
        catch { return ""; }
    }

    private string GetStoreFilePathForAccount(dynamic account)
    {
        try { return (string)(GetStoreForAccount(account)?.FilePath ?? ""); }
        catch { return ""; }
    }

    private static string GetProviderHint(dynamic account)
    {
        try
        {
            int accountType = (int)account.AccountType;
            return accountType switch
            {
                olExchange => "Exchange",
                olImap => "IMAP",
                olPop3 => "POP3",
                olHttp => "Outlook.com",
                _ => "Unknown"
            };
        }
        catch { return "Unknown"; }
    }

    private static string GetProviderHintFromStore(dynamic store)
    {
        try
        {
            int storeType = (int)store.ExchangeStoreType;
            if (storeType == olPrimaryExchangeMailbox || storeType == olAdditionalExchangeMailbox)
                return "Exchange";

            try
            {
                string fp = (string)store.FilePath;
                if (fp.EndsWith(".pst", StringComparison.OrdinalIgnoreCase))
                    return "POP3/IMAP";
            }
            catch { }

            return "Unknown";
        }
        catch { return "Unknown"; }
    }

    private static bool CheckPatternFileExists(string accountKey)
    {
        var patternPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "EmailManage", "accounts", accountKey, "pattern_v1.yaml");
        return File.Exists(patternPath);
    }

    private static bool IsOutlookInstalled()
    {
        try
        {
            var outlookType = Type.GetTypeFromProgID("Outlook.Application");
            return outlookType is not null;
        }
        catch { return false; }
    }

    private static bool IsOutlookRunning()
    {
        return Process.GetProcessesByName("OUTLOOK").Length > 0;
    }

    /// <summary>
    /// Replacement for Marshal.GetActiveObject which was removed in .NET 5+.
    /// Uses OLE32 GetActiveObject via P/Invoke.
    /// </summary>
    private static object GetActiveComObject(string progId)
    {
        var clsid = Type.GetTypeFromProgID(progId, true)!.GUID;
        int hr = GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
        if (hr < 0)
            Marshal.ThrowExceptionForHR(hr);
        return obj;
    }

    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    /// <summary>
    /// Disconnects from Outlook COM objects and releases resources.
    /// </summary>
    public void Disconnect()
    {
        if (_session is not null)
        {
            try
            {
                _session.Logoff();
                Marshal.ReleaseComObject(_session);
            }
            catch { /* best effort */ }
            _session = null;
        }

        if (_outlookApp is not null)
        {
            try { Marshal.ReleaseComObject(_outlookApp); }
            catch { /* best effort */ }
            _outlookApp = null;
        }

        _log.Info("Outlook disconnected.");
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            Disconnect();
            _disposed = true;
        }
    }
}
