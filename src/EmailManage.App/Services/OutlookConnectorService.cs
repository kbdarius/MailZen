using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using EmailManage.Models;

namespace EmailManage.Services;

/// <summary>
/// Implements IMessageFilter to automatically retry COM calls that are rejected
/// because the server (Outlook) is busy. Without this, RPC_E_CALL_REJECTED crashes the app.
/// </summary>
internal class RetryMessageFilter : IOleMessageFilter
{
    public static void Register()
    {
        IOleMessageFilter newFilter = new RetryMessageFilter();
        CoRegisterMessageFilter(newFilter, out _);
    }

    public static void Revoke()
    {
        CoRegisterMessageFilter(null, out _);
    }

    // IOleMessageFilter methods
    int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) => 0; // SERVERCALL_ISHANDLED

    int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    {
        if (dwRejectType == 2) // SERVERCALL_RETRYLATER
            return 99; // Retry after ~100ms
        return -1; // Cancel
    }

    int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) => 2; // PENDINGMSG_WAITDEFPROCESS

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter? lpMessageFilter, out IOleMessageFilter? lplpMessageFilter);
}

[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig]
    int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig]
    int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

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
    public const int OlFolderInbox = 6;
    public const int OlFolderDeletedItems = 3;

    // Rule constants
    private const int olRuleReceive = 0;
    private const int olRuleActionMoveToFolder = 1;

    // ── Folder Names ──
    public const string ReviewForDeletionFolderName = "Review for Deletion";

    // Legacy folder names (kept for migration/cleanup)
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

    // ── Sandbox Method (Data Harvester) ──
    public async Task<List<EmailMessageInfo>> GetRecentEmailsAsync(string storeId, int folderType, int maxCount)
    {
        var results = new List<EmailMessageInfo>();

        await Task.Run(() =>
        {
            // Thread in STA for COM
            var thread = new Thread(() =>
            {
                // We create a new local session because we are on a new STA thread.
                dynamic? localApp = null;
                dynamic? localSession = null;
                dynamic? store = null;
                dynamic? folder = null;
                dynamic? items = null;

                try
                {
                    localApp = GetActiveComObject("Outlook.Application");
                    localSession = localApp.GetNamespace("MAPI");
                    localSession.Logon("", "", false, false);
                    
                    _log.Info($"GetRecentEmailsAsync: Looking for StoreID '{storeId}'...");

                    if (!string.IsNullOrEmpty(storeId))
                    {
                        try { store = localSession.GetStoreFromID(storeId); }
                        catch (Exception ex) { _log.Warn("GetStoreFromID failed: {0}", ex.Message); }
                    }

                    if (store == null)
                    {
                        _log.Info("GetRecentEmailsAsync: Fallback search by ID...");
                        try
                        {
                            dynamic allStores = localSession.Stores;
                            int countS = allStores.Count;
                            for (int k=1; k<=countS; k++)
                            {
                                dynamic s = allStores[k];
                                string sid = "";
                                try { sid = (string)s.StoreID; } catch {}
                                
                                if (!string.IsNullOrEmpty(sid) && sid == storeId)
                                {
                                    store = s;
                                    break;
                                }
                                if (s != null && s != store) Marshal.ReleaseComObject(s);
                            }
                            if (allStores != null) Marshal.ReleaseComObject(allStores);
                        }
                        catch (Exception ex) { _log.Error(ex, "Fallback loop failed."); }
                    }
                    
                    if (store == null) 
                    {
                        _log.Warn("GetRecentEmailsAsync: Store NOT found.");
                        return;
                    }

                    folder = store.GetDefaultFolder(folderType);
                    items = folder.Items;
                    // Use GetFirst/GetNext for better performance on large folders
                    // Sort descending by ReceivedTime
                    items.Sort("[ReceivedTime]", true);
                    
                    dynamic item = items.GetFirst();
                    int countFound = 0;
                    
                    while (item != null && countFound < maxCount)
                    {
                        try 
                        {
                            // Verify it's a MailItem (Class 43)
                            int cls = 43; // Default or assume? No, check property
                            try { cls = (int)item.Class; } catch {}

                            if (cls == 43) 
                            {
                                var info = new EmailMessageInfo
                                {
                                    // Use safe accessors with try-catch per property if needed
                                    // But typically if item is valid, these work 
                                    Subject = "", 
                                    SenderName = "",
                                    Body = ""
                                };

                                try { info.EntryId = (string)item.EntryID; } catch {}
                                try { info.Subject = (string)item.Subject ?? "(No Subject)"; } catch {}
                                try { info.SenderName = (string)item.SenderName ?? ""; } catch {}
                                try { info.Body = (string)item.Body ?? ""; } catch {}
                                try { info.ReceivedTime = (DateTime)item.ReceivedTime; } catch { info.ReceivedTime = DateTime.MinValue; }

                                // Get Sender Address with helper (handles Exchange / O=...)
                                try { info.SenderEmailAddress = GetSenderAddress(item); } catch { info.SenderEmailAddress = ""; }

                                results.Add(info);
                                countFound++;
                            }
                        }
                        catch (Exception exItem) 
                        {
                            // Log but continue
                             _log.Warn($"Item processing error: {exItem.Message}");
                        }

                        // Move to next item safely
                        dynamic nextItem = null;
                        try 
                        {
                            nextItem = items.GetNext(); 
                        } 
                        catch 
                        {
                            // If GetNext fails, stop
                            nextItem = null;
                        }

                        if (item != null) Marshal.ReleaseComObject(item);
                        item = nextItem;
                    }
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Failed to get recent emails from Sandbox.");
                }
                finally
                {
                    if (items != null) Marshal.ReleaseComObject(items);
                    if (folder != null) Marshal.ReleaseComObject(folder);
                    if (store != null) Marshal.ReleaseComObject(store);
                    if (localSession != null) Marshal.ReleaseComObject(localSession);
                    if (localApp != null) Marshal.ReleaseComObject(localApp);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        });
        
        return results;
    }

    public async Task<List<EmailMessageInfo>> GetEmailsByStoreIdAsync(
        string storeId,
        int folderType,
        DateTime startDate,
        DateTime endDate,
        int maxItems = 5000)
    {
        return await Task.Run(() =>
        {
            var results = new List<EmailMessageInfo>();
            Exception? staException = null;

            var thread = new Thread(() =>
            {
                dynamic? localApp = null;
                dynamic? localSession = null;
                dynamic? store = null;
                dynamic? folder = null;
                dynamic? items = null;
                dynamic? restricted = null;

                try
                {
                    RetryMessageFilter.Register();

                    localApp = GetActiveComObject("Outlook.Application");
                    localSession = localApp.GetNamespace("MAPI");
                    localSession.Logon("", "", false, false);

                    try { store = localSession.GetStoreFromID(storeId); }
                    catch
                    {
                        dynamic stores = localSession.Stores;
                        int count = stores.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            dynamic candidate = stores[i];
                            try
                            {
                                string sid = candidate.StoreID ?? "";
                                if (sid.Equals(storeId, StringComparison.OrdinalIgnoreCase))
                                {
                                    store = candidate;
                                    break;
                                }
                            }
                            catch
                            {
                                Marshal.ReleaseComObject(candidate);
                            }
                        }
                    }

                    if (store == null)
                        throw new Exception($"Could not find Outlook store '{storeId}'.");

                    folder = store.GetDefaultFolder(folderType);
                    items = folder.Items;
                    items.Sort("[ReceivedTime]", true);

                    string startS = startDate.ToString("g");
                    string endS = endDate.ToString("g");
                    string filter = $"[ReceivedTime] >= '{startS}' AND [ReceivedTime] <= '{endS}'";

                    try { restricted = items.Restrict(filter); }
                    catch { restricted = items; }

                    dynamic item = restricted.GetFirst();
                    int fetched = 0;
                    while (item != null && fetched < maxItems)
                    {
                        try
                        {
                            if ((int)item.Class == 43)
                            {
                                DateTime received = (DateTime)item.ReceivedTime;
                                string senderEmail = GetSenderAddress(item);

                                results.Add(new EmailMessageInfo
                                {
                                    EntryId = (string)(item.EntryID ?? ""),
                                    Subject = (string)(item.Subject ?? ""),
                                    SenderName = (string)(item.SenderName ?? ""),
                                    SenderEmailAddress = senderEmail ?? "",
                                    ReceivedTime = received,
                                    UnRead = (bool)(item.UnRead ?? false),
                                    IsDeleted = folderType == OlFolderDeletedItems
                                });
                                fetched++;
                            }
                        }
                        catch
                        {
                            // Skip malformed items and continue scanning.
                        }
                        finally
                        {
                            dynamic next = restricted.GetNext();
                            Marshal.ReleaseComObject(item);
                            item = next;
                        }
                    }
                }
                catch (Exception ex)
                {
                    staException = ex;
                }
                finally
                {
                    if (restricted != null && !ReferenceEquals(restricted, items))
                        Marshal.ReleaseComObject(restricted);
                    if (items != null) Marshal.ReleaseComObject(items);
                    if (folder != null) Marshal.ReleaseComObject(folder);
                    if (store != null) Marshal.ReleaseComObject(store);
                    if (localSession != null) Marshal.ReleaseComObject(localSession);
                    if (localApp != null) Marshal.ReleaseComObject(localApp);
                    RetryMessageFilter.Revoke();
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (staException != null)
                throw new Exception($"Could not read Outlook folder state: {staException.Message}", staException);

            return results;
        });
    }

    private string GetSenderAddress(dynamic mailItem)
    {
        try
        {
            string address = mailItem.SenderEmailAddress;
            if (string.IsNullOrEmpty(address)) return "";

            // If it's an Exchange address (/o=...), try to resolve it
            if (address.Contains("/o=") || address.Contains("/O="))
            {
                dynamic sender = mailItem.Sender;
                if (sender != null)
                {
                    dynamic exchangeUser = sender.GetExchangeUser();
                    if (exchangeUser != null)
                    {
                        address = exchangeUser.PrimarySmtpAddress;
                        Marshal.ReleaseComObject(exchangeUser);
                    }
                    Marshal.ReleaseComObject(sender);
                }
            }
            return address ?? "";
        }
        catch { return ""; }
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
                        StoreID = GetDefaultStoreID(acct), // Added for Sandbox
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
                                dynamic inbox = store.GetDefaultFolder(OlFolderInbox);
                                info.InboxCount = (int)inbox.Items.Count;
                            }
                            catch { }

                            try
                            {
                                dynamic deleted = store.GetDefaultFolder(OlFolderDeletedItems);
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
                            StoreID = (string)store.StoreID, // Added for Sandbox
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

    public async Task<List<EmailMessageInfo>> GetEmailsAsync(string emailAddress, string storeName, int folderType, int maxMonths, int maxItems = 1000, bool skipUnread = true)
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

                    dynamic folder = targetStore.GetDefaultFolder(folderType);
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
                            if (skipUnread && folderType == OlFolderInbox && unread) continue;

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
                                IsDeleted = folderType == OlFolderDeletedItems
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

                    dynamic deletedFolder = targetStore.GetDefaultFolder(OlFolderDeletedItems);
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

                    dynamic inboxFolder = targetStore.GetDefaultFolder(OlFolderInbox);
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
                    dynamic inboxFolder = targetStore.GetDefaultFolder(OlFolderInbox);
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

    // ═══════════════════════════════════════════════════════════════
    // New "Review for Deletion" top-level folder methods
    // ═══════════════════════════════════════════════════════════════

    /// <summary>
    /// Creates the "Review for Deletion" folder at the top level of the account store.
    /// Returns true on success.
    /// </summary>
    public async Task<bool> EnsureReviewFolderAsync(string emailAddress, string storeName)
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

                    dynamic? targetStore = FindStore(session, emailAddress, storeName);
                    if (targetStore is null) { _log.Warn("EnsureReviewFolder: Store not found"); return; }

                    dynamic rootFolder = targetStore.GetRootFolder();
                    dynamic folders = rootFolder.Folders;

                    dynamic? reviewFolder = null;
                    try { reviewFolder = folders[ReviewForDeletionFolderName]; } catch { }

                    if (reviewFolder is null)
                    {
                        reviewFolder = folders.Add(ReviewForDeletionFolderName);
                        _log.Info("Created top-level folder: {FolderName}", ReviewForDeletionFolderName);
                    }

                    success = true;
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error ensuring Review for Deletion folder exists");
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromSeconds(30));
            return success;
        });
    }

    /// <summary>
    /// Takes a snapshot of all items in the top-level "Review for Deletion" folder.
    /// </summary>
    public async Task<List<FolderEmailSnapshot>> GetReviewFolderSnapshotAsync(
        string emailAddress, string storeName, int maxItems = 1000)
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
                    if (targetStore is null) { _log.Warn("GetReviewSnapshot: Store not found"); return; }

                    dynamic rootFolder = targetStore.GetRootFolder();
                    dynamic reviewFolder = rootFolder.Folders[ReviewForDeletionFolderName];
                    dynamic items = reviewFolder.Items;
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
                                        senderEmail = (string)pa.GetProperty(
                                            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E");
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
                        catch (Exception ex) { _log.Warn("ReviewSnapshot: skip item {Index}: {Error}", i, ex.Message); }
                    }

                    _log.Info("Review snapshot taken: {Count} items", snapshot.Count);
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error taking Review for Deletion snapshot");
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromMinutes(2));
            return snapshot;
        });
    }

    /// <summary>
    /// Bulk-moves emails (by EntryId) into the top-level "Review for Deletion" folder.
    /// Returns count of successfully moved items.
    /// </summary>
    public async Task<int> BulkMoveToReviewFolderAsync(
        List<string> entryIds, string emailAddress, string storeName,
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
                    if (targetStore is null) { _log.Warn("BulkMoveToReview: Store not found"); return; }

                    dynamic rootFolder = targetStore.GetRootFolder();
                    dynamic? reviewFolder = null;
                    try { reviewFolder = rootFolder.Folders[ReviewForDeletionFolderName]; } catch { }
                    if (reviewFolder is null) reviewFolder = rootFolder.Folders.Add(ReviewForDeletionFolderName);

                    for (int i = 0; i < entryIds.Count; i++)
                    {
                        try
                        {
                            dynamic item = session.GetItemFromID(entryIds[i], targetStore.StoreID);
                            item.Move(reviewFolder);
                            moved++;
                            if ((i + 1) % 10 == 0)
                                progress?.Report($"Moving to review... {i + 1}/{entryIds.Count}");
                        }
                        catch (Exception ex)
                        {
                            _log.Warn("BulkMoveToReview: skip item {Index}: {Error}", i, ex.Message);
                        }
                    }

                    _log.Info("BulkMoveToReview complete: {Moved}/{Total}", moved, entryIds.Count);
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error during bulk move to Review for Deletion");
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromMinutes(5));
            return moved;
        });
    }

    /// <summary>
    /// Moves all remaining items from "Review for Deletion" back to Inbox.
    /// Returns count of successfully moved items.
    /// </summary>
    public async Task<int> BulkMoveReviewToInboxAsync(
        string emailAddress, string storeName, IProgress<string>? progress = null)
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
                    dynamic reviewFolder = rootFolder.Folders[ReviewForDeletionFolderName];
                    dynamic inboxFolder = targetStore.GetDefaultFolder(OlFolderInbox);
                    dynamic items = reviewFolder.Items;

                    int count = (int)items.Count;
                    // Move in reverse order — COM collections shift indices on removal
                    for (int i = count; i >= 1; i--)
                    {
                        try
                        {
                            dynamic item = items[i];
                            item.Move(inboxFolder);
                            moved++;
                            if (moved % 10 == 0)
                                progress?.Report($"Returning to Inbox... {moved}/{count}");
                        }
                        catch (Exception ex)
                        {
                            _log.Warn("ReviewToInbox: skip item {Index}: {Error}", i, ex.Message);
                        }
                    }

                    _log.Info("ReviewToInbox complete: {Moved} items returned", moved);
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error moving Review for Deletion items to Inbox");
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromMinutes(5));
            return moved;
        });
    }

    /// <summary>
    /// Gets the count of items in the "Review for Deletion" folder.
    /// Returns 0 if folder doesn't exist.
    /// </summary>
    public async Task<int> GetReviewFolderCountAsync(string emailAddress, string storeName)
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
                    dynamic reviewFolder = rootFolder.Folders[ReviewForDeletionFolderName];
                    count = (int)reviewFolder.Items.Count;
                }
                catch { /* folder may not exist yet — that's fine */ }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromSeconds(15));
            return count;
        });
    }

    /// <summary>
    /// Creates Outlook rules that move emails from the given sender addresses
    /// to Deleted Items. Batches senders into rules (max 10 per rule) to stay
    /// within Outlook's rule-size limits. Returns the number of rules created.
    /// </summary>
    public async Task<int> CreateSenderRulesAsync(
        string emailAddress, string storeName,
        List<string> senderAddresses,
        IProgress<string>? progress = null)
    {
        return await Task.Run(() =>
        {
            int created = 0;
            var thread = new Thread(() =>
            {
                try
                {
                    dynamic outlookApp = GetActiveComObject("Outlook.Application");
                    dynamic session = outlookApp.GetNamespace("MAPI");
                    session.Logon("", "", false, false);

                    dynamic? targetStore = FindStore(session, emailAddress, storeName);
                    if (targetStore is null)
                    {
                        _log.Warn("CreateSenderRules: Store not found");
                        return;
                    }

                    dynamic deletedFolder = targetStore.GetDefaultFolder(OlFolderDeletedItems);
                    dynamic rules = targetStore.GetRules();

                    // Collect existing MailZen rule names to avoid duplicates
                    var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    int ruleCount = (int)rules.Count;
                    for (int i = 1; i <= ruleCount; i++)
                    {
                        try
                        {
                            string name = (string)(rules[i].Name ?? "");
                            if (name.StartsWith("MailZen:", StringComparison.OrdinalIgnoreCase))
                                existingNames.Add(name);
                        }
                        catch { }
                    }

                    // Deduplicate and batch senders (max 10 per rule)
                    var uniqueSenders = senderAddresses
                        .Where(s => !string.IsNullOrWhiteSpace(s))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    const int batchSize = 10;
                    int batchIndex = 0;

                    for (int start = 0; start < uniqueSenders.Count; start += batchSize)
                    {
                        var batch = uniqueSenders.Skip(start).Take(batchSize).ToArray();
                        batchIndex++;

                        // Generate a unique rule name
                        var ruleName = batch.Length == 1
                            ? $"MailZen: {batch[0]}"
                            : $"MailZen: auto-delete batch {batchIndex} ({batch.Length} senders)";

                        int suffix = 1;
                        var baseName = ruleName;
                        while (existingNames.Contains(ruleName))
                            ruleName = $"{baseName} #{suffix++}";

                        try
                        {
                            // Create a receive rule
                            dynamic rule = rules.Create(ruleName, olRuleReceive);

                            // Condition: sender address matches
                            dynamic senderCondition = rule.Conditions.SenderAddress;
                            senderCondition.Address = batch;
                            senderCondition.Enabled = true;

                            // Action: move to Deleted Items
                            dynamic moveAction = rule.Actions.MoveToFolder;
                            moveAction.Folder = deletedFolder;
                            moveAction.Enabled = true;

                            rule.Enabled = true;
                            created++;
                            existingNames.Add(ruleName);

                            progress?.Report($"Created rule: {ruleName}");
                            _log.Info("Created Outlook rule [{Name}] for {Count} senders",
                                ruleName, batch.Length);
                        }
                        catch (Exception ex)
                        {
                            _log.Warn("Failed to create rule [{Name}]: {Error}",
                                ruleName, ex.Message);
                        }
                    }

                    // Persist all rules to Outlook
                    if (created > 0)
                    {
                        rules.Save();
                        _log.Info("Saved {Count} new Outlook rules", created);
                    }
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Error creating Outlook sender rules");
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(TimeSpan.FromMinutes(2));
            return created;
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

    private string GetDefaultStoreID(dynamic account)
    {
        try 
        { 
            dynamic store = GetStoreForAccount(account);
            if(store != null) return (string)store.StoreID;
        }
        catch { }
        return "";
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

    // ═══════════════════════════════════════════════════════════════════
    // Scoring System — Sender Reputation + Junk Score
    // ═══════════════════════════════════════════════════════════════════

    /// <summary>Aggregated statistics per unique sender, built during extraction.</summary>
    private class SenderStats
    {
        public string SenderEmail { get; set; } = "";
        public string SenderName { get; set; } = "";
        public int TotalEmails { get; set; }
        public int DeletedCount { get; set; }
        public int ReadCount { get; set; }
        public int UnreadCount { get; set; }
        public int HasUnsubscribeCount { get; set; }
        public int IsBulkCount { get; set; }
        public int HasFeedbackIdCount { get; set; }
        public int RepliedToCount { get; set; } // sender email appears in Sent Items
        public int FocusedCount { get; set; }
        public int OtherCount { get; set; }
    }

    /// <summary>Calculate a 0-100 reputation score for a sender (50 = neutral, 100 = suspicious).</summary>
    private static int CalculateSenderReputation(SenderStats stats)
    {
        double score = 50; // neutral

        if (stats.TotalEmails > 0)
        {
            // Delete ratio: if you delete most of their emails → bad sender
            double deleteRatio = (double)stats.DeletedCount / stats.TotalEmails;
            score += deleteRatio * 40; // max +40

            // Read ratio: if you read most → good sender
            double readRatio = (double)stats.ReadCount / stats.TotalEmails;
            score -= readRatio * 20; // max -20
        }

        // All emails have unsubscribe header → commercial sender
        if (stats.TotalEmails > 0 && stats.HasUnsubscribeCount == stats.TotalEmails)
            score += 20;

        // You've replied to (or sent to) this sender → trusted
        if (stats.RepliedToCount > 0)
            score -= 30;

        // Microsoft classified most as "Other"
        int classifiedTotal = stats.FocusedCount + stats.OtherCount;
        if (classifiedTotal > 0)
        {
            double otherRatio = (double)stats.OtherCount / classifiedTotal;
            score += otherRatio * 15;
        }

        return (int)Math.Max(0, Math.Min(100, score));
    }

    /// <summary>Calculate a 0-100 junk score for an individual email.</summary>
    private static int CalculateJunkScore(
        bool hasUnsubscribe, bool isBulk, string focusedOrOther,
        bool hasFeedbackId, bool isMailingList, bool isRead,
        bool dmarcPass, bool spfPass, bool dkimPass,
        string folder, DateTime receivedTime, int senderReputation)
    {
        double score = 0;

        // Signal-based scoring
        if (hasUnsubscribe) score += 30;
        if (isBulk) score += 25;
        if (focusedOrOther.Equals("Other", StringComparison.OrdinalIgnoreCase)) score += 20;
        if (hasFeedbackId) score += 15;
        if (isMailingList) score += 15;

        // Unread old email (> 7 days and never opened)
        if (!isRead && (DateTime.Now - receivedTime).TotalDays > 7) score += 10;

        // Authentication failures
        if (!dmarcPass) score += 10;
        if (!spfPass) score += 5;
        if (!dkimPass) score += 5;

        // Behavioral / folder signals
        if (folder.Equals("Deleted Items", StringComparison.OrdinalIgnoreCase)) score += 35;
        if (folder.Equals("Sent Items", StringComparison.OrdinalIgnoreCase)) score -= 50;
        if (isRead) score -= 10;

        // Blend in sender reputation (0-100 scale, 50 = neutral)
        score += (senderReputation - 50) * 0.3;

        return (int)Math.Max(0, Math.Min(100, score));
    }

    /// <summary>Map a junk score to a recommendation label.</summary>
    private static string GetRecommendation(int junkScore)
    {
        if (junkScore <= 30) return "Keep";
        if (junkScore <= 60) return "Review";
        return "Delete";
    }

    private static int ApplyLearnedProfileAdjustments(int junkScore, string senderEmail, LearnedProfile? profile)
    {
        if (profile == null)
            return junkScore;

        var sender = (senderEmail ?? "").Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(sender))
            return junkScore;

        if (profile.DoNotDeleteSenders.Contains(sender))
            return Math.Min(junkScore, 15);

        if (profile.DeletedSenderCounts.TryGetValue(sender, out var senderDeletes))
            junkScore += Math.Min(30, senderDeletes * 8);

        var domain = ExtractDomainFromEmail(sender);
        if (!string.IsNullOrEmpty(domain) && profile.DeletedDomainCounts.TryGetValue(domain, out var domainDeletes))
            junkScore += Math.Min(15, domainDeletes * 3);

        return Math.Clamp(junkScore, 0, 100);
    }

    // ═══════════════════════════════════════════════════════════════════
    // Outlook Category Labelling
    // ═══════════════════════════════════════════════════════════════════

    /// <summary>Ensure MailZen categories exist in Outlook.</summary>
    private static void EnsureCategoriesExist(dynamic ns)
    {
        try
        {
            dynamic categories = ns.Categories;
            bool hasKeep = false, hasReview = false, hasDelete = false;

            foreach (dynamic cat in categories)
            {
                string name = cat.Name;
                if (name == "MailZen: Keep") hasKeep = true;
                else if (name == "MailZen: Review") hasReview = true;
                else if (name == "MailZen: Delete") hasDelete = true;
            }

            // OlCategoryColor enum: 5 = Green, 14 = Yellow, 1 = Red
            if (!hasKeep) categories.Add("MailZen: Keep", 5);
            if (!hasReview) categories.Add("MailZen: Review", 14);
            if (!hasDelete) categories.Add("MailZen: Delete", 1);
        }
        catch { /* Categories API not available on some stores */ }
    }

    /// <summary>Apply MailZen color categories to emails in Outlook.</summary>
    private static void ApplyCategoriesToOutlook(
        dynamic ns,
        List<Dictionary<string, object>> allEmails,
        IProgress<string> progress,
        CancellationToken cancellationToken)
    {
        progress.Report("Applying color categories in Outlook...");

        EnsureCategoriesExist(ns);

        int applied = 0;
        int errors = 0;

        for (int i = 0; i < allEmails.Count; i++)
        {
            if (cancellationToken.IsCancellationRequested) break;

            try
            {
                string entryId = (string)allEmails[i]["EntryID"];
                string recommendation = (string)allEmails[i]["Recommendation"];

                if (string.IsNullOrEmpty(entryId)) continue;

                dynamic mailItem = ns.GetItemFromID(entryId);
                if (mailItem == null) continue;

                try
                {
                    // Remove any existing MailZen categories, keep user categories
                    string existing = mailItem.Categories ?? "";
                    var cats = existing.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(c => c.Trim())
                        .Where(c => !c.StartsWith("MailZen:"))
                        .ToList();

                    string newCategory = recommendation switch
                    {
                        "Keep" => "MailZen: Keep",
                        "Review" => "MailZen: Review",
                        "Delete" => "MailZen: Delete",
                        _ => ""
                    };

                    if (!string.IsNullOrEmpty(newCategory))
                    {
                        cats.Add(newCategory);
                        mailItem.Categories = string.Join(", ", cats);
                        mailItem.Save();
                        applied++;
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(mailItem);
                }
            }
            catch
            {
                errors++;
            }

            if ((i + 1) % 100 == 0)
                progress.Report($"Labelling... {applied} tagged ({errors} skipped) of {allEmails.Count}");
        }

        progress.Report($"Labelling complete: {applied} tagged, {errors} skipped.");
    }

    // ═══════════════════════════════════════════════════════════════════
    // Dataset Export (Two-Pass: Extract → Score → Write)
    // ═══════════════════════════════════════════════════════════════════

    /// <summary>
    /// Exports a dataset of emails to CSV + XLSX with scoring.
    /// Pass 1: Extract all emails and build sender stats.
    /// Pass 2: Calculate sender reputation + per-email junk scores.
    /// Pass 3 (optional): Apply color categories in Outlook.
    /// Pass 4: Write CSV + 3-sheet XLSX (Dataset, Sender Report, Column Guide).
    /// </summary>
    public async Task ExportDataset(
        IEnumerable<string> accountIds,
        IReadOnlyDictionary<string, string>? accountKeysByStoreId,
        DateTime startDate,
        DateTime endDate,
        bool includeInbox,
        bool includeSent,
        bool includeDeleted,
        bool includeRead,
        bool applyOutlookCategories,
        bool saveXlsx,
        string outputFilePath,
        IProgress<string> progress,
        CancellationToken cancellationToken)
    {
        await Task.Run(() =>
        {
            // MAPI property tags
            const string PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

            // Focused Inbox — comprehensive list of MAPI property paths to try
            string[] focusedPropertyPaths = new[]
            {
                "http://schemas.microsoft.com/mapi/string/{00062008-0000-0000-C000-000000000046}/InferenceClassification",
                "http://schemas.microsoft.com/mapi/string/{00062008-0000-0000-C000-000000000046}/InferenceClassificationResult",
                "http://schemas.microsoft.com/mapi/string/{00062008-0000-0000-C000-000000000046}/InferenceClassificationOverride",
                "http://schemas.microsoft.com/mapi/string/{00062008-0000-0000-C000-000000000046}/IsClutter",
                "http://schemas.microsoft.com/mapi/string/{23239608-685D-4732-9C55-4C95CB4E8E33}/InferenceClassification",
                "http://schemas.microsoft.com/mapi/string/{23239608-685D-4732-9C55-4C95CB4E8E33}/InferenceClassificationResult",
                "http://schemas.microsoft.com/mapi/string/{23239608-685D-4732-9C55-4C95CB4E8E33}/InferredClass",
                "http://schemas.microsoft.com/mapi/string/{23239608-685D-4732-9C55-4C95CB4E8E33}/OverrideClass",
                "http://schemas.microsoft.com/mapi/string/{23239608-685D-4732-9C55-4C95CB4E8E33}/IsClutter",
                "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/InferenceClassification",
                "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/InferenceClassificationResult",
                "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/InferenceClassification",
                "http://schemas.microsoft.com/mapi/string/{41F28F13-83F4-4114-A584-EEDB5A6B0BFF}/InferenceClassification",
                "http://schemas.microsoft.com/mapi/string/{41F28F13-83F4-4114-A584-EEDB5A6B0BFF}/InferenceClassificationResult",
            };

            // ══════════════════════════════════════════════════════════════
            // Two-pass approach: collect all emails first, then score them
            // ══════════════════════════════════════════════════════════════
            var allEmails = new List<Dictionary<string, object>>();
            var senderStatsMap = new Dictionary<string, SenderStats>(StringComparer.OrdinalIgnoreCase);
            var learnedProfilesByAccountKey = new Dictionary<string, LearnedProfile>(StringComparer.OrdinalIgnoreCase);

            if (accountKeysByStoreId != null)
            {
                foreach (var accountKey in accountKeysByStoreId.Values.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    var profile = LearnedProfile.Load(accountKey);
                    if (profile != null)
                        learnedProfilesByAccountKey[accountKey] = profile;
                }
            }

            // COM STA Thread
            Exception? staException = null;
            dynamic? staNamespace = null; // keep ref for category labelling

            var thread = new Thread(() =>
            {
                dynamic? app = null;
                dynamic? ns = null;

                try
                {
                    RetryMessageFilter.Register();

                    // Connect to Outlook
                    const int maxRetries = 3;
                    for (int attempt = 1; attempt <= maxRetries; attempt++)
                    {
                        try
                        {
                            app = GetActiveComObject("Outlook.Application");
                            break;
                        }
                        catch
                        {
                            if (attempt == maxRetries)
                            {
                                try
                                {
                                    var outlookType = Type.GetTypeFromProgID("Outlook.Application", true)!;
                                    app = Activator.CreateInstance(outlookType);
                                }
                                catch (COMException)
                                {
                                    throw new Exception("Could not connect to Outlook. Make sure Outlook is open and not showing any dialog boxes, then try again.");
                                }
                            }
                            else
                            {
                                Thread.Sleep(1000 * attempt);
                            }
                        }
                    }
                    ns = app.GetNamespace("MAPI");
                    ns.Logon("", "", false, false);
                    staNamespace = ns;

                    // ══════════════════════════════════════════
                    // PASS 1: Extract all emails into memory
                    // ══════════════════════════════════════════
                    progress.Report("Pass 1: Extracting emails...");

                    int totalExtracted = 0;
                    dynamic stores = ns.Stores;
                    int storeCount = stores.Count;

                    for (int i = 1; i <= storeCount; i++)
                    {
                        if (cancellationToken.IsCancellationRequested) break;

                        dynamic store = stores[i];
                        string storeId = "";
                        string displayName = "";

                        try
                        {
                            storeId = store.StoreID;
                            displayName = store.DisplayName;
                        }
                        catch { continue; }

                        bool isSelected = false;
                        foreach (var id in accountIds)
                        {
                            if (id == storeId) { isSelected = true; break; }
                        }

                        if (!isSelected)
                        {
                            Marshal.ReleaseComObject(store);
                            continue;
                        }

                        string accountKey = "";
                        if (accountKeysByStoreId != null)
                            accountKeysByStoreId.TryGetValue(storeId, out accountKey);

                        progress.Report($"Pass 1: Scanning {displayName}...");

                        var foldersToScan = new List<(dynamic Folder, string Name)>();
                        dynamic root = store.GetRootFolder();

                        try
                        {
                            if (includeInbox)
                            {
                                dynamic f = root.Folders["Inbox"];
                                if (f == null) f = store.GetDefaultFolder(6);
                                if (f != null) foldersToScan.Add((f, "Inbox"));
                            }
                        } catch {}

                        try
                        {
                            if (includeSent)
                            {
                                dynamic f = root.Folders["Sent Items"];
                                if (f == null) f = store.GetDefaultFolder(5);
                                if (f != null) foldersToScan.Add((f, "Sent Items"));
                            }
                        } catch {}

                        try
                        {
                            if (includeDeleted)
                            {
                                dynamic f = root.Folders["Deleted Items"];
                                if (f == null) f = store.GetDefaultFolder(3);
                                if (f != null) foldersToScan.Add((f, "Deleted Items"));
                            }
                        } catch {}

                        Marshal.ReleaseComObject(root);

                        foreach (var item in foldersToScan)
                        {
                            if (cancellationToken.IsCancellationRequested) break;

                            dynamic folder = item.Folder;
                            string folderName = item.Name;

                            progress.Report($"Pass 1: {displayName} > {folderName}...");

                            string startS = startDate.ToString("g");
                            string endS = endDate.ToString("g");
                            string filter = $"[ReceivedTime] >= '{startS}' AND [ReceivedTime] <= '{endS}'";

                            dynamic items = folder.Items;
                            items.Sort("[ReceivedTime]", true);

                            dynamic restricted = null;
                            try { restricted = items.Restrict(filter); }
                            catch { restricted = items; }

                            dynamic mail = restricted.GetFirst();

                            while (mail != null)
                            {
                                if (cancellationToken.IsCancellationRequested) break;

                                try
                                {
                                    if ((int)mail.Class == 43) // olMail
                                    {
                                        bool isUnread = false;
                                        try { isUnread = mail.UnRead; } catch {}

                                        bool isInbox = folderName.Equals("Inbox", StringComparison.OrdinalIgnoreCase);
                                        if (isInbox && !isUnread && !includeRead)
                                        {
                                            // Skip read Inbox items unless user opts in to include them.
                                        }
                                        else
                                        {
                                            string entryId = ""; try { entryId = mail.EntryID; } catch {}
                                            DateTime received = DateTime.MinValue; try { received = mail.ReceivedTime; } catch {}
                                            string senderName = ""; try { senderName = mail.SenderName; } catch {}
                                            string senderEmail = ""; try { senderEmail = GetSenderAddress(mail); } catch {}
                                            string subject = ""; try { subject = mail.Subject; } catch {}
                                            string body = ""; try { body = mail.Body; } catch {}
                                            string categories = ""; try { categories = mail.Categories; } catch {}

                                            string rawHeaders = "";
                                            bool hasUnsub = false, isBulk = false;
                                            bool hasFeedbackId = false, isMailingList = false, isGoogleGroup = false;
                                            bool dmarcPass = false, spfPass = false, dkimPass = false;
                                            bool yahooBulk = false, yahooNewman = false, yahooAd = false;
                                            string yahooClassification = "";
                                            string focusedOrOther = "";

                                            try
                                            {
                                                dynamic pa = mail.PropertyAccessor;
                                                try { rawHeaders = pa.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS); } catch {}

                                                // Try MAPI property paths for Focused/Other
                                                foreach (var propPath in focusedPropertyPaths)
                                                {
                                                    try
                                                    {
                                                        object val = pa.GetProperty(propPath);
                                                        if (val != null)
                                                        {
                                                            if (val is int intVal)
                                                            {
                                                                focusedOrOther = (intVal & 0x00C00000) != 0 ? "Other" : "Focused";
                                                                break;
                                                            }

                                                            string strVal = val.ToString();
                                                            if (!string.IsNullOrWhiteSpace(strVal))
                                                            {
                                                                if (int.TryParse(strVal, out int parsedInt))
                                                                {
                                                                    focusedOrOther = (parsedInt & 0x00C00000) != 0 ? "Other" : "Focused";
                                                                    break;
                                                                }

                                                                if (strVal == "0") focusedOrOther = "Focused";
                                                                else if (strVal == "1") focusedOrOther = "Other";
                                                                else if (strVal == "2") focusedOrOther = "Focused";
                                                                else if (strVal.Equals("True", StringComparison.OrdinalIgnoreCase)) focusedOrOther = "Other";
                                                                else if (strVal.Equals("False", StringComparison.OrdinalIgnoreCase)) focusedOrOther = "Focused";
                                                                else if (strVal.Equals("focused", StringComparison.OrdinalIgnoreCase)) focusedOrOther = "Focused";
                                                                else if (strVal.Equals("other", StringComparison.OrdinalIgnoreCase)) focusedOrOther = "Other";
                                                                else if (strVal.Equals("clutter", StringComparison.OrdinalIgnoreCase)) focusedOrOther = "Other";
                                                                else continue; // unknown value, try next path

                                                                break;
                                                            }
                                                        }
                                                    }
                                                    catch { }
                                                }

                                                // Fallback: Exchange transport headers
                                                if (string.IsNullOrEmpty(focusedOrOther) && !string.IsNullOrEmpty(rawHeaders))
                                                {
                                                    if (rawHeaders.IndexOf("X-MS-Exchange-Organization-Clutter", StringComparison.OrdinalIgnoreCase) >= 0)
                                                    {
                                                        try
                                                        {
                                                            int cStart = rawHeaders.IndexOf("X-MS-Exchange-Organization-Clutter:", StringComparison.OrdinalIgnoreCase);
                                                            int cEnd = rawHeaders.IndexOf("\n", cStart + 1);
                                                            if (cEnd > cStart)
                                                            {
                                                                string cVal = rawHeaders.Substring(cStart + 35, Math.Min(cEnd - cStart - 35, 50)).Trim();
                                                                focusedOrOther = cVal.Equals("true", StringComparison.OrdinalIgnoreCase) ? "Other" : "Focused";
                                                            }
                                                        }
                                                        catch { }
                                                    }

                                                    if (string.IsNullOrEmpty(focusedOrOther) && rawHeaders.IndexOf("X-Forefront-Antispam-Report", StringComparison.OrdinalIgnoreCase) >= 0)
                                                    {
                                                        try
                                                        {
                                                            int fStart = rawHeaders.IndexOf("X-Forefront-Antispam-Report:", StringComparison.OrdinalIgnoreCase);
                                                            int fEnd = rawHeaders.IndexOf("\n", fStart + 1);
                                                            if (fEnd > fStart)
                                                            {
                                                                string fVal = rawHeaders.Substring(fStart + 28, Math.Min(fEnd - fStart - 28, 500)).Trim();
                                                                if (fVal.IndexOf("SFV:NSPM", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                                    fVal.IndexOf("SFV=NSPM", StringComparison.OrdinalIgnoreCase) >= 0)
                                                                    focusedOrOther = "Focused";
                                                                else if (fVal.IndexOf("SFV:SPM", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                                         fVal.IndexOf("SFV=SPM", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                                         fVal.IndexOf("SFV:SKB", StringComparison.OrdinalIgnoreCase) >= 0)
                                                                    focusedOrOther = "Other";
                                                            }
                                                        }
                                                        catch { }
                                                    }
                                                }

                                                // Extract header signals
                                                if (!string.IsNullOrEmpty(rawHeaders))
                                                {
                                                    hasUnsub = rawHeaders.IndexOf("List-Unsubscribe", StringComparison.OrdinalIgnoreCase) >= 0;
                                                    isBulk = rawHeaders.IndexOf("Precedence: bulk", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                             rawHeaders.IndexOf("X-Auto-Response-Suppress", StringComparison.OrdinalIgnoreCase) >= 0;
                                                    hasFeedbackId = rawHeaders.IndexOf("Feedback-ID", StringComparison.OrdinalIgnoreCase) >= 0;
                                                    isMailingList = rawHeaders.IndexOf("Mailing-List", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                                   rawHeaders.IndexOf("List-Id", StringComparison.OrdinalIgnoreCase) >= 0;
                                                    isGoogleGroup = rawHeaders.IndexOf("X-Google-Group-Id", StringComparison.OrdinalIgnoreCase) >= 0;

                                                    if (rawHeaders.IndexOf("Authentication-Results", StringComparison.OrdinalIgnoreCase) >= 0)
                                                    {
                                                        try
                                                        {
                                                            int authStart = rawHeaders.IndexOf("Authentication-Results:", StringComparison.OrdinalIgnoreCase);
                                                            int authEnd = rawHeaders.IndexOf("\n", authStart + 1);
                                                            if (authEnd > authStart)
                                                            {
                                                                string authLine = rawHeaders.Substring(authStart, Math.Min(authEnd - authStart, 500));
                                                                dmarcPass = authLine.IndexOf("dmarc=pass", StringComparison.OrdinalIgnoreCase) >= 0;
                                                                spfPass = authLine.IndexOf("spf=pass", StringComparison.OrdinalIgnoreCase) >= 0;
                                                                dkimPass = authLine.IndexOf("dkim=pass", StringComparison.OrdinalIgnoreCase) >= 0;
                                                            }
                                                        }
                                                        catch { }
                                                    }

                                                    yahooBulk = rawHeaders.IndexOf("X-YahooFilteredBulk", StringComparison.OrdinalIgnoreCase) >= 0;
                                                    yahooNewman = rawHeaders.IndexOf("X-Yahoo-Newman-Id", StringComparison.OrdinalIgnoreCase) >= 0;
                                                    yahooAd = rawHeaders.IndexOf("X-Rocket-MIMEInfo", StringComparison.OrdinalIgnoreCase) >= 0;

                                                    if (rawHeaders.IndexOf("X-Yahoo-Classification", StringComparison.OrdinalIgnoreCase) >= 0)
                                                    {
                                                        try
                                                        {
                                                            int ycStart = rawHeaders.IndexOf("X-Yahoo-Classification:", StringComparison.OrdinalIgnoreCase);
                                                            int ycEnd = rawHeaders.IndexOf("\n", ycStart + 1);
                                                            if (ycEnd > ycStart)
                                                                yahooClassification = rawHeaders.Substring(ycStart + 23, Math.Min(ycEnd - ycStart - 23, 100)).Trim();
                                                        }
                                                        catch { }
                                                    }
                                                }
                                                Marshal.ReleaseComObject(pa);
                                            }
                                            catch { /* No headers available */ }

                                            // Sanitize text fields
                                            string safeSub = (subject ?? "").Replace("\"", "").Replace(",", " ");
                                            string safeSend = (senderName ?? "").Replace("\"", "").Replace(",", " ");
                                            string safeAddr = (senderEmail ?? "").Replace("\"", "").Replace(",", " ");
                                            string safeBody = (body ?? "").Replace("\r", " ").Replace("\n", " ").Replace(",", " ").Replace("\"", "");
                                            if (safeBody.Length > 200) safeBody = safeBody.Substring(0, 200);
                                            string safeFocused = (focusedOrOther ?? "").Replace(",", " ");
                                            string safeYahooCls = (yahooClassification ?? "").Replace(",", " ");
                                            string safeCategories = (categories ?? "").Replace("\"", "").Replace(",", ";").Trim();

                                            // Store email data in memory
                                            var emailData = new Dictionary<string, object>
                                            {
                                                ["Account"] = displayName,
                                                ["AccountKey"] = accountKey,
                                                ["Folder"] = folderName,
                                                ["Categories"] = safeCategories,
                                                ["EntryID"] = entryId,
                                                ["ReceivedTime"] = received,
                                                ["SenderName"] = safeSend,
                                                ["SenderEmail"] = safeAddr,
                                                ["Subject"] = safeSub,
                                                ["BodySnippet"] = safeBody,
                                                ["IsRead"] = !isUnread,
                                                ["HasUnsubscribe"] = hasUnsub,
                                                ["IsBulk"] = isBulk,
                                                ["FocusedOrOther"] = safeFocused,
                                                ["HasFeedbackId"] = hasFeedbackId,
                                                ["IsMailingList"] = isMailingList,
                                                ["IsGoogleGroup"] = isGoogleGroup,
                                                ["DMARC_Pass"] = dmarcPass,
                                                ["SPF_Pass"] = spfPass,
                                                ["DKIM_Pass"] = dkimPass,
                                                ["YahooBulk"] = yahooBulk,
                                                ["YahooNewman"] = yahooNewman,
                                                ["YahooAd"] = yahooAd,
                                                ["YahooClassification"] = safeYahooCls
                                            };

                                            allEmails.Add(emailData);

                                            // Build sender stats
                                            string senderKey = safeAddr.ToLowerInvariant();
                                            if (!string.IsNullOrWhiteSpace(senderKey))
                                            {
                                                if (!senderStatsMap.ContainsKey(senderKey))
                                                {
                                                    senderStatsMap[senderKey] = new SenderStats
                                                    {
                                                        SenderEmail = safeAddr,
                                                        SenderName = safeSend
                                                    };
                                                }
                                                var stats = senderStatsMap[senderKey];
                                                stats.TotalEmails++;
                                                if (!isUnread) stats.ReadCount++; else stats.UnreadCount++;
                                                if (folderName.Equals("Deleted Items", StringComparison.OrdinalIgnoreCase)) stats.DeletedCount++;
                                                if (folderName.Equals("Sent Items", StringComparison.OrdinalIgnoreCase)) stats.RepliedToCount++;
                                                if (hasUnsub) stats.HasUnsubscribeCount++;
                                                if (isBulk) stats.IsBulkCount++;
                                                if (hasFeedbackId) stats.HasFeedbackIdCount++;
                                                if (safeFocused.Equals("Focused", StringComparison.OrdinalIgnoreCase)) stats.FocusedCount++;
                                                if (safeFocused.Equals("Other", StringComparison.OrdinalIgnoreCase)) stats.OtherCount++;
                                            }

                                            totalExtracted++;
                                            if (totalExtracted % 100 == 0)
                                                progress.Report($"Pass 1: Extracted {totalExtracted} emails...");
                                        }
                                    }
                                }
                                catch { /* Skip bad item */ }
                                finally
                                {
                                    dynamic next = restricted.GetNext();
                                    Marshal.ReleaseComObject(mail);
                                    mail = next;
                                }
                            }

                            if (restricted != null) Marshal.ReleaseComObject(restricted);
                            Marshal.ReleaseComObject(folder);
                        }

                        Marshal.ReleaseComObject(store);
                    }

                    if (stores != null) Marshal.ReleaseComObject(stores);

                    progress.Report($"Pass 1 complete: {allEmails.Count} emails extracted.");

                    // ══════════════════════════════════════════
                    // PASS 2: Score every email
                    // ══════════════════════════════════════════
                    progress.Report($"Pass 2: Scoring {allEmails.Count} emails...");

                    // Calculate sender reputations
                    var senderReputations = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    foreach (var kvp in senderStatsMap)
                    {
                        senderReputations[kvp.Key] = CalculateSenderReputation(kvp.Value);
                    }

                    // Score each email
                    foreach (var email in allEmails)
                    {
                        string senderKey = ((string)email["SenderEmail"]).ToLowerInvariant();
                        int senderRep = senderReputations.ContainsKey(senderKey) ? senderReputations[senderKey] : 50;
                        string accountKey = email.TryGetValue("AccountKey", out var accountKeyObj)
                            ? (string)(accountKeyObj ?? "")
                            : "";
                        learnedProfilesByAccountKey.TryGetValue(accountKey, out var learnedProfile);

                        int junkScore = CalculateJunkScore(
                            (bool)email["HasUnsubscribe"],
                            (bool)email["IsBulk"],
                            (string)email["FocusedOrOther"],
                            (bool)email["HasFeedbackId"],
                            (bool)email["IsMailingList"],
                            (bool)email["IsRead"],
                            (bool)email["DMARC_Pass"],
                            (bool)email["SPF_Pass"],
                            (bool)email["DKIM_Pass"],
                            (string)email["Folder"],
                            (DateTime)email["ReceivedTime"],
                            senderRep
                        );

                        junkScore = ApplyLearnedProfileAdjustments(
                            junkScore,
                            (string)email["SenderEmail"],
                            learnedProfile);

                        email["SenderReputation"] = senderRep;
                        email["JunkScore"] = junkScore;
                        email["Recommendation"] = GetRecommendation(junkScore);
                    }

                    progress.Report("Pass 2 complete: All emails scored.");

                    // ══════════════════════════════════════════
                    // PASS 3 (optional): Apply Outlook categories
                    // ══════════════════════════════════════════
                    if (applyOutlookCategories && ns != null)
                    {
                        progress.Report("Pass 3: Applying color categories in Outlook...");
                        ApplyCategoriesToOutlook(ns, allEmails, progress, cancellationToken);
                    }

                    // Release COM (ns stays alive until after categories are applied)
                    if (ns != null) Marshal.ReleaseComObject(ns);
                    if (app != null) Marshal.ReleaseComObject(app);

                    progress.Report("Extraction and scoring complete.");
                }
                catch (Exception ex)
                {
                    _log.Error(ex, "Dataset generation failed.");
                    staException = ex;
                }
                finally
                {
                    RetryMessageFilter.Revoke();
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (staException != null)
                throw new Exception($"Dataset extraction failed: {staException.Message}", staException);

            // ══════════════════════════════════════════
            // PASS 4: Write CSV + XLSX output
            // ══════════════════════════════════════════
            progress.Report("Writing output files...");

            try
            {
                // Write CSV
                var csvHeader = "Account,Folder,Categories,EntryID,ReceivedTime,SenderName,SenderEmail,Subject,BodySnippet,IsRead,HasUnsubscribe,IsBulk,FocusedOrOther,HasFeedbackId,IsMailingList,IsGoogleGroup,DMARC_Pass,SPF_Pass,DKIM_Pass,YahooBulk,YahooNewman,YahooAd,YahooClassification,SenderReputation,JunkScore,Recommendation";
                var csvLines = new List<string> { csvHeader };

                foreach (var email in allEmails)
                {
                    string safeRec = ((DateTime)email["ReceivedTime"]).ToString("yyyy-MM-dd HH:mm:ss");
                    string line = $"{email["Account"]},{email["Folder"]},{email.GetValueOrDefault("Categories", "")},{email["EntryID"]},{safeRec}," +
                                  $"{email["SenderName"]},{email["SenderEmail"]},{email["Subject"]},{email["BodySnippet"]}," +
                                  $"{email["IsRead"]},{email["HasUnsubscribe"]},{email["IsBulk"]},{email["FocusedOrOther"]}," +
                                  $"{email["HasFeedbackId"]},{email["IsMailingList"]},{email["IsGoogleGroup"]}," +
                                  $"{email["DMARC_Pass"]},{email["SPF_Pass"]},{email["DKIM_Pass"]}," +
                                  $"{email["YahooBulk"]},{email["YahooNewman"]},{email["YahooAd"]},{email["YahooClassification"]}," +
                                  $"{email["SenderReputation"]},{email["JunkScore"]},{email["Recommendation"]}";
                    csvLines.Add(line);
                }

                File.WriteAllLines(outputFilePath, csvLines, System.Text.Encoding.UTF8);

                // Write XLSX with 3 sheets (optional)
                if (saveXlsx)
                    WriteExcelWithScoring(outputFilePath, allEmails, senderStatsMap, progress);

                progress.Report($"Done: {allEmails.Count} emails scored and saved.");
            }
            catch (Exception ex)
            {
                _log.Error(ex, "Failed to write output files.");
                throw new Exception($"Failed to write output: {ex.Message}", ex);
            }
        });
    }

    // ═══════════════════════════════════════════════════════════════════
    // Score Existing Dataset (load CSV → validate → score → write)
    // ═══════════════════════════════════════════════════════════════════

    /// <summary>Required CSV columns for a valid MailZen dataset.</summary>
    private static readonly string[] RequiredCsvColumns = {
        "Account", "Folder", "EntryID", "ReceivedTime", "SenderName", "SenderEmail",
        "Subject", "BodySnippet", "IsRead", "HasUnsubscribe", "IsBulk", "FocusedOrOther",
        "HasFeedbackId", "IsMailingList", "IsGoogleGroup", "DMARC_Pass", "SPF_Pass", "DKIM_Pass",
        "YahooBulk", "YahooNewman", "YahooAd", "YahooClassification"
    };

    /// <summary>
    /// Validates a CSV file has the correct MailZen dataset format.
    /// Returns null if valid, or an error message if invalid.
    /// </summary>
    public static string? ValidateDatasetCsv(string filePath)
    {
        if (!File.Exists(filePath))
            return "File not found.";

        string[] lines;
        try { lines = File.ReadAllLines(filePath); }
        catch (Exception ex) { return $"Cannot read file: {ex.Message}"; }

        if (lines.Length < 2)
            return "File is empty or has no data rows (only a header).";

        var headerCols = lines[0].Split(',');
        var missing = RequiredCsvColumns.Where(c =>
            !headerCols.Any(h => h.Trim().Equals(c, StringComparison.OrdinalIgnoreCase))).ToList();

        if (missing.Count > 0)
            return $"Invalid dataset — missing columns: {string.Join(", ", missing)}";

        return null; // valid
    }

    /// <summary>
    /// Loads a previously-exported CSV dataset, scores every email,
    /// optionally applies Outlook categories, and writes scored CSV + XLSX.
    /// </summary>
    public async Task ScoreExistingDataset(
        string inputCsvPath,
        string outputDirectory,
        bool applyOutlookCategories,
        bool saveXlsx,
        IProgress<string> progress,
        CancellationToken cancellationToken)
    {
        await Task.Run(() =>
        {
            // ── Load and validate ─────────────────────────────────────
            progress.Report("Loading dataset...");
            var validationError = ValidateDatasetCsv(inputCsvPath);
            if (validationError != null)
                throw new Exception(validationError);

            var lines = File.ReadAllLines(inputCsvPath);
            var headerCols = lines[0].Split(',');

            // Build column index map
            var colIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headerCols.Length; i++)
                colIndex[headerCols[i].Trim()] = i;

            progress.Report($"Loaded {lines.Length - 1} emails from dataset.");

            // ── Parse rows into memory ─────────────────────────────────
            var allEmails = new List<Dictionary<string, object>>();
            var senderStatsMap = new Dictionary<string, SenderStats>(StringComparer.OrdinalIgnoreCase);
            int parseErrors = 0;

            for (int row = 1; row < lines.Length; row++)
            {
                if (cancellationToken.IsCancellationRequested) break;

                try
                {
                    var parts = lines[row].Split(',');
                    if (parts.Length < RequiredCsvColumns.Length) { parseErrors++; continue; }

                    string Col(string name) => colIndex.ContainsKey(name) && colIndex[name] < parts.Length
                        ? parts[colIndex[name]].Trim() : "";

                    bool ParseBool(string name)
                    {
                        var v = Col(name);
                        return v.Equals("True", StringComparison.OrdinalIgnoreCase) ||
                               v.Equals("TRUE", StringComparison.OrdinalIgnoreCase);
                    }

                    DateTime received = DateTime.TryParse(Col("ReceivedTime"), out var dt) ? dt : DateTime.MinValue;
                    bool isRead = ParseBool("IsRead");
                    string senderEmail = Col("SenderEmail");
                    string senderName = Col("SenderName");
                    string folder = Col("Folder");

                    var emailData = new Dictionary<string, object>
                    {
                        ["Account"] = Col("Account"),
                        ["Folder"] = folder,
                        ["Categories"] = Col("Categories"),
                        ["EntryID"] = Col("EntryID"),
                        ["ReceivedTime"] = received,
                        ["SenderName"] = senderName,
                        ["SenderEmail"] = senderEmail,
                        ["Subject"] = Col("Subject"),
                        ["BodySnippet"] = Col("BodySnippet"),
                        ["IsRead"] = isRead,
                        ["HasUnsubscribe"] = ParseBool("HasUnsubscribe"),
                        ["IsBulk"] = ParseBool("IsBulk"),
                        ["FocusedOrOther"] = Col("FocusedOrOther"),
                        ["HasFeedbackId"] = ParseBool("HasFeedbackId"),
                        ["IsMailingList"] = ParseBool("IsMailingList"),
                        ["IsGoogleGroup"] = ParseBool("IsGoogleGroup"),
                        ["DMARC_Pass"] = ParseBool("DMARC_Pass"),
                        ["SPF_Pass"] = ParseBool("SPF_Pass"),
                        ["DKIM_Pass"] = ParseBool("DKIM_Pass"),
                        ["YahooBulk"] = ParseBool("YahooBulk"),
                        ["YahooNewman"] = ParseBool("YahooNewman"),
                        ["YahooAd"] = ParseBool("YahooAd"),
                        ["YahooClassification"] = Col("YahooClassification")
                    };

                    allEmails.Add(emailData);

                    // Build sender stats
                    string senderKey = senderEmail.ToLowerInvariant();
                    if (!string.IsNullOrWhiteSpace(senderKey))
                    {
                        if (!senderStatsMap.ContainsKey(senderKey))
                            senderStatsMap[senderKey] = new SenderStats { SenderEmail = senderEmail, SenderName = senderName };

                        var stats = senderStatsMap[senderKey];
                        stats.TotalEmails++;
                        if (isRead) stats.ReadCount++; else stats.UnreadCount++;
                        if (folder.Equals("Deleted Items", StringComparison.OrdinalIgnoreCase)) stats.DeletedCount++;
                        if (folder.Equals("Sent Items", StringComparison.OrdinalIgnoreCase)) stats.RepliedToCount++;
                        if ((bool)emailData["HasUnsubscribe"]) stats.HasUnsubscribeCount++;
                        if ((bool)emailData["IsBulk"]) stats.IsBulkCount++;
                        if ((bool)emailData["HasFeedbackId"]) stats.HasFeedbackIdCount++;
                        string focused = Col("FocusedOrOther");
                        if (focused.Equals("Focused", StringComparison.OrdinalIgnoreCase)) stats.FocusedCount++;
                        if (focused.Equals("Other", StringComparison.OrdinalIgnoreCase)) stats.OtherCount++;
                    }
                }
                catch { parseErrors++; }

                if (row % 500 == 0)
                    progress.Report($"Parsing... {row}/{lines.Length - 1}");
            }

            if (allEmails.Count == 0)
                throw new Exception("No valid email rows found in the dataset.");

            progress.Report($"Parsed {allEmails.Count} emails ({parseErrors} skipped).");

            // ── Score ──────────────────────────────────────────────────
            progress.Report("Scoring emails...");

            var senderReputations = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in senderStatsMap)
                senderReputations[kvp.Key] = CalculateSenderReputation(kvp.Value);

            foreach (var email in allEmails)
            {
                string senderKey = ((string)email["SenderEmail"]).ToLowerInvariant();
                int senderRep = senderReputations.ContainsKey(senderKey) ? senderReputations[senderKey] : 50;

                int junkScore = CalculateJunkScore(
                    (bool)email["HasUnsubscribe"], (bool)email["IsBulk"],
                    (string)email["FocusedOrOther"], (bool)email["HasFeedbackId"],
                    (bool)email["IsMailingList"], (bool)email["IsRead"],
                    (bool)email["DMARC_Pass"], (bool)email["SPF_Pass"], (bool)email["DKIM_Pass"],
                    (string)email["Folder"], (DateTime)email["ReceivedTime"], senderRep);

                email["SenderReputation"] = senderRep;
                email["JunkScore"] = junkScore;
                email["Recommendation"] = GetRecommendation(junkScore);
            }

            progress.Report("Scoring complete.");

            // ── Apply Outlook categories (optional, needs COM) ─────────
            if (applyOutlookCategories)
            {
                Exception? staException = null;
                var thread = new Thread(() =>
                {
                    try
                    {
                        RetryMessageFilter.Register();
                        dynamic app = GetActiveComObject("Outlook.Application");
                        dynamic ns = app.GetNamespace("MAPI");
                        ns.Logon("", "", false, false);

                        ApplyCategoriesToOutlook(ns, allEmails, progress, cancellationToken);

                        Marshal.ReleaseComObject(ns);
                        Marshal.ReleaseComObject(app);
                    }
                    catch (Exception ex) { staException = ex; }
                    finally { RetryMessageFilter.Revoke(); }
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();

                if (staException != null)
                    progress.Report($"Warning: Could not apply Outlook categories — {staException.Message}");
            }

            // ── Write output ───────────────────────────────────────────
            progress.Report("Writing scored output...");

            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            string outputPath = Path.Combine(
                outputDirectory,
                Path.GetFileNameWithoutExtension(inputCsvPath) + "_scored.csv");

            var csvHeader = "Account,Folder,Categories,EntryID,ReceivedTime,SenderName,SenderEmail,Subject,BodySnippet,IsRead,HasUnsubscribe,IsBulk,FocusedOrOther,HasFeedbackId,IsMailingList,IsGoogleGroup,DMARC_Pass,SPF_Pass,DKIM_Pass,YahooBulk,YahooNewman,YahooAd,YahooClassification,SenderReputation,JunkScore,Recommendation";
            var csvLines = new List<string> { csvHeader };

            foreach (var email in allEmails)
            {
                string safeRec = ((DateTime)email["ReceivedTime"]).ToString("yyyy-MM-dd HH:mm:ss");
                string line = $"{email["Account"]},{email["Folder"]},{email.GetValueOrDefault("Categories", "")},{email["EntryID"]},{safeRec}," +
                              $"{email["SenderName"]},{email["SenderEmail"]},{email["Subject"]},{email["BodySnippet"]}," +
                              $"{email["IsRead"]},{email["HasUnsubscribe"]},{email["IsBulk"]},{email["FocusedOrOther"]}," +
                              $"{email["HasFeedbackId"]},{email["IsMailingList"]},{email["IsGoogleGroup"]}," +
                              $"{email["DMARC_Pass"]},{email["SPF_Pass"]},{email["DKIM_Pass"]}," +
                              $"{email["YahooBulk"]},{email["YahooNewman"]},{email["YahooAd"]},{email["YahooClassification"]}," +
                              $"{email["SenderReputation"]},{email["JunkScore"]},{email["Recommendation"]}";
                csvLines.Add(line);
            }

            File.WriteAllLines(outputPath, csvLines, System.Text.Encoding.UTF8);
            if (saveXlsx)
                WriteExcelWithScoring(outputPath, allEmails, senderStatsMap, progress);

            progress.Report($"Done: {allEmails.Count} emails scored → {Path.GetFileName(outputPath)}");
        });
    }

    /// <summary>
    /// Writes XLSX output with 3 sheets: Dataset, Sender Report, and Column Guide.
    /// </summary>
    private void WriteExcelWithScoring(
        string csvPath,
        List<Dictionary<string, object>> allEmails,
        Dictionary<string, SenderStats> senderStatsMap,
        IProgress<string> progress)
    {
        var xlsxPath = Path.ChangeExtension(csvPath, ".xlsx");
        progress.Report("Creating Excel file...");

        using var workbook = new XLWorkbook();

        // ═══════════════════════════════════════
        // Sheet 1: Dataset
        // ═══════════════════════════════════════
        var ws = workbook.Worksheets.Add("Dataset");
        string[] headers = { "Account", "Folder", "Categories", "EntryID", "ReceivedTime", "SenderName", "SenderEmail",
            "SenderReputation", "JunkScore", "Recommendation" };

        // Header row
        for (int c = 0; c < headers.Length; c++)
        {
            var cell = ws.Cell(1, c + 1);
            cell.Value = headers[c];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#4A148C");
            cell.Style.Font.FontColor = XLColor.White;
        }

        // Data rows
        for (int r = 0; r < allEmails.Count; r++)
        {
            var email = allEmails[r];
            int row = r + 2;

            ws.Cell(row, 1).Value = (string)email["Account"];
            ws.Cell(row, 2).Value = (string)email["Folder"];
            ws.Cell(row, 3).Value = (string)email["EntryID"];
            ws.Cell(row, 4).Value = (DateTime)email["ReceivedTime"];
            ws.Cell(row, 4).Style.NumberFormat.Format = "yyyy-MM-dd HH:mm";
            ws.Cell(row, 5).Value = (string)email["SenderName"];
            ws.Cell(row, 6).Value = (string)email["SenderEmail"];
            ws.Cell(row, 7).Value = (string)email["Subject"];
            ws.Cell(row, 8).Value = (string)email["BodySnippet"];
            ws.Cell(row, 9).Value = (bool)email["IsRead"];
            ws.Cell(row, 10).Value = (bool)email["HasUnsubscribe"];
            ws.Cell(row, 11).Value = (bool)email["IsBulk"];
            ws.Cell(row, 12).Value = (string)email["FocusedOrOther"];
            ws.Cell(row, 13).Value = (bool)email["HasFeedbackId"];
            ws.Cell(row, 14).Value = (bool)email["IsMailingList"];
            ws.Cell(row, 15).Value = (bool)email["IsGoogleGroup"];
            ws.Cell(row, 16).Value = (bool)email["DMARC_Pass"];
            ws.Cell(row, 17).Value = (bool)email["SPF_Pass"];
            ws.Cell(row, 18).Value = (bool)email["DKIM_Pass"];
            ws.Cell(row, 19).Value = (bool)email["YahooBulk"];
            ws.Cell(row, 20).Value = (bool)email["YahooNewman"];
            ws.Cell(row, 21).Value = (bool)email["YahooAd"];
            ws.Cell(row, 22).Value = (string)email["YahooClassification"];
            ws.Cell(row, 23).Value = (int)email["SenderReputation"];
            ws.Cell(row, 24).Value = (int)email["JunkScore"];
            ws.Cell(row, 25).Value = (string)email["Recommendation"];

            // Color-code JunkScore column
            int score = (int)email["JunkScore"];
            var scoreCell = ws.Cell(row, 24);
            if (score <= 30)
                scoreCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#C8E6C9"); // green
            else if (score <= 60)
                scoreCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF3E0"); // yellow
            else
                scoreCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFCDD2"); // red

            // Color-code Recommendation column
            string rec = (string)email["Recommendation"];
            var recCell = ws.Cell(row, 25);
            if (rec == "Keep")
                recCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#C8E6C9");
            else if (rec == "Review")
                recCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF3E0");
            else
                recCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFCDD2");
        }

        ws.Columns().AdjustToContents(1, 1);
        ws.SheetView.FreezeRows(1);

        // ═══════════════════════════════════════
        // Sheet 2: Sender Report
        // ═══════════════════════════════════════
        var sr = workbook.Worksheets.Add("Sender Report");
        string[] srHeaders = { "SenderEmail", "SenderName", "TotalEmails", "ReadCount",
            "UnreadCount", "DeletedCount", "RepliedTo", "HasUnsub%", "IsBulk%",
            "Focused%", "Other%", "Reputation", "Rating" };

        for (int c = 0; c < srHeaders.Length; c++)
        {
            var cell = sr.Cell(1, c + 1);
            cell.Value = srHeaders[c];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#4A148C");
            cell.Style.Font.FontColor = XLColor.White;
        }

        int srRow = 2;
        foreach (var kvp in senderStatsMap.OrderByDescending(x => x.Value.TotalEmails))
        {
            var s = kvp.Value;
            int rep = CalculateSenderReputation(s);
            string rating = rep <= 30 ? "Trusted" : rep <= 60 ? "Neutral" : "Suspicious";

            sr.Cell(srRow, 1).Value = s.SenderEmail;
            sr.Cell(srRow, 2).Value = s.SenderName;
            sr.Cell(srRow, 3).Value = s.TotalEmails;
            sr.Cell(srRow, 4).Value = s.ReadCount;
            sr.Cell(srRow, 5).Value = s.UnreadCount;
            sr.Cell(srRow, 6).Value = s.DeletedCount;
            sr.Cell(srRow, 7).Value = s.RepliedToCount;
            sr.Cell(srRow, 8).Value = s.TotalEmails > 0 ? Math.Round((double)s.HasUnsubscribeCount / s.TotalEmails * 100, 1) : 0;
            sr.Cell(srRow, 9).Value = s.TotalEmails > 0 ? Math.Round((double)s.IsBulkCount / s.TotalEmails * 100, 1) : 0;
            int classTotal = s.FocusedCount + s.OtherCount;
            sr.Cell(srRow, 10).Value = classTotal > 0 ? Math.Round((double)s.FocusedCount / classTotal * 100, 1) : 0;
            sr.Cell(srRow, 11).Value = classTotal > 0 ? Math.Round((double)s.OtherCount / classTotal * 100, 1) : 0;
            sr.Cell(srRow, 12).Value = rep;
            sr.Cell(srRow, 13).Value = rating;

            // Color-code rating
            var ratingCell = sr.Cell(srRow, 13);
            if (rating == "Trusted")
                ratingCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#C8E6C9");
            else if (rating == "Neutral")
                ratingCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF3E0");
            else
                ratingCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFCDD2");

            srRow++;
        }

        sr.Columns().AdjustToContents(1, 1);
        sr.SheetView.FreezeRows(1);

        // ═══════════════════════════════════════
        // Sheet 3: Column Guide
        // ═══════════════════════════════════════
        var info = workbook.Worksheets.Add("Column Guide");

        var colGuide = new (string Name, string Description, string Values, string Scoring)[]
        {
            ("Account", "The email address of the Outlook account that received this email.",
             "Email address string", "Used to segment analysis by account."),
            ("Folder", "The Outlook folder the email resides in.",
             "Inbox / Sent Items / Deleted Items",
             "+35 if Deleted Items, -50 if Sent Items."),
            ("EntryID", "Unique MAPI identifier for deduplication.",
             "Long hexadecimal string", "Not used in scoring."),
            ("ReceivedTime", "Date and time the email arrived.",
             "YYYY-MM-DD HH:MM:SS", "+10 if unread and older than 7 days."),
            ("SenderName", "Display name of the sender.",
             "Free text", "Used for Sender Reputation."),
            ("SenderEmail", "Actual email address of the sender.",
             "Email address", "Core grouping key for Sender Reputation."),
            ("Subject", "Subject line of the email.",
             "Free text (commas replaced)", "Not directly scored (future NLP)."),
            ("BodySnippet", "First 200 characters of the email body.",
             "Free text (up to 200 chars)", "Not directly scored (future NLP)."),
            ("IsRead", "Whether the email was opened in Outlook.",
             "TRUE / FALSE", "-10 if True (engagement signal)."),
            ("HasUnsubscribe", "Email has a List-Unsubscribe header.",
             "TRUE / FALSE", "+30 — strong newsletter/marketing signal."),
            ("IsBulk", "Server flagged as bulk/mass mail.",
             "TRUE / FALSE", "+25 — confirmed mass email."),
            ("FocusedOrOther", "Microsoft Exchange AI classification.",
             "Focused / Other / blank", "+20 if Other — Microsoft AI thinks it's low priority."),
            ("HasFeedbackId", "Google marketing campaign tracker present.",
             "TRUE / FALSE", "+15 — confirmed marketing campaign."),
            ("IsMailingList", "Email from a mailing list system.",
             "TRUE / FALSE", "+15 — automated list, not personal."),
            ("IsGoogleGroup", "Email from a Google Group.",
             "TRUE / FALSE", "Informational — group discussion email."),
            ("DMARC_Pass", "Sender domain passed DMARC authentication.",
             "TRUE / FALSE", "+10 if False — sender may be spoofed."),
            ("SPF_Pass", "Sending server IP verified by sender domain.",
             "TRUE / FALSE", "+5 if False — IP not authorized."),
            ("DKIM_Pass", "Email digital signature verified.",
             "TRUE / FALSE", "+5 if False — signature invalid."),
            ("YahooBulk", "Yahoo flagged this as bulk mail.",
             "TRUE / FALSE (blank for non-Yahoo)", "Yahoo's bulk filter signal."),
            ("YahooNewman", "Yahoo system notification email.",
             "TRUE / FALSE (blank for non-Yahoo)", "Automated Yahoo alert."),
            ("YahooAd", "Yahoo advertising network email.",
             "TRUE / FALSE (blank for non-Yahoo)", "Paid promotional email."),
            ("YahooClassification", "Yahoo's internal category label.",
             "String (blank for non-Yahoo)", "Yahoo's own email classification."),
            ("SenderReputation", "Aggregated score for this sender across ALL their emails. Based on delete ratio, read ratio, unsubscribe presence, and whether you've sent emails to them.",
             "0 = Trusted, 50 = Neutral, 100 = Suspicious",
             "Blended into JunkScore. High rep adds +15, low rep subtracts -15."),
            ("JunkScore", "Final composite unwanted-email score combining all signals plus sender reputation.",
             "0-100 integer",
             "0-30 = Keep, 31-60 = Review, 61-100 = Delete."),
            ("Recommendation", "Action recommendation based on JunkScore thresholds.",
             "Keep / Review / Delete",
             "Green = Keep (score 0-30), Yellow = Review (31-60), Red = Delete (61-100)."),
        };

        string[] infoHeaders = { "Column Name", "Description", "Values / Format", "Scoring Impact" };
        int[] infoColWidths = { 22, 70, 45, 55 };

        for (int c = 0; c < infoHeaders.Length; c++)
        {
            var hCell = info.Cell(1, c + 1);
            hCell.Value = infoHeaders[c];
            hCell.Style.Font.Bold = true;
            hCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#4A148C");
            hCell.Style.Font.FontColor = XLColor.White;
            hCell.Style.Font.FontSize = 11;
            info.Column(c + 1).Width = infoColWidths[c];
        }

        for (int r = 0; r < colGuide.Length; r++)
        {
            var (colName, desc, values, scoring) = colGuide[r];
            int rowNum = r + 2;
            bool isAlt = r % 2 == 0;
            var bgColor = isAlt ? XLColor.FromHtml("#EDE7F6") : XLColor.White;

            var nameCell = info.Cell(rowNum, 1);
            nameCell.Value = colName;
            nameCell.Style.Font.Bold = true;
            nameCell.Style.Fill.BackgroundColor = bgColor;
            nameCell.Style.Alignment.WrapText = true;
            nameCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

            foreach (var (col, text) in new[] { (2, desc), (3, values), (4, scoring) })
            {
                var cell = info.Cell(rowNum, col);
                cell.Value = text;
                cell.Style.Alignment.WrapText = true;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                cell.Style.Fill.BackgroundColor = bgColor;
            }
        }

        info.SheetView.FreezeRows(1);
        info.Row(1).Height = 20;

        workbook.SaveAs(xlsxPath);
        progress.Report("Excel file created.");
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

    // ═══════════════════════════════════════════════════════════════════
    // Color Coding Test (works for any account type)
    // ═══════════════════════════════════════════════════════════════════

    /// <summary>
    /// Fetches the first 3 newest Inbox emails from each specified store,
    /// assigning Delete / Review / Keep categories for preview display.
    /// </summary>
    public async Task<List<ColorTestEmailItem>> FetchEmailPreview(
        List<string> storeIds,
        IProgress<string> progress,
        CancellationToken cancellationToken)
    {
        var result = new List<ColorTestEmailItem>();

        await Task.Run(() =>
        {
            Exception? staException = null;

            var thread = new Thread(() =>
            {
                dynamic? app = null;
                dynamic? ns = null;

                try
                {
                    RetryMessageFilter.Register();

                    app = GetActiveComObject("Outlook.Application");
                    ns = app.GetNamespace("MAPI");
                    ns.Logon("", "", false, false);

                    dynamic stores = ns.Stores;
                    int storeCount = stores.Count;
                    string[] categories = { "MailZen: Delete", "MailZen: Review", "MailZen: Keep" };
                    string[] colors = { "#D32F2F", "#E65100", "#2E7D32" };

                    for (int s = 1; s <= storeCount; s++)
                    {
                        if (cancellationToken.IsCancellationRequested) break;

                        dynamic store = stores[s];
                        try
                        {
                            string sid = store.StoreID ?? "";
                            if (!storeIds.Contains(sid)) continue;

                            string displayName = store.DisplayName ?? "";
                            progress.Report($"Loading preview from {displayName}...");

                            dynamic inbox;
                            try { inbox = store.GetDefaultFolder(6); } // olFolderInbox
                            catch { inbox = store.GetRootFolder().Folders["Inbox"]; }

                            dynamic items = inbox.Items;
                            items.Sort("[ReceivedTime]", true); // newest first
                            int count = Math.Min(3, items.Count);

                            for (int i = 1; i <= count; i++)
                            {
                                dynamic mail = items[i];
                                try
                                {
                                    string sender = "";
                                    try { sender = mail.SenderName ?? ""; } catch { }
                                    if (string.IsNullOrEmpty(sender))
                                        try { sender = mail.SenderEmailAddress ?? ""; } catch { }

                                    string subject = "";
                                    try { subject = mail.Subject ?? ""; } catch { }
                                    if (subject.Length > 50) subject = subject.Substring(0, 50) + "...";

                                    string date = "";
                                    try { date = ((DateTime)mail.ReceivedTime).ToString("MMM d, yyyy"); } catch { }

                                    string entryId = "";
                                    try { entryId = mail.EntryID; } catch { }

                                    int idx = i - 1;
                                    result.Add(new ColorTestEmailItem
                                    {
                                        AccountName = displayName,
                                        StoreId = sid,
                                        EntryId = entryId,
                                        Sender = sender,
                                        SubjectSnippet = subject,
                                        ReceivedDate = date,
                                        Category = categories[idx],
                                        DisplayColor = colors[idx]
                                    });
                                }
                                finally { Marshal.ReleaseComObject(mail); }
                            }

                            Marshal.ReleaseComObject(items);
                            Marshal.ReleaseComObject(inbox);
                        }
                        finally { Marshal.ReleaseComObject(store); }
                    }

                    Marshal.ReleaseComObject(stores);
                    Marshal.ReleaseComObject(ns);
                    Marshal.ReleaseComObject(app);
                }
                catch (Exception ex) { staException = ex; }
                finally { RetryMessageFilter.Revoke(); }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (staException != null)
                throw new Exception($"Preview fetch failed: {staException.Message}", staException);
        });

        progress.Report($"Loaded {result.Count} emails for preview.");
        return result;
    }

    /// <summary>
    /// Applies MailZen categories to the specified emails and sets up
    /// AutoFormatRules (conditional formatting) on each account's Inbox view
    /// so rows display in colored text. Works for any account type.
    /// </summary>
    public async Task ApplyColorCodingTest(
        List<ColorTestEmailItem> emails,
        IProgress<string> progress,
        CancellationToken cancellationToken)
    {
        await Task.Run(() =>
        {
            Exception? staException = null;

            var thread = new Thread(() =>
            {
                dynamic? app = null;
                dynamic? ns = null;

                try
                {
                    RetryMessageFilter.Register();

                    app = GetActiveComObject("Outlook.Application");
                    ns = app.GetNamespace("MAPI");
                    ns.Logon("", "", false, false);

                    // ── Ensure MailZen categories exist ──
                    EnsureCategoriesExist(ns);

                    // ── Group emails by store and assign categories ──
                    var byStore = emails.GroupBy(e => e.StoreId).ToList();
                    dynamic stores = ns.Stores;
                    int storeCount = stores.Count;
                    int tagged = 0;

                    for (int s = 1; s <= storeCount; s++)
                    {
                        if (cancellationToken.IsCancellationRequested) break;

                        dynamic store = stores[s];
                        try
                        {
                            string sid = store.StoreID ?? "";
                            var group = byStore.FirstOrDefault(g => g.Key == sid);
                            if (group == null) continue;

                            string displayName = store.DisplayName ?? "";
                            progress.Report($"Tagging emails in {displayName}...");

                            dynamic inbox;
                            try { inbox = store.GetDefaultFolder(6); }
                            catch { inbox = store.GetRootFolder().Folders["Inbox"]; }

                            // Build a lookup of EntryIDs we need to tag
                            var entryIdToCategory = group.ToDictionary(e => e.EntryId, e => e);
                            var remaining = new HashSet<string>(entryIdToCategory.Keys);

                            dynamic items = inbox.Items;
                            items.Sort("[ReceivedTime]", true); // newest first
                            int scanLimit = Math.Min(20, items.Count); // scan a few more than 3 in case of ordering

                            for (int i = 1; i <= scanLimit && remaining.Count > 0; i++)
                            {
                                if (cancellationToken.IsCancellationRequested) break;

                                dynamic mail = items[i];
                                try
                                {
                                    string eid = mail.EntryID ?? "";
                                    if (!remaining.Contains(eid)) continue;

                                    var emailInfo = entryIdToCategory[eid];

                                    // Remove existing MailZen categories, keep user categories
                                    string existing = mail.Categories ?? "";
                                    var cats = existing.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                        .Select(c => c.Trim())
                                        .Where(c => !c.StartsWith("MailZen:"))
                                        .ToList();

                                    cats.Add(emailInfo.Category);
                                    mail.Categories = string.Join(", ", cats);
                                    mail.Save();
                                    tagged++;
                                    remaining.Remove(eid);

                                    progress.Report($"Tagged {tagged}/{emails.Count}: [{emailInfo.Category}] {emailInfo.SubjectSnippet}");
                                }
                                finally { Marshal.ReleaseComObject(mail); }
                            }

                            Marshal.ReleaseComObject(items);

                            // ── Set up AutoFormatRules on this Inbox view ──
                            try
                            {
                                progress.Report($"Setting up conditional formatting for {displayName}...");

                                dynamic currentView = inbox.CurrentView;
                                int viewType = currentView.ViewType;

                                if (viewType != 0) // 0 = olTableView
                                {
                                    progress.Report($"⚠ {displayName}: view is not Table type — skipping conditional formatting.");
                                    Marshal.ReleaseComObject(currentView);
                                }
                                else
                                {
                                    dynamic rules = currentView.AutoFormatRules;

                                    // Remove existing MailZen rules to avoid duplicates
                                    for (int r = rules.Count; r >= 1; r--)
                                    {
                                        try
                                        {
                                            dynamic rule = rules[r];
                                            string ruleName = rule.Name ?? "";
                                            if (ruleName.StartsWith("MailZen:"))
                                                rules.Remove(r);
                                        }
                                        catch { }
                                    }

                                    string daslProp = "urn:schemas-microsoft-com:office:office#Keywords";

                                    // Delete → Red bold
                                    dynamic rd = rules.Add("MailZen: Delete");
                                    rd.Filter = $"@SQL=\"{daslProp}\" ci_phrasematch 'MailZen: Delete'";
                                    rd.Font.Color = 255;     // OLE Red
                                    rd.Font.Bold = true;
                                    rd.Enabled = true;

                                    // Review → Orange bold
                                    dynamic rr = rules.Add("MailZen: Review");
                                    rr.Filter = $"@SQL=\"{daslProp}\" ci_phrasematch 'MailZen: Review'";
                                    rr.Font.Color = 33023;   // OLE Orange
                                    rr.Font.Bold = true;
                                    rr.Enabled = true;

                                    // Keep → Green bold
                                    dynamic rk = rules.Add("MailZen: Keep");
                                    rk.Filter = $"@SQL=\"{daslProp}\" ci_phrasematch 'MailZen: Keep'";
                                    rk.Font.Color = 32768;   // OLE Green
                                    rk.Font.Bold = true;
                                    rk.Enabled = true;

                                    currentView.Save();
                                    currentView.Apply();
                                    progress.Report($"✅ Conditional formatting applied to {displayName} Inbox.");
                                    Marshal.ReleaseComObject(currentView);
                                }
                            }
                            catch (Exception cfEx)
                            {
                                progress.Report($"⚠ {displayName}: conditional formatting not supported — categories still applied.");
                                _log.Warn("AutoFormatRules failed for {Store}: {Error}", displayName, cfEx.Message);
                            }

                            Marshal.ReleaseComObject(inbox);
                        }
                        finally { Marshal.ReleaseComObject(store); }
                    }

                    Marshal.ReleaseComObject(stores);
                    Marshal.ReleaseComObject(ns);
                    Marshal.ReleaseComObject(app);
                }
                catch (Exception ex) { staException = ex; }
                finally { RetryMessageFilter.Revoke(); }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (staException != null)
                throw new Exception($"Color coding test failed: {staException.Message}", staException);

            progress.Report("✅ Test complete! Check your Inbox in Outlook for colored rows.");
        });
    }

    /// <summary>
    /// Removes MailZen categories and AutoFormatRules from the test emails.
    /// </summary>
    public async Task CleanUpColorTest(
        List<ColorTestEmailItem> emails,
        IProgress<string> progress,
        CancellationToken cancellationToken)
    {
        await Task.Run(() =>
        {
            Exception? staException = null;

            var thread = new Thread(() =>
            {
                dynamic? app = null;
                dynamic? ns = null;

                try
                {
                    RetryMessageFilter.Register();

                    app = GetActiveComObject("Outlook.Application");
                    ns = app.GetNamespace("MAPI");
                    ns.Logon("", "", false, false);

                    var byStore = emails.GroupBy(e => e.StoreId).ToList();
                    dynamic stores = ns.Stores;
                    int storeCount = stores.Count;
                    int cleaned = 0;

                    for (int s = 1; s <= storeCount; s++)
                    {
                        if (cancellationToken.IsCancellationRequested) break;

                        dynamic store = stores[s];
                        try
                        {
                            string sid = store.StoreID ?? "";
                            var group = byStore.FirstOrDefault(g => g.Key == sid);
                            if (group == null) continue;

                            string displayName = store.DisplayName ?? "";
                            progress.Report($"Cleaning categories from {displayName}...");

                            dynamic inbox;
                            try { inbox = store.GetDefaultFolder(6); }
                            catch { inbox = store.GetRootFolder().Folders["Inbox"]; }

                            var entryIds = new HashSet<string>(group.Select(e => e.EntryId));
                            var remaining = new HashSet<string>(entryIds);

                            dynamic items = inbox.Items;
                            items.Sort("[ReceivedTime]", true);
                            int scanLimit = Math.Min(20, items.Count);

                            for (int i = 1; i <= scanLimit && remaining.Count > 0; i++)
                            {
                                if (cancellationToken.IsCancellationRequested) break;

                                dynamic mail = items[i];
                                try
                                {
                                    string eid = mail.EntryID ?? "";
                                    if (!remaining.Contains(eid)) continue;

                                    string existing = mail.Categories ?? "";
                                    var cats = existing.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                        .Select(c => c.Trim())
                                        .Where(c => !c.StartsWith("MailZen:"))
                                        .ToList();

                                    mail.Categories = cats.Count > 0 ? string.Join(", ", cats) : "";
                                    mail.Save();
                                    cleaned++;
                                    remaining.Remove(eid);

                                    progress.Report($"Cleaned {cleaned}/{emails.Count}");
                                }
                                finally { Marshal.ReleaseComObject(mail); }
                            }

                            Marshal.ReleaseComObject(items);

                            // Remove AutoFormatRules
                            try
                            {
                                dynamic currentView = inbox.CurrentView;
                                if ((int)currentView.ViewType == 0)
                                {
                                    dynamic rules = currentView.AutoFormatRules;
                                    for (int r = rules.Count; r >= 1; r--)
                                    {
                                        try
                                        {
                                            dynamic rule = rules[r];
                                            string ruleName = rule.Name ?? "";
                                            if (ruleName.StartsWith("MailZen:"))
                                                rules.Remove(r);
                                        }
                                        catch { }
                                    }
                                    currentView.Save();
                                    currentView.Apply();
                                }
                                Marshal.ReleaseComObject(currentView);
                            }
                            catch { }

                            Marshal.ReleaseComObject(inbox);
                        }
                        finally { Marshal.ReleaseComObject(store); }
                    }

                    Marshal.ReleaseComObject(stores);
                    Marshal.ReleaseComObject(ns);
                    Marshal.ReleaseComObject(app);
                }
                catch (Exception ex) { staException = ex; }
                finally { RetryMessageFilter.Revoke(); }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (staException != null)
                throw new Exception($"Cleanup failed: {staException.Message}", staException);

            progress.Report("✅ Categories and formatting removed.");
        });
    }

    /// <summary>
    /// Global production cleanup: removes MailZen: category tags from ALL emails in
    /// the Inbox of each selected store. Optionally removes AutoFormatRules.
    /// </summary>
    public async Task ClearMailZenCategoriesFromInbox(
        List<string> storeIds,
        bool alsoRemoveFormatting,
        IProgress<string> progress,
        CancellationToken cancellationToken)
    {
        await Task.Run(() =>
        {
            Exception? staException = null;

            var thread = new Thread(() =>
            {
                dynamic? app = null;
                dynamic? ns = null;

                try
                {
                    RetryMessageFilter.Register();

                    app = GetActiveComObject("Outlook.Application");
                    ns = app.GetNamespace("MAPI");
                    ns.Logon("", "", false, false);

                    var targetIds = new HashSet<string>(storeIds, StringComparer.OrdinalIgnoreCase);
                    dynamic stores = ns.Stores;
                    int storeCount = stores.Count;
                    int totalScanned = 0;
                    int totalCleaned = 0;

                    for (int s = 1; s <= storeCount; s++)
                    {
                        if (cancellationToken.IsCancellationRequested) break;

                        dynamic store = stores[s];
                        try
                        {
                            string sid = store.StoreID ?? "";
                            if (!targetIds.Contains(sid)) continue;

                            string displayName = store.DisplayName ?? "";
                            progress.Report($"Scanning {displayName}...");

                            dynamic inbox;
                            try { inbox = store.GetDefaultFolder(6); }
                            catch { inbox = store.GetRootFolder().Folders["Inbox"]; }

                            dynamic items = inbox.Items;
                            int count = items.Count;
                            int storeScanned = 0;
                            int storeCleaned = 0;

                            for (int i = 1; i <= count; i++)
                            {
                                if (cancellationToken.IsCancellationRequested) break;

                                dynamic mail = items[i];
                                try
                                {
                                    string existing = mail.Categories ?? "";
                                    if (existing.Contains("MailZen:"))
                                    {
                                        var cats = existing.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                            .Select(c => c.Trim())
                                            .Where(c => !c.StartsWith("MailZen:"))
                                            .ToList();
                                        mail.Categories = cats.Count > 0 ? string.Join(", ", cats) : "";
                                        mail.Save();
                                        storeCleaned++;
                                    }
                                    storeScanned++;
                                }
                                catch { }
                                finally { Marshal.ReleaseComObject(mail); }

                                if (i % 100 == 0)
                                    progress.Report($"{displayName}: {i}/{count} scanned, {storeCleaned} cleared...");
                            }

                            totalScanned += storeScanned;
                            totalCleaned += storeCleaned;

                            Marshal.ReleaseComObject(items);

                            if (alsoRemoveFormatting)
                            {
                                try
                                {
                                    dynamic currentView = inbox.CurrentView;
                                    if ((int)currentView.ViewType == 0)
                                    {
                                        dynamic rules = currentView.AutoFormatRules;
                                        for (int r = rules.Count; r >= 1; r--)
                                        {
                                            try
                                            {
                                                dynamic rule = rules[r];
                                                string ruleName = rule.Name ?? "";
                                                if (ruleName.StartsWith("MailZen:"))
                                                    rules.Remove(r);
                                            }
                                            catch { }
                                        }
                                        currentView.Save();
                                        currentView.Apply();
                                    }
                                    Marshal.ReleaseComObject(currentView);
                                }
                                catch { /* AutoFormatRules not supported on this store type */ }
                            }

                            Marshal.ReleaseComObject(inbox);
                        }
                        finally { Marshal.ReleaseComObject(store); }
                    }

                    Marshal.ReleaseComObject(stores);
                    Marshal.ReleaseComObject(ns);
                    Marshal.ReleaseComObject(app);

                    progress.Report($"✅ Done — Scanned {totalScanned} · Cleared {totalCleaned}");
                }
                catch (Exception ex) { staException = ex; }
                finally { RetryMessageFilter.Revoke(); }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (staException != null)
                throw new Exception($"Inbox cleanup failed: {staException.Message}", staException);
        });
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
