using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using EmailManage.Models;
using EmailManage.Services;

namespace EmailManage.ViewModels;

/// <summary>
/// ViewModel for the MailZen linear workflow.
/// Steps: Connect → Learn → Triage → Review → Automate
/// </summary>
public partial class MainViewModel : ObservableObject
{
    private readonly OutlookConnectorService _connector;
    private readonly DiagnosticLogger _log;
    private LearningService? _learningService;
    private TriageService? _triageService;
    private OllamaClient? _ollamaClient;
    private LearnedProfile? _learnedProfile;

    // ── Step indicator colors ──
    private const string ColorActive  = "#6C63FF"; // Indigo — current step
    private const string ColorDone    = "#00C853"; // Green  — completed step
    private const string ColorPending = "#D1D5DB"; // Gray   — not yet reached

    // ═════════════════════════════════════════════
    //   OBSERVABLE PROPERTIES
    // ═════════════════════════════════════════════

    // ── General ──
    [ObservableProperty] private string _appTitle = "✨ MailZen";
    [ObservableProperty] private string _statusBarText = "Starting up...";
    [ObservableProperty] private string _statusMessage = "";

    // ── Connection ──
    [ObservableProperty] private string _connectionStatus = "Disconnected";
    [ObservableProperty] private string _connectionStatusColor = "#EF5350";
    [ObservableProperty] private bool   _isConnecting;
    [ObservableProperty] private bool   _isConnected;
    [ObservableProperty] private bool   _hasError;
    [ObservableProperty] private string? _errorMessage;
    [ObservableProperty] private OutlookAccountInfo? _selectedAccount;

    // ── Workflow step ──
    [ObservableProperty] private string _currentStep = "Connect";

    // ── Step indicator colors (bound to the horizontal pipeline) ──
    [ObservableProperty] private string _step1Color = ColorActive;
    [ObservableProperty] private string _step2Color = ColorPending;
    [ObservableProperty] private string _step3Color = ColorPending;
    [ObservableProperty] private string _step4Color = ColorPending;
    [ObservableProperty] private string _step5Color = ColorPending;

    // ── Settings flyout ──
    [ObservableProperty] private bool _showSettings;

    // ── Step 2: Learn ──
    [ObservableProperty] private bool _isLearning;
    [ObservableProperty] private bool _showAccountPicker;
    [ObservableProperty] private bool _showStartLearning;

    // ── Step 3: Triage ──
    [ObservableProperty] private bool _isTriaging;

    // ── Step 4: Review ──
    [ObservableProperty] private int  _reviewCount;
    [ObservableProperty] private bool _isProcessing;

    // ── Step 5: Automate ──
    [ObservableProperty] private int _deletedCount;
    [ObservableProperty] private int _keptCount;
    [ObservableProperty] private int _rulesCreatedCount;

    // ── AI Engine ──
    [ObservableProperty] private string _aiStatus = "Not checked";
    [ObservableProperty] private string _aiStatusColor = "#9E9E9E";

    [ObservableProperty] private bool   _isAiReady;
    [ObservableProperty] private string? _aiModelName;
    [ObservableProperty] private string? _aiModelSize;
    [ObservableProperty] private string? _aiOllamaVersion;
    [ObservableProperty] private bool   _isInstallingOllama;

    // ── Settings: Profile stats ──
    [ObservableProperty] private int  _protectedSenderCount;
    [ObservableProperty] private int  _ruledSenderCount;
    [ObservableProperty] private int  _learnedDomainCount;
    [ObservableProperty] private string? _profileLastLearnedAt;

    private OllamaSetupService? _ollamaSetup;
    private List<FolderEmailSnapshot>? _reviewSnapshot;
    private CancellationTokenSource? _cts;

    // ── Stop / Cancel ──
    [ObservableProperty] private bool _canStop;
    [ObservableProperty] private bool _wasStopped;

    public ObservableCollection<OutlookAccountInfo> Accounts { get; } = [];

    // ═════════════════════════════════════════════
    //   CONSTRUCTOR
    // ═════════════════════════════════════════════

    public MainViewModel()
    {
        _connector = new OutlookConnectorService();
        _log = DiagnosticLogger.Instance;
    }

    // ═════════════════════════════════════════════
    //   STEP NAVIGATION HELPERS
    // ═════════════════════════════════════════════

    private void GoToStep(string step)
    {
        CurrentStep = step;
        RefreshStepColors();
        _log.Info("Workflow → Step: {Step}", step);
    }

    private void RefreshStepColors()
    {
        var steps = new[] { "Connect", "Learn", "Triage", "Review", "Automate" };
        var current = Array.IndexOf(steps, CurrentStep);

        Step1Color = current > 0 ? ColorDone : current == 0 ? ColorActive : ColorPending;
        Step2Color = current > 1 ? ColorDone : current == 1 ? ColorActive : ColorPending;
        Step3Color = current > 2 ? ColorDone : current == 2 ? ColorActive : ColorPending;
        Step4Color = current > 3 ? ColorDone : current == 3 ? ColorActive : ColorPending;
        Step5Color = current >= 4 ? ColorActive : ColorPending;
    }

    // ═════════════════════════════════════════════
    //   STEP 1: CONNECT (auto-runs on startup)
    // ═════════════════════════════════════════════

    /// <summary>Called from Window_Loaded via MainWindow.xaml.cs.</summary>
    public async Task InitializeAsync()
    {
        _log.Info("MailZen starting — auto-connecting...");
        GoToStep("Connect");

        // Run Outlook connection and AI setup in parallel
        var connectTask = ConnectAsync();
        var aiTask = SetupAiSilentAsync();

        await connectTask;
        await aiTask;
    }

    private async Task ConnectAsync()
    {
        IsConnecting = true;
        HasError = false;
        ErrorMessage = null;
        ConnectionStatus = "Connecting...";
        ConnectionStatusColor = "#2196F3";
        StatusMessage = "Looking for Outlook Desktop...";
        StatusBarText = "Connecting to Outlook...";
        Accounts.Clear();
        SelectedAccount = null;

        _log.Info("Auto-connect to Outlook started.");

        var result = await _connector.ConnectAsync();

        IsConnecting = false;

        if (result.Success)
        {
            IsConnected = true;
            ConnectionStatus = "Connected";
            ConnectionStatusColor = "#00C853";
            foreach (var account in result.Accounts)
                Accounts.Add(account);

            StatusBarText = $"Connected. {Accounts.Count} account(s) found.";
            StatusMessage = "";
            _log.Info("Connected. {Count} accounts found.", Accounts.Count);

            // Auto-select first if only one account
            if (Accounts.Count == 1)
                SelectedAccount = Accounts[0];

            // Advance to Learn step
            ShowAccountPicker = Accounts.Count > 1;
            ShowStartLearning = Accounts.Count == 1;
            GoToStep("Learn");
        }
        else
        {
            IsConnected = false;
            HasError = true;
            ErrorMessage = result.ErrorMessage;
            ConnectionStatus = "Error";
            ConnectionStatusColor = "#F44336";
            StatusBarText = $"Connection failed: {result.ErrorCode}";
            StatusMessage = "";
        }
    }

    [RelayCommand]
    private async Task RetryConnectAsync()
    {
        await ConnectAsync();
    }

    // ═════════════════════════════════════════════
    //   STEP 2: LEARN (AI scans Deleted Items)
    // ═════════════════════════════════════════════

    partial void OnSelectedAccountChanged(OutlookAccountInfo? value)
    {
        if (value is not null)
        {
            _log.Info("Account selected: {Email}", value.EmailAddress);
            StatusBarText = $"Active: {value.EmailAddress}";
            ShowStartLearning = true;
        }
        else
        {
            ShowStartLearning = false;
        }
    }

    [RelayCommand]
    private async Task StartLearningAsync()
    {
        if (SelectedAccount == null) return;

        IsLearning = true;
        ShowStartLearning = false;
        ShowAccountPicker = false;
        WasStopped = false;
        StatusMessage = "Scanning your Deleted Items to learn what you delete...";

        // Create a fresh cancellation token
        _cts?.Dispose();
        _cts = new CancellationTokenSource();
        CanStop = true;

        try
        {
            _learningService ??= new LearningService(_connector);
            var progress = new Progress<string>(msg => StatusMessage = msg);

            _learnedProfile = await _learningService.LearnFromDeletedItemsAsync(
                SelectedAccount, progress, _cts.Token);

            StatusMessage = $"Learned from {_learnedProfile.TotalDeletedScanned} deleted emails. " +
                            $"{_learnedProfile.DeletedDomainCounts.Count} unique domains identified.";
            _log.Info("Learning step completed for {Account}. {Domains} domains, {Senders} senders.",
                SelectedAccount.EmailAddress,
                _learnedProfile.DeletedDomainCounts.Count,
                _learnedProfile.DeletedSenderCounts.Count);

            RefreshProfileStats();
        }
        catch (OperationCanceledException)
        {
            WasStopped = true;
            StatusMessage = "Learning stopped. You can resume or skip to triage.";
            _log.Info("Learning cancelled by user.");
            IsLearning = false;
            CanStop = false;
            return; // Don't auto-advance
        }
        catch (Exception ex)
        {
            HasError = true;
            if (IsComException(ex))
            {
                ErrorMessage = "Lost connection to Outlook. Make sure Outlook is open and click Retry.";
                StatusMessage = "Outlook disconnected.";
            }
            else
            {
                ErrorMessage = $"Learning error: {ex.Message}";
                StatusMessage = "Learning failed. Check Settings → Diagnostics for details.";
            }
            _log.Error(ex, "Learning step failed.");
        }

        IsLearning = false;
        CanStop = false;

        if (HasError)
        {
            ShowStartLearning = true; // Let user retry
            return;
        }

        // Auto-advance to Triage
        GoToStep("Triage");
        await RunTriageAsync();
    }

    [RelayCommand]
    private void StopOperation()
    {
        if (_cts is not null && !_cts.IsCancellationRequested)
        {
            _cts.Cancel();
            CanStop = false;
            StatusMessage = "Stopping... (finishing current item safely)";
            _log.Info("User requested stop.");
        }
    }

    [RelayCommand]
    private async Task ResumeAsync()
    {
        WasStopped = false;

        if (CurrentStep == "Learn")
        {
            await StartLearningAsync();
        }
        else if (CurrentStep == "Triage")
        {
            await RunTriageAsync();
        }
    }

    [RelayCommand]
    private async Task SkipToTriageAsync()
    {
        WasStopped = false;
        // Use whatever profile we have so far (may be empty)
        _learnedProfile ??= new LearnedProfile { AccountKey = SelectedAccount?.AccountKey ?? "" };
        GoToStep("Triage");
        await RunTriageAsync();
    }

    // ═════════════════════════════════════════════
    //   STEP 3: TRIAGE (AI scans inbox → Review folder)
    // ═════════════════════════════════════════════

    private async Task RunTriageAsync()
    {
        if (SelectedAccount == null) return;

        IsTriaging = true;
        WasStopped = false;
        StatusMessage = "AI is scanning your inbox...";

        // Create a fresh cancellation token
        _cts?.Dispose();
        _cts = new CancellationTokenSource();
        CanStop = true;

        try
        {
            // Ensure we have a profile (even if empty)
            _learnedProfile ??= LearnedProfile.Load(SelectedAccount.AccountKey)
                                ?? new LearnedProfile { AccountKey = SelectedAccount.AccountKey };

            // Create Ollama client and triage service
            _ollamaClient ??= new OllamaClient();
            _triageService ??= new TriageService(_connector, _ollamaClient);

            var progress = new Progress<string>(msg => StatusMessage = msg);
            var triageResult = await _triageService.TriageInboxAsync(
                SelectedAccount, _learnedProfile, progress, _cts.Token);

            ReviewCount = triageResult.MovedToReview;

            // Take a snapshot of the review folder for before/after comparison
            if (ReviewCount > 0)
            {
                StatusMessage = "Taking snapshot for review comparison...";
                _reviewSnapshot = await _connector.GetReviewFolderSnapshotAsync(
                    SelectedAccount.EmailAddress, SelectedAccount.StoreName);
            }

            StatusMessage = ReviewCount > 0
                ? $"Triage complete. {ReviewCount} emails moved to Review for Deletion."
                : "Triage complete. No junk detected — your inbox looks clean!";

            _log.Info("Triage step completed. {Count} moved to review. {Scanned} scanned, {Errors} errors.",
                triageResult.MovedToReview, triageResult.TotalScanned, triageResult.Errors);
        }
        catch (OperationCanceledException)
        {
            WasStopped = true;
            // Triage was stopped — any emails already moved are safe in the review folder.
            // Check how many were moved so far.
            var movedSoFar = await _connector.GetReviewFolderCountAsync(
                SelectedAccount.EmailAddress, SelectedAccount.StoreName);
            ReviewCount = movedSoFar;

            StatusMessage = movedSoFar > 0
                ? $"Stopped. {movedSoFar} emails already moved to Review for Deletion. You can resume or proceed to review."
                : "Stopped. No emails were moved yet. You can resume.";
            _log.Info("Triage cancelled by user. {Moved} emails already moved.", movedSoFar);

            // Take a snapshot of whatever was moved
            if (movedSoFar > 0)
            {
                _reviewSnapshot = await _connector.GetReviewFolderSnapshotAsync(
                    SelectedAccount.EmailAddress, SelectedAccount.StoreName);
            }

            IsTriaging = false;
            CanStop = false;
            return; // Stay on Triage step so user can Resume or Skip
        }
        catch (Exception ex)
        {
            HasError = true;
            if (IsComException(ex))
            {
                ErrorMessage = "Lost connection to Outlook during triage. Make sure Outlook is open.";
                StatusMessage = "Outlook disconnected.";
            }
            else if (IsOllamaException(ex))
            {
                ErrorMessage = "Could not reach the AI engine (Ollama). Is it running? Check Settings → AI Engine.";
                StatusMessage = "AI engine unreachable.";
            }
            else
            {
                ErrorMessage = $"Triage error: {ex.Message}";
                StatusMessage = "Triage failed. Check Settings → Diagnostics for details.";
            }
            _log.Error(ex, "Triage step failed.");
        }

        IsTriaging = false;
        CanStop = false;

        if (HasError) return; // Stay on Triage step so user sees the error

        GoToStep("Review");
    }

    // ═════════════════════════════════════════════
    //   STEP 4: REVIEW (user reviews in Outlook)
    // ═════════════════════════════════════════════

    [RelayCommand]
    private async Task ContinueAfterReviewAsync()
    {
        if (SelectedAccount == null) return;

        // If coming from stopped triage, just advance to Review step
        if (CurrentStep == "Triage" && WasStopped)
        {
            WasStopped = false;
            GoToStep("Review");
            return;
        }

        IsProcessing = true;
        WasStopped = false;
        StatusMessage = "Comparing before & after...";

        try
        {
            // 1. Take a fresh snapshot of the review folder
            var afterSnapshot = await _connector.GetReviewFolderSnapshotAsync(
                SelectedAccount.EmailAddress, SelectedAccount.StoreName);

            // 2. Compare with the pre-review snapshot
            var beforeIds = new HashSet<string>(_reviewSnapshot?.Select(s => s.EntryId) ?? []);
            var afterIds = new HashSet<string>(afterSnapshot.Select(s => s.EntryId));

            // Items that were in the folder before but are gone now → user deleted them (agrees with AI)
            var userDeletedIds = beforeIds.Except(afterIds).ToList();
            // Items still remaining → user kept them (disagrees with AI — these are NOT junk)
            var userKeptItems = afterSnapshot;

            // Get sender info for confirmed-junk items (from pre-review snapshot)
            var confirmedJunkItems = _reviewSnapshot!
                .Where(s => userDeletedIds.Contains(s.EntryId))
                .ToList();

            DeletedCount = userDeletedIds.Count;
            KeptCount = userKeptItems.Count;
            RulesCreatedCount = 0;

            _log.Info("Review comparison: {Deleted} agreed (deleted), {Kept} disagreed (kept)",
                DeletedCount, KeptCount);

            // 3. If user kept some emails, move them back to Inbox
            if (userKeptItems.Count > 0)
            {
                StatusMessage = $"Moving {userKeptItems.Count} kept emails back to Inbox...";
                var movedBack = await _connector.BulkMoveReviewToInboxAsync(
                    SelectedAccount.EmailAddress, SelectedAccount.StoreName);

                _log.Info("Moved {Count} kept emails back to Inbox.", movedBack);
            }

            // 4. Update learned profile with corrections and create rules
            if (_learnedProfile != null)
            {
                // ── Do-Not-Delete list: senders the user explicitly kept ──
                foreach (var kept in userKeptItems)
                {
                    var sender = kept.SenderEmail?.ToLowerInvariant() ?? "";
                    if (!string.IsNullOrEmpty(sender))
                    {
                        _learnedProfile.DoNotDeleteSenders.Add(sender);
                        _learnedProfile.DeletedSenderCounts.Remove(sender);
                    }

                    var domain = kept.Domain?.ToLowerInvariant() ?? "";
                    if (!string.IsNullOrEmpty(domain))
                    {
                        if (_learnedProfile.DeletedDomainCounts.TryGetValue(domain, out int current))
                        {
                            if (current <= 1)
                                _learnedProfile.DeletedDomainCounts.Remove(domain);
                            else
                                _learnedProfile.DeletedDomainCounts[domain] = current - 1;
                        }
                    }
                }

                _log.Info("Updated DoNotDelete list with {Count} senders from kept emails.", userKeptItems.Count);

                // ── Create Outlook rules for confirmed-junk senders ──
                var newJunkSenders = confirmedJunkItems
                    .Select(s => s.SenderEmail)
                    .Where(s => !string.IsNullOrEmpty(s)
                             && !_learnedProfile.DoNotDeleteSenders.Contains(s)
                             && !_learnedProfile.RuleCreatedSenders.Contains(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (newJunkSenders.Count > 0)
                {
                    StatusMessage = $"Creating Outlook rules for {newJunkSenders.Count} junk senders...";
                    var rulesProgress = new Progress<string>(msg => StatusMessage = msg);
                    RulesCreatedCount = await _connector.CreateSenderRulesAsync(
                        SelectedAccount.EmailAddress, SelectedAccount.StoreName,
                        newJunkSenders, rulesProgress);

                    // Track which senders now have rules
                    foreach (var s in newJunkSenders)
                        _learnedProfile.RuleCreatedSenders.Add(s);

                    _log.Info("Created {Count} Outlook rules for confirmed junk senders.", RulesCreatedCount);
                }

                _learnedProfile.Save();
                _log.Info("Saved learned profile with {DoNot} protected senders, {Rules} ruled senders.",
                    _learnedProfile.DoNotDeleteSenders.Count, _learnedProfile.RuleCreatedSenders.Count);
            }

            StatusMessage = RulesCreatedCount > 0
                ? $"Done! Deleted: {DeletedCount}, Kept: {KeptCount}, Rules created: {RulesCreatedCount}."
                : $"Done! Deleted: {DeletedCount}, Kept: {KeptCount}.";
            _log.Info("Review step completed. Deleted={Deleted}, Kept={Kept}, Rules={Rules}",
                DeletedCount, KeptCount, RulesCreatedCount);
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error: {ex.Message}";
            _log.Error(ex, "Continue-after-review failed.");
        }

        IsProcessing = false;
        GoToStep("Automate");
    }

    // ═════════════════════════════════════════════
    //   STEP 5: AUTOMATE (results + rules created)
    // ═════════════════════════════════════════════

    [RelayCommand]
    private void RunAgain()
    {
        _cts?.Cancel();
        _cts?.Dispose();
        _cts = null;
        _reviewSnapshot = null;
        ReviewCount = 0;
        DeletedCount = 0;
        KeptCount = 0;
        RulesCreatedCount = 0;
        CanStop = false;
        WasStopped = false;
        StatusMessage = "";
        GoToStep("Learn");

        ShowStartLearning = true;
        ShowAccountPicker = Accounts.Count > 1;
    }

    // ═════════════════════════════════════════════
    //   SETTINGS
    // ═════════════════════════════════════════════

    [RelayCommand]
    private void ToggleSettings()
    {
        ShowSettings = !ShowSettings;
        if (ShowSettings) RefreshProfileStats();
    }

    [RelayCommand]
    private async Task CheckAiUpdatesAsync()
    {
        _ollamaSetup ??= new OllamaSetupService();
        StatusMessage = "Checking for AI updates...";

        try
        {
            var progress = new Progress<string>(msg => StatusMessage = msg);
            var result = await _ollamaSetup.CheckForUpdatesAsync(progress);

            AiOllamaVersion = result.CurrentVersion;
            AiModelName = result.ModelName;
            AiModelSize = result.ModelSize;

            StatusMessage = "Checking for model updates...";
            await _ollamaSetup.PullModelAsync("gemma3:4b", progress);

            var (_, newName, newSize) = await _ollamaSetup.CheckModelInstalledAsync();
            AiModelName = newName;
            AiModelSize = newSize;

            StatusMessage = $"Up to date. Ollama v{result.CurrentVersion} — {newName} ({newSize}).";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Update check failed: {ex.Message}";
        }
    }

    [RelayCommand]
    private async Task InstallOllamaAsync()
    {
        _ollamaSetup ??= new OllamaSetupService();
        IsInstallingOllama = true;
        StatusMessage = "Downloading Ollama...";

        try
        {
            var progress = new Progress<string>(msg => StatusMessage = msg);
            var ok = await _ollamaSetup.InstallOllamaAsync(progress);

            if (ok)
            {
                AiStatus = "Installed";
                AiStatusColor = "#2196F3";
                StatusMessage = "Ollama installed! Pulling AI model...";

                await _ollamaSetup.PullModelAsync("gemma3:4b", progress);
                var (installed, name, size) = await _ollamaSetup.CheckModelInstalledAsync();
                if (installed)
                {
                    AiModelName = name;
                    AiModelSize = size;
                    IsAiReady = true;
                    AiStatus = "Ready";
                    AiStatusColor = "#00C853";
                    AiOllamaVersion = await _ollamaSetup.GetOllamaVersionAsync();
                    StatusMessage = $"AI ready: {name} ({size})";
                }
            }
            else
            {
                StatusMessage = "Ollama installation failed. Please install manually from ollama.com.";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Install failed: {ex.Message}";
            _log.Error(ex, "Ollama installation failed.");
        }

        IsInstallingOllama = false;
    }

    [RelayCommand]
    private void ExportLogs()
    {
        try
        {
            var logDir = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "EmailManage");
            System.IO.Directory.CreateDirectory(logDir);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = logDir,
                UseShellExecute = true
            });
            StatusMessage = $"Opened log folder: {logDir}";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Could not open log folder: {ex.Message}";
        }
    }

    [RelayCommand]
    private void ResetProfile()
    {
        if (SelectedAccount == null) return;

        try
        {
            var profile = new LearnedProfile { AccountKey = SelectedAccount.AccountKey };
            profile.Save();
            _learnedProfile = profile;
            RefreshProfileStats();
            StatusMessage = "Learned profile has been reset.";
            _log.Info("Profile reset for {Account}", SelectedAccount.EmailAddress);
        }
        catch (Exception ex)
        {
            StatusMessage = $"Reset failed: {ex.Message}";
        }
    }

    private void RefreshProfileStats()
    {
        if (_learnedProfile == null)
        {
            ProtectedSenderCount = 0;
            RuledSenderCount = 0;
            LearnedDomainCount = 0;
            ProfileLastLearnedAt = null;
            return;
        }

        ProtectedSenderCount = _learnedProfile.DoNotDeleteSenders.Count;
        RuledSenderCount = _learnedProfile.RuleCreatedSenders.Count;
        LearnedDomainCount = _learnedProfile.DeletedDomainCounts.Count;
        ProfileLastLearnedAt = _learnedProfile.LearnedAt == default
            ? "Never"
            : _learnedProfile.LearnedAt.ToLocalTime().ToString("MMM d, yyyy h:mm tt");
    }

    // ═════════════════════════════════════════════
    //   AI SETUP (runs silently during connect)
    // ═════════════════════════════════════════════

    private async Task SetupAiSilentAsync()
    {
        _ollamaSetup ??= new OllamaSetupService();

        try
        {
            if (!_ollamaSetup.IsOllamaInstalled())
            {
                AiStatus = "Not installed";
                AiStatusColor = "#FF9800";
                _log.Info("Ollama not found — user can install via Settings.");
                return;
            }

            if (!await _ollamaSetup.IsOllamaRunningAsync())
            {
                AiStatus = "Starting...";
                AiStatusColor = "#2196F3";
                var progress = new Progress<string>(_ => { });
                var started = await _ollamaSetup.StartOllamaAsync(progress);
                if (!started)
                {
                    AiStatus = "Not running";
                    AiStatusColor = "#F44336";
                    return;
                }
            }

            AiOllamaVersion = await _ollamaSetup.GetOllamaVersionAsync();
            var (modelInstalled, modelName, modelSize) = await _ollamaSetup.CheckModelInstalledAsync();

            if (!modelInstalled)
            {
                AiStatus = "No model";
                AiStatusColor = "#FF9800";
                return;
            }

            AiModelName = modelName;
            AiModelSize = modelSize;
            IsAiReady = true;
            AiStatus = "Ready";
            AiStatusColor = "#00C853";
            _log.Info("AI ready: {Model} ({Size}), Ollama v{Ver}", modelName, modelSize, AiOllamaVersion);
        }
        catch (Exception ex)
        {
            AiStatus = "Error";
            AiStatusColor = "#F44336";
            _log.Error(ex, "AI silent setup failed.");
        }
    }

    // ═════════════════════════════════════════════
    //   ERROR DETECTION HELPERS
    // ═════════════════════════════════════════════

    /// <summary>
    /// Detects COM/Outlook disconnection errors by HResult or type.
    /// </summary>
    private static bool IsComException(Exception ex)
    {
        if (ex is System.Runtime.InteropServices.COMException) return true;
        if (ex is InvalidComObjectException) return true;
        // Check HResult for RPC server unavailable (0x800706BA)
        // or call rejected (0x80010001)
        if (ex.HResult == unchecked((int)0x800706BA) ||
            ex.HResult == unchecked((int)0x80010001) ||
            ex.HResult == unchecked((int)0x80080005))
            return true;
        return ex.InnerException != null && IsComException(ex.InnerException);
    }

    /// <summary>
    /// Detects Ollama connectivity errors (connection refused, timeout, etc.).
    /// </summary>
    private static bool IsOllamaException(Exception ex)
    {
        if (ex is System.Net.Http.HttpRequestException) return true;
        if (ex is TaskCanceledException tce && tce.InnerException is TimeoutException) return true;
        var msg = ex.Message.ToLowerInvariant();
        if (msg.Contains("connection refused") || msg.Contains("actively refused") ||
            msg.Contains("127.0.0.1:11434"))
            return true;
        return ex.InnerException != null && IsOllamaException(ex.InnerException);
    }

    // ═════════════════════════════════════════════
    //   CLEANUP
    // ═════════════════════════════════════════════

    public void Cleanup()
    {
        _cts?.Cancel();
        _cts?.Dispose();
        _connector.Dispose();
    }
}
