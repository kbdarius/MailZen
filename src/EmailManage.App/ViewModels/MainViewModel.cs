using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using EmailManage.Models;
using EmailManage.Services;

namespace EmailManage.ViewModels;

/// <summary>
/// ViewModel for the main window. Manages Outlook connection and account selection.
/// </summary>
public partial class MainViewModel : ObservableObject
{
    private readonly OutlookConnectorService _connector;
    private readonly DiagnosticLogger _log;

    [ObservableProperty]
    private string _connectionStatus = "Disconnected";

    [ObservableProperty]
    private string _connectionStatusColor = "#EF5350"; // Red for disconnected

    [ObservableProperty]
    private string? _errorMessage;

    [ObservableProperty]
    private string? _errorCode;

    [ObservableProperty]
    private bool _isConnecting;

    [ObservableProperty]
    private bool _isConnected;

    [ObservableProperty]
    private bool _hasError;

    [ObservableProperty]
    private OutlookAccountInfo? _selectedAccount;

    [ObservableProperty]
    private string _statusBarText = "Ready. Click Connect to start.";

    [ObservableProperty]
    private string _appTitle = "✨ MailZen";

    [ObservableProperty]
    private string _currentView = "Connect"; // "Connect" or "Analyze"

    [ObservableProperty]
    private bool _isAnalyzing;

    [ObservableProperty]
    private string _analysisProgress = "";

    [ObservableProperty]
    private AnalysisResult? _lastAnalysisResult;

    [ObservableProperty]
    private bool _isTriaging;

    [ObservableProperty]
    private string _triageProgress = "";

    [ObservableProperty]
    private TriageResult? _lastTriageResult;

    [ObservableProperty]
    private bool _isLearning;

    [ObservableProperty]
    private string _learnProgress = "";

    [ObservableProperty]
    private LearnResult? _lastLearnResult;

    // ── Improve Model state ──
    [ObservableProperty]
    private bool _isLoadingReview;

    [ObservableProperty]
    private bool _isProcessingAction;

    [ObservableProperty]
    private string _improveProgress = "";

    [ObservableProperty]
    private int _improveNeedsReviewCount;

    [ObservableProperty]
    private bool _isImproveReviewing; // user is reviewing in Outlook

    [ObservableProperty]
    private int _improveDeletedCount;

    [ObservableProperty]
    private int _improveKeptCount;

    [ObservableProperty]
    private int _improveNewRulesAdded;

    [ObservableProperty]
    private bool _improveShowResults; // show comparison results

    private List<FolderEmailSnapshot>? _improveSnapshot; // snapshot before user reviews

    // ── AI Smart Scan state ──
    [ObservableProperty]
    private string _aiStatus = "Not installed";

    [ObservableProperty]
    private string _aiStatusColor = "#9E9E9E"; // Gray

    [ObservableProperty]
    private bool _isAiSetupRunning;

    [ObservableProperty]
    private string _aiSetupProgress = "";

    [ObservableProperty]
    private bool _isAiReady;

    [ObservableProperty]
    private bool _isAiScanning;

    [ObservableProperty]
    private string _aiScanProgress = "";

    [ObservableProperty]
    private int _aiScanTotal;

    [ObservableProperty]
    private int _aiGapCount;

    [ObservableProperty]
    private int _aiScanSeconds;

    [ObservableProperty]
    private long _aiAvgLatency;

    [ObservableProperty]
    private bool _isAiReviewPending; // emails moved to AI Review, waiting for user

    [ObservableProperty]
    private bool _isAiMoveBackPending; // user done reviewing, confirm move-back

    [ObservableProperty]
    private int _aiUserDeletedCount;

    [ObservableProperty]
    private int _aiUserKeptCount;

    [ObservableProperty]
    private int _aiNewRulesAdded;

    [ObservableProperty]
    private bool _aiShowResults;

    [ObservableProperty]
    private string? _aiModelName;

    [ObservableProperty]
    private string? _aiModelSize;

    [ObservableProperty]
    private string? _aiOllamaVersion;

    private OllamaSetupService? _ollamaSetup;
    private CancellationTokenSource? _aiScanCts;
    private List<FolderEmailSnapshot>? _aiReviewSnapshot; // snapshot when moved to AI Review
    private List<AiScanItem>? _lastAiGapItems; // the gap items for rule learning

    public ObservableCollection<OutlookAccountInfo> Accounts { get; } = [];

    public MainViewModel()
    {
        _connector = new OutlookConnectorService();
        _log = DiagnosticLogger.Instance;
    }

    [RelayCommand]
    private void NavigateTo(string viewName)
    {
        if (IsConnected && SelectedAccount != null)
        {
            CurrentView = viewName;
        }
    }

    [RelayCommand]
    private async Task AnalyzePatternsAsync()
    {
        if (SelectedAccount == null) return;

        IsAnalyzing = true;
        AnalysisProgress = "Starting analysis...";
        LastAnalysisResult = null;

        var progress = new Progress<string>(msg => AnalysisProgress = msg);
        var service = new PatternAnalysisService(_connector);

        var result = await Task.Run(() => service.AnalyzeAccountAsync(SelectedAccount, progress));

        LastAnalysisResult = result;
        IsAnalyzing = false;
        AnalysisProgress = result.Success ? "Analysis complete." : $"Analysis failed: {result.ErrorMessage}";

        // Update account status
        if (result.Success)
        {
            SelectedAccount.HasPatternFile = true;
            // Force UI refresh
            var index = Accounts.IndexOf(SelectedAccount);
            if (index >= 0)
            {
                Accounts[index] = SelectedAccount;
                SelectedAccount = Accounts[index];
            }
        }
    }

    [RelayCommand]
    private async Task RunTriageAsync()
    {
        if (SelectedAccount == null) return;

        IsTriaging = true;
        TriageProgress = "Starting triage...";
        LastTriageResult = null;

        var progress = new Progress<string>(msg => TriageProgress = msg);
        var service = new TriageService(_connector);

        var result = await Task.Run(() => service.RunTriageAsync(SelectedAccount, progress));

        LastTriageResult = result;
        IsTriaging = false;
        TriageProgress = result.Success ? "Triage complete." : $"Triage failed: {result.ErrorMessage}";
    }

    [RelayCommand]
    private async Task LearnFromDeletionsAsync()
    {
        if (SelectedAccount == null) return;

        IsLearning = true;
        LearnProgress = "Scanning recent deletions...";
        LastLearnResult = null;

        var progress = new Progress<string>(msg => LearnProgress = msg);
        var service = new PatternAnalysisService(_connector);

        var result = await Task.Run(() => service.LearnFromRecentDeletionsAsync(SelectedAccount, 7, progress));

        LastLearnResult = result;
        IsLearning = false;
        LearnProgress = result.Success 
            ? $"Done! Added {result.NewPatternsAdded} new rules." 
            : $"Failed: {result.ErrorMessage}";
    }

    // ═══ Improve Model Commands (Folder-based: review in Outlook) ═══

    [RelayCommand]
    private async Task StartReviewAsync()
    {
        if (SelectedAccount == null) return;

        IsLoadingReview = true;
        ImproveProgress = "Checking Needs Review folder...";
        ImproveShowResults = false;

        try
        {
            // Get count of items in Needs Review
            var count = await _connector.GetFolderItemCountAsync(
                SelectedAccount.EmailAddress, SelectedAccount.StoreName, OutlookConnectorService.NeedsReviewFolderName);

            ImproveNeedsReviewCount = count;

            if (count == 0)
            {
                ImproveProgress = "No items in Needs Review folder. Run Triage first to populate it.";
                IsLoadingReview = false;
                return;
            }

            // Take snapshot of the folder
            ImproveProgress = $"Taking snapshot of {count} items...";
            _improveSnapshot = await _connector.GetFolderSnapshotAsync(
                SelectedAccount.EmailAddress, SelectedAccount.StoreName, OutlookConnectorService.NeedsReviewFolderName);

            ImproveProgress = $"Needs Review has {count} emails. Open Outlook → Smart Cleanup → Needs Review. Delete what's junk, then click 'Done Reviewing'.";
            IsImproveReviewing = true;
        }
        catch (Exception ex)
        {
            ImproveProgress = $"Error: {ex.Message}";
        }

        IsLoadingReview = false;
    }

    [RelayCommand]
    private async Task DoneReviewingImproveAsync()
    {
        if (SelectedAccount == null || _improveSnapshot == null) return;

        IsProcessingAction = true;
        ImproveProgress = "Comparing before & after...";

        try
        {
            // Get current state of the folder
            var afterSnapshot = await _connector.GetFolderSnapshotAsync(
                SelectedAccount.EmailAddress, SelectedAccount.StoreName, OutlookConnectorService.NeedsReviewFolderName);

            var afterIds = new HashSet<string>(afterSnapshot.Select(s => s.EntryId));
            var deletedItems = _improveSnapshot.Where(s => !afterIds.Contains(s.EntryId)).ToList();
            var keptItems = afterSnapshot; // what's still there

            ImproveDeletedCount = deletedItems.Count;
            ImproveKeptCount = keptItems.Count;

            // Learn from deletions: add deleted senders to rules
            if (deletedItems.Count > 0)
            {
                ImproveProgress = $"You deleted {deletedItems.Count} emails (agreed they're junk). Learning from your choices...";
                var existingPatterns = LoadPatternsForAccount(SelectedAccount.AccountKey);
                var existingSenders = new HashSet<string>(
                    existingPatterns.Where(p => p.Type == "Sender" || p.Type == "SenderWithUnsubscribe")
                                    .Select(p => p.Value.ToLowerInvariant()));

                var newPatterns = new List<PatternRule>();
                foreach (var item in deletedItems)
                {
                    if (string.IsNullOrWhiteSpace(item.SenderEmail)) continue;
                    if (existingSenders.Contains(item.SenderEmail)) continue;

                    newPatterns.Add(new PatternRule
                    {
                        RuleId = $"rule_improve_{Guid.NewGuid().ToString("N")[..8]}",
                        Type = "Sender",
                        Value = item.SenderEmail,
                        Weight = 0.9,
                        EvidenceCount = 1
                    });
                    existingSenders.Add(item.SenderEmail);
                }

                if (newPatterns.Count > 0)
                {
                    var merged = existingPatterns.Concat(newPatterns).ToList();
                    SavePatternFile(SelectedAccount.AccountKey, merged);
                }

                ImproveNewRulesAdded = newPatterns.Count;
            }

            // Move remaining items back to Inbox
            if (keptItems.Count > 0)
            {
                ImproveProgress = $"Moving {keptItems.Count} kept emails back to Inbox...";
                var progress = new Progress<string>(msg => ImproveProgress = msg);
                var movedBack = await _connector.BulkMoveAllToInboxAsync(
                    SelectedAccount.EmailAddress, SelectedAccount.StoreName,
                    OutlookConnectorService.NeedsReviewFolderName, progress);

                ImproveProgress = $"Done! Deleted: {deletedItems.Count} (learned {ImproveNewRulesAdded} new rules). " +
                                  $"Kept: {keptItems.Count} (moved back to Inbox).";
            }
            else
            {
                ImproveProgress = $"Done! You deleted all {deletedItems.Count} items. Learned {ImproveNewRulesAdded} new rules.";
            }

            IsImproveReviewing = false;
            ImproveShowResults = true;
            _improveSnapshot = null;
        }
        catch (Exception ex)
        {
            ImproveProgress = $"Error: {ex.Message}";
        }

        IsProcessingAction = false;
    }

    [RelayCommand]
    private async Task ConnectAsync()
    {
        IsConnecting = true;
        HasError = false;
        ErrorMessage = null;
        ErrorCode = null;
        ConnectionStatus = "Connecting...";
        ConnectionStatusColor = "#2196F3"; // Blue
        StatusBarText = "Connecting to Outlook...";
        Accounts.Clear();
        SelectedAccount = null;

        _log.Info("User initiated Outlook connection.");

        var result = await _connector.ConnectAsync();

        IsConnecting = false;

        if (result.Success)
        {
            IsConnected = true;
            ConnectionStatus = "Connected";
        ConnectionStatusColor = "#00C853"; // Vivid green
            foreach (var account in result.Accounts)
                Accounts.Add(account);

            StatusBarText = $"Connected. {Accounts.Count} account(s) found. ({result.Elapsed.TotalMilliseconds:F0}ms)";
            _log.Info("AccountsLoaded: {Count} accounts displayed.", Accounts.Count);

            // Auto-select first account if only one
            if (Accounts.Count == 1)
                SelectedAccount = Accounts[0];
        }
        else
        {
            IsConnected = false;
            HasError = true;
            ErrorMessage = result.ErrorMessage;
            ErrorCode = result.ErrorCode;
            ConnectionStatus = "Error";
            ConnectionStatusColor = "#F44336"; // Red
            StatusBarText = $"Connection failed: {result.ErrorCode}";
        }
    }

    [RelayCommand]
    private async Task RetryAsync()
    {
        await ConnectAsync();
    }

    [RelayCommand]
    private void Disconnect()
    {
        _connector.Disconnect();
        IsConnected = false;
        HasError = false;
        ErrorMessage = null;
        ConnectionStatus = "Disconnected";
        ConnectionStatusColor = "#EF5350"; // Red
        Accounts.Clear();
        SelectedAccount = null;
        StatusBarText = "Disconnected.";
        _log.Info("User disconnected from Outlook.");
    }

    partial void OnSelectedAccountChanged(OutlookAccountInfo? value)
    {
        if (value is not null)
        {
            _log.Info("AccountSelected: {Key} ({Email})", value.AccountKey, value.EmailAddress);
            StatusBarText = $"Active account: {value.EmailAddress}";
        }
    }

    // ═══ AI Smart Scan Commands ═══

    [RelayCommand]
    private async Task SetupAiAsync()
    {
        IsAiSetupRunning = true;
        AiSetupProgress = "Checking Ollama status...";
        _ollamaSetup ??= new OllamaSetupService();
        var progress = new Progress<string>(msg => AiSetupProgress = msg);

        try
        {
            // Step 1: Check if Ollama is installed
            if (!_ollamaSetup.IsOllamaInstalled())
            {
                AiSetupProgress = "Ollama not found. Downloading and installing...";
                AiStatus = "Installing...";
                AiStatusColor = "#2196F3"; // Blue

                var installed = await _ollamaSetup.InstallOllamaAsync(progress);
                if (!installed)
                {
                    AiSetupProgress = "Failed to install Ollama. Please install manually from ollama.com";
                    AiStatus = "Install failed";
                    AiStatusColor = "#F44336";
                    IsAiSetupRunning = false;
                    return;
                }
            }

            // Step 2: Check if Ollama API is running
            if (!await _ollamaSetup.IsOllamaRunningAsync())
            {
                AiSetupProgress = "Starting Ollama service...";
                AiStatus = "Starting...";
                AiStatusColor = "#2196F3";

                var started = await _ollamaSetup.StartOllamaAsync(progress);
                if (!started)
                {
                    AiSetupProgress = "Could not start Ollama. Try opening the Ollama app manually.";
                    AiStatus = "Not running";
                    AiStatusColor = "#F44336";
                    IsAiSetupRunning = false;
                    return;
                }
            }

            // Step 3: Check if model is pulled
            AiOllamaVersion = await _ollamaSetup.GetOllamaVersionAsync();
            var (modelInstalled, modelName, modelSize) = await _ollamaSetup.CheckModelInstalledAsync();

            if (!modelInstalled)
            {
                AiSetupProgress = "Downloading AI model (gemma3:4b ~3GB)... this takes a few minutes.";
                AiStatus = "Pulling model...";
                AiStatusColor = "#FF9800"; // Orange

                var pulled = await _ollamaSetup.PullModelAsync("gemma3:4b", progress);
                if (!pulled)
                {
                    AiSetupProgress = "Failed to download model. Check your internet connection and try again.";
                    AiStatus = "Model failed";
                    AiStatusColor = "#F44336";
                    IsAiSetupRunning = false;
                    return;
                }

                (modelInstalled, modelName, modelSize) = await _ollamaSetup.CheckModelInstalledAsync();
            }

            AiModelName = modelName;
            AiModelSize = modelSize;
            IsAiReady = true;
            AiStatus = "Ready";
            AiStatusColor = "#00C853"; // Green
            AiSetupProgress = $"AI is ready! Model: {modelName} ({modelSize}). Ollama v{AiOllamaVersion}.";
        }
        catch (Exception ex)
        {
            AiSetupProgress = $"Setup error: {ex.Message}";
            AiStatus = "Error";
            AiStatusColor = "#F44336";
        }

        IsAiSetupRunning = false;
    }

    [RelayCommand]
    private async Task CheckAiUpdatesAsync()
    {
        _ollamaSetup ??= new OllamaSetupService();
        AiSetupProgress = "Checking for updates...";

        try
        {
            var progress = new Progress<string>(msg => AiSetupProgress = msg);
            var result = await _ollamaSetup.CheckForUpdatesAsync(progress);

            AiOllamaVersion = result.CurrentVersion;
            AiModelName = result.ModelName;
            AiModelSize = result.ModelSize;

            // Re-pull model to get latest version (no-op if already up to date)
            AiSetupProgress = "Checking for model updates...";
            await _ollamaSetup.PullModelAsync("gemma3:4b", progress);

            var (_, newName, newSize) = await _ollamaSetup.CheckModelInstalledAsync();
            AiModelName = newName;
            AiModelSize = newSize;

            AiSetupProgress = $"Up to date. Ollama v{result.CurrentVersion} — Model: {newName} ({newSize}).";
        }
        catch (Exception ex)
        {
            AiSetupProgress = $"Update check failed: {ex.Message}";
        }
    }

    [RelayCommand]
    private async Task RunAiScanAsync()
    {
        if (SelectedAccount == null || !IsAiReady) return;

        IsAiScanning = true;
        AiScanProgress = "Starting AI scan...";
        AiShowResults = false;
        IsAiReviewPending = false;
        IsAiMoveBackPending = false;
        _aiScanCts = new CancellationTokenSource();

        var progress = new Progress<string>(msg => AiScanProgress = msg);
        var service = new AiScanService(_connector);

        var result = await service.RunBenchmarkAsync(SelectedAccount, 100, progress, _aiScanCts.Token);

        IsAiScanning = false;

        if (result.Success)
        {
            AiScanTotal = result.TotalScanned;
            AiGapCount = result.GapCount;
            AiScanSeconds = result.TotalTimeSeconds;
            AiAvgLatency = result.AvgLatencyMs;
            _lastAiGapItems = result.Items.Where(i => i.IsAiGap).ToList();

            if (result.GapCount > 0)
            {
                AiScanProgress = $"AI found {result.GapCount} junk emails that rules missed. Click 'Move to AI Review' to review them in Outlook.";
                AiShowResults = true;
            }
            else
            {
                AiScanProgress = $"Scan complete. AI agrees with your rules — no gaps found! ({result.TotalScanned} scanned in {result.TotalTimeSeconds}s)";
            }
        }
    }

    [RelayCommand]
    private void CancelAiScan()
    {
        _aiScanCts?.Cancel();
        AiScanProgress = "Cancelling...";
    }

    [RelayCommand]
    private async Task MoveToAiReviewAsync()
    {
        if (SelectedAccount == null || _lastAiGapItems == null || _lastAiGapItems.Count == 0) return;

        IsAiScanning = true;
        AiScanProgress = $"Moving {_lastAiGapItems.Count} gap emails to AI Review folder...";

        try
        {
            var progress = new Progress<string>(msg => AiScanProgress = msg);
            var entryIds = _lastAiGapItems.Select(i => i.EntryId).ToList();

            var moved = await _connector.BulkMoveToTriageFolderAsync(
                entryIds, SelectedAccount.EmailAddress, SelectedAccount.StoreName,
                OutlookConnectorService.AiReviewFolderName, progress);

            // Take snapshot of AI Review folder (items now have new EntryIDs)
            _aiReviewSnapshot = await _connector.GetFolderSnapshotAsync(
                SelectedAccount.EmailAddress, SelectedAccount.StoreName,
                OutlookConnectorService.AiReviewFolderName);

            AiScanProgress = $"Moved {moved} emails to Smart Cleanup → AI Review.\n" +
                             $"Open Outlook → Smart Cleanup → AI Review. Delete what's junk, then click 'Done Reviewing'.";
            IsAiReviewPending = true;
            IsAiScanning = false;
        }
        catch (Exception ex)
        {
            AiScanProgress = $"Error: {ex.Message}";
            IsAiScanning = false;
        }
    }

    [RelayCommand]
    private async Task DoneReviewingAiAsync()
    {
        if (SelectedAccount == null || _aiReviewSnapshot == null) return;

        IsAiScanning = true;
        AiScanProgress = "Comparing your review...";

        try
        {
            // Get current state of AI Review folder
            var afterSnapshot = await _connector.GetFolderSnapshotAsync(
                SelectedAccount.EmailAddress, SelectedAccount.StoreName,
                OutlookConnectorService.AiReviewFolderName);

            var afterIds = new HashSet<string>(afterSnapshot.Select(s => s.EntryId));
            var deletedItems = _aiReviewSnapshot.Where(s => !afterIds.Contains(s.EntryId)).ToList();
            var keptItems = afterSnapshot;

            AiUserDeletedCount = deletedItems.Count;
            AiUserKeptCount = keptItems.Count;

            // Learn from deletions
            if (deletedItems.Count > 0)
            {
                AiScanProgress = $"You deleted {deletedItems.Count} emails (agreed with AI). Learning new rules...";
                var existingPatterns = LoadPatternsForAccount(SelectedAccount.AccountKey);
                var existingSenders = new HashSet<string>(
                    existingPatterns.Where(p => p.Type == "Sender" || p.Type == "SenderWithUnsubscribe")
                                    .Select(p => p.Value.ToLowerInvariant()));

                var newPatterns = new List<PatternRule>();
                foreach (var item in deletedItems)
                {
                    if (string.IsNullOrWhiteSpace(item.SenderEmail)) continue;
                    if (existingSenders.Contains(item.SenderEmail)) continue;

                    newPatterns.Add(new PatternRule
                    {
                        RuleId = $"rule_ai_{Guid.NewGuid().ToString("N")[..8]}",
                        Type = "Sender",
                        Value = item.SenderEmail,
                        Weight = 0.9,
                        EvidenceCount = 1
                    });
                    existingSenders.Add(item.SenderEmail);
                }

                if (newPatterns.Count > 0)
                {
                    var merged = existingPatterns.Concat(newPatterns).ToList();
                    SavePatternFile(SelectedAccount.AccountKey, merged);
                }

                AiNewRulesAdded = newPatterns.Count;
            }

            IsAiReviewPending = false;

            if (keptItems.Count > 0)
            {
                AiScanProgress = $"Agreed: {deletedItems.Count} deleted ({AiNewRulesAdded} new rules learned). " +
                                 $"{keptItems.Count} emails remain — move them back to Inbox?";
                IsAiMoveBackPending = true;
            }
            else
            {
                AiScanProgress = $"Done! You deleted all {deletedItems.Count} items. Learned {AiNewRulesAdded} new rules. Run Triage again to apply.";
                AiShowResults = false;
                _aiReviewSnapshot = null;
            }
        }
        catch (Exception ex)
        {
            AiScanProgress = $"Error: {ex.Message}";
        }

        IsAiScanning = false;
    }

    [RelayCommand]
    private async Task MoveAiRemainingBackAsync()
    {
        if (SelectedAccount == null) return;

        IsAiScanning = true;
        AiScanProgress = "Moving remaining emails back to Inbox...";

        try
        {
            var progress = new Progress<string>(msg => AiScanProgress = msg);
            var movedBack = await _connector.BulkMoveAllToInboxAsync(
                SelectedAccount.EmailAddress, SelectedAccount.StoreName,
                OutlookConnectorService.AiReviewFolderName, progress);

            AiScanProgress = $"Complete! Deleted: {AiUserDeletedCount} ({AiNewRulesAdded} new rules). " +
                             $"Moved back to Inbox: {movedBack}. Run Triage again to apply new rules.";
            IsAiMoveBackPending = false;
            AiShowResults = false;
            _aiReviewSnapshot = null;
        }
        catch (Exception ex)
        {
            AiScanProgress = $"Error: {ex.Message}";
        }

        IsAiScanning = false;
    }

    private List<PatternRule> LoadPatternsForAccount(string accountKey)
    {
        var patterns = new List<PatternRule>();
        var path = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "EmailManage", "accounts", accountKey, "pattern_v1.yaml");
        if (!System.IO.File.Exists(path)) return patterns;

        var lines = System.IO.File.ReadAllLines(path);
        PatternRule? cur = null;
        foreach (var line in lines)
        {
            var t = line.Trim();
            if (t.StartsWith("- id:")) { if (cur != null) patterns.Add(cur); cur = new PatternRule { RuleId = t[5..].Trim() }; }
            else if (cur != null)
            {
                if (t.StartsWith("type:")) cur.Type = t[5..].Trim();
                else if (t.StartsWith("value:")) cur.Value = t[6..].Trim().Trim('"');
                else if (t.StartsWith("weight:") && double.TryParse(t[7..].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var w)) cur.Weight = w;
                else if (t.StartsWith("evidence_count:") && int.TryParse(t[15..].Trim(), out var e)) cur.EvidenceCount = e;
            }
        }
        if (cur != null) patterns.Add(cur);
        return patterns;
    }

    private void SavePatternFile(string accountKey, List<PatternRule> patterns)
    {
        var dir = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "EmailManage", "accounts", accountKey);
        System.IO.Directory.CreateDirectory(dir);
        var path = System.IO.Path.Combine(dir, "pattern_v1.yaml");
        using var writer = new System.IO.StreamWriter(path);
        writer.WriteLine("version: 1");
        writer.WriteLine($"last_updated: {DateTime.UtcNow:O}");
        writer.WriteLine("rules:");
        foreach (var p in patterns)
        {
            writer.WriteLine($"  - id: {p.RuleId}");
            writer.WriteLine($"    type: {p.Type}");
            writer.WriteLine($"    value: \"{p.Value}\"");
            writer.WriteLine($"    weight: {p.Weight.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)}");
            writer.WriteLine($"    evidence_count: {p.EvidenceCount}");
        }
    }

    public void Cleanup()
    {
        _aiScanCts?.Cancel();
        _connector.Dispose();
    }
}
