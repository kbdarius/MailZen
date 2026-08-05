using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using System.Net.Http;
using System.Reflection;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Globalization;
using EmailManage.Models;
using EmailManage.Services;

namespace EmailManage;

public partial class SearchWindow : Window
{
    public SearchWindow()
    {
        InitializeComponent();
        DataContext = new SearchWorkspaceViewModel();
        Loaded += async (_, _) => await ((SearchWorkspaceViewModel)DataContext).LoadAccountsAsync();
    }
}

public sealed class SearchWorkspaceViewModel : INotifyPropertyChanged
{
    private readonly OutlookReadAdapter _outlook = new();
    private readonly MailZenDatabase _database = new();
    private string _query = string.Empty;
    private string _status = "Connecting to Outlook…";
    private string _resultSummary = string.Empty;
    private string _scopeSummary = "No accounts selected.";
    private string _indexSummary = "Index coverage will appear here.";
    private bool _localSearchOnly;
    private string _apiKey = string.Empty;
    private DateTime? _indexFromDate = DateTime.Today.AddYears(-1);
    private DateTime? _indexToDate = DateTime.Today;
    private string _selectedModelId = MailZenModelRegistry.FastModelId;
    private SearchModeOption _selectedSearchMode = null!;
    private double _dateRangeMinimum;
    private double _dateRangeMaximum;
    private double _searchStartValue;
    private double _searchEndValue;
    private bool _isBusy;
    private string _activityStatus = string.Empty;
    private string _dailySyncTime = "02:00";
    private readonly OpenAiCredentialStore _credentials = new();
    public ObservableCollection<SearchAccountRow> Accounts { get; } = new();
    public ObservableCollection<SearchResultRow> Results { get; } = new();
    public string Query { get => _query; set { _query = value; OnChanged(); } }
    public string AppVersion { get; } = Assembly.GetEntryAssembly()?.GetName().Version?.ToString(3) ?? "2.0.0";
    public string Status { get => _status; private set { _status = value; OnChanged(); } }
    public bool IsBusy { get => _isBusy; private set { _isBusy = value; OnChanged(); } }
    public string ActivityStatus { get => _activityStatus; private set { _activityStatus = value; OnChanged(); } }
    public string DailySyncTime { get => _dailySyncTime; set { _dailySyncTime = value; OnChanged(); } }
    public string ResultSummary { get => _resultSummary; private set { _resultSummary = value; OnChanged(); } }
    public string ScopeSummary { get => _scopeSummary; private set { _scopeSummary = value; OnChanged(); } }
    public string IndexSummary { get => _indexSummary; private set { _indexSummary = value; OnChanged(); } }
    public string DatabasePath => _database.DatabasePath;
    public bool LocalSearchOnly { get => _localSearchOnly; set { _localSearchOnly = value; OnChanged(); } }
    public string ApiKey { get => _apiKey; set { _apiKey = value; OnChanged(); } }
    public DateTime? IndexFromDate { get => _indexFromDate; set { _indexFromDate = value; OnChanged(); } }
    public DateTime? IndexToDate { get => _indexToDate; set { _indexToDate = value; OnChanged(); } }
    public IReadOnlyList<string> ModelOptions { get; } = MailZenModelRegistry.SupportedModelIds;
    public string SelectedModelId { get => _selectedModelId; set { if (MailZenModelRegistry.SupportedModelIds.Contains(value)) { _selectedModelId = value; OnChanged(); } } }
    public IReadOnlyList<SearchModeOption> SearchModeOptions { get; } = new[]
    {
        new SearchModeOption(SearchMode.SmartLocal, "Smart local words (all terms)", "Fast local search. Filler words are removed and every meaningful term must appear in the email."),
        new SearchModeOption(SearchMode.ConversationalAi, "Conversational AI", "Describe what you mean; the selected model interprets the request and searches the indexed messages."),
        new SearchModeOption(SearchMode.Boolean, "Boolean words (AND / OR)", "Use parentheses with AND or &, OR or |, and NOT for explicit local logic.")
    };
    public SearchModeOption SelectedSearchMode { get => _selectedSearchMode; set { if (value is not null) { _selectedSearchMode = value; OnChanged(); } } }
    public double DateRangeMinimum { get => _dateRangeMinimum; private set { _dateRangeMinimum = value; OnChanged(); } }
    public double DateRangeMaximum { get => _dateRangeMaximum; private set { _dateRangeMaximum = value; OnChanged(); } }
    public double SearchStartValue { get => _searchStartValue; set { _searchStartValue = Math.Min(value, SearchEndValue); OnChanged(); OnChanged(nameof(SearchFromDate)); OnChanged(nameof(DateRangeSummary)); } }
    public double SearchEndValue { get => _searchEndValue; set { _searchEndValue = Math.Max(value, SearchStartValue); OnChanged(); OnChanged(nameof(SearchToDate)); OnChanged(nameof(DateRangeSummary)); } }
    public DateTime SearchFromDate => DateTime.FromOADate(SearchStartValue).Date;
    public DateTime SearchToDate => DateTime.FromOADate(SearchEndValue).Date;
    public string DateRangeSummary => $"{SearchFromDate:MMM d, yyyy} – {SearchToDate:MMM d, yyyy}";
    public bool UseAi => !LocalSearchOnly;
    public ICommand SearchCommand { get; }
    public ICommand SyncCommand { get; }
    public ICommand SyncNewCommand { get; }
    public ICommand ScheduleSyncCommand { get; }
    public ICommand RemoveScheduleCommand { get; }
    public ICommand SelectAllCommand { get; }
    public ICommand ClearAllCommand { get; }
    public ICommand SaveKeyCommand { get; }
    public ICommand TestAiCommand { get; }
    public ICommand RemoveKeyCommand { get; }
    public ICommand OpenDatabaseLocationCommand { get; }
    public ICommand OpenResultCommand { get; }
    public ICommand ExportResultCommand { get; }

    public SearchWorkspaceViewModel()
    {
        _selectedSearchMode = SearchModeOptions[0];
        SetDateRange(DateTime.Today.AddYears(-1), DateTime.Today);
        SearchCommand = new AsyncCommand(RunSearchAsync);
        SyncCommand = new AsyncCommand(SyncAsync);
        SyncNewCommand = new AsyncCommand(SyncNewAsync);
        ScheduleSyncCommand = new AsyncCommand(ScheduleSyncAsync);
        RemoveScheduleCommand = new AsyncCommand(async () => { await new WindowsSyncScheduler().RemoveAsync(); Status = "Daily sync schedule removed."; });
        SelectAllCommand = new ActionCommand(() => { foreach (var a in Accounts) a.IsSelected = true; });
        ClearAllCommand = new ActionCommand(() => { foreach (var a in Accounts) a.IsSelected = false; });
        SaveKeyCommand = new AsyncCommand(async () => { await _credentials.SetApiKeyAsync(ApiKey); ApiKey = string.Empty; Status = "API key saved in Windows Credential Manager."; });
        TestAiCommand = new AsyncCommand(TestAiAsync);
        RemoveKeyCommand = new AsyncCommand(async () => { await _credentials.RemoveApiKeyAsync(); Status = "API key removed. Local search remains available."; });
        OpenDatabaseLocationCommand = new ActionCommand(OpenDatabaseLocation);
        OpenResultCommand = new AsyncParameterCommand(async value => { if (value is SearchResultRow row && !await new OutlookOpenService().TryOpenAsync(row.Message)) Status = "Could not open in Outlook; use the .msg fallback."; });
        ExportResultCommand = new AsyncParameterCommand(async value => { if (value is SearchResultRow row) Status = $"Exported to {await new OutlookOpenService().ExportToMsgAsync(row.Message)}"; });
    }

    private void OpenDatabaseLocation()
    {
        try
        {
            var directory = Path.GetDirectoryName(_database.DatabasePath) ?? Environment.CurrentDirectory;
            Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{_database.DatabasePath}\"") { UseShellExecute = true });
            Status = $"Database folder opened: {directory}";
        }
        catch (Exception ex) { Status = $"Could not open database folder: {ex.Message}"; }
    }

    public async Task LoadAccountsAsync()
    {
        try
        {
            foreach (var account in await _outlook.GetAccountsAsync()) Accounts.Add(new SearchAccountRow(account));
            await RefreshAccountCoverageAsync();
            Status = $"{Accounts.Count} Outlook account(s) available.";
            await RefreshCoverageAsync();
        }
        catch (Exception ex) { Status = $"Outlook unavailable: {ex.Message}"; }
    }
    private async Task SyncAsync()
    {
        var ids = Accounts.Where(a => a.IsSelected).Select(a => a.AccountId).ToHashSet();
        if (!IndexFromDate.HasValue || !IndexToDate.HasValue || IndexFromDate > IndexToDate) { Status = "Choose a valid indexing date range."; return; }
        Status = "Syncing selected accounts…";
        var fromUtc = DateTime.SpecifyKind(IndexFromDate.Value.Date, DateTimeKind.Local).ToUniversalTime();
        var toUtc = DateTime.SpecifyKind(IndexToDate.Value.Date.AddDays(1), DateTimeKind.Local).ToUniversalTime();
        await new EmailIndexCoordinator(_outlook, _database).SyncAsync(ids, fromUtc, toUtc, new Progress<string>(s => Status = s));
        Status = $"Sync complete. Indexed Outlook messages from {IndexFromDate:MMM d, yyyy} through {IndexToDate:MMM d, yyyy}.";
        await RefreshAccountCoverageAsync();
        await RefreshCoverageAsync();
    }

    private async Task SyncNewAsync()
    {
        var ids = Accounts.Where(a => a.IsSelected).Select(a => a.AccountId).ToHashSet();
        if (ids.Count == 0) { Status = "Select at least one account to sync."; return; }
        if (IsBusy) return;
        IsBusy = true;
        ActivityStatus = "Syncing new emails\u2026";
        try
        {
            var lastSuccessful = await _database.GetEarliestSuccessfulSyncUtcAsync(ids);
            var sinceUtc = (lastSuccessful ?? DateTime.UtcNow.AddDays(-2)).AddDays(lastSuccessful.HasValue ? -2 : 0);
            Status = lastSuccessful.HasValue
                ? $"Syncing selected accounts from {sinceUtc.ToLocalTime():g}\u2026"
                : "No prior checkpoint found; syncing the last two days for the selected accounts\u2026";
            await new EmailIndexCoordinator(_outlook, _database).SyncAsync(ids, sinceUtc, null,
                new Progress<string>(s => Status = s));
            Status = "New email sync complete.";
            await RefreshAccountCoverageAsync();
            await RefreshCoverageAsync();
        }
        catch (Exception ex) { Status = $"New email sync failed: {ex.Message}"; }
        finally { IsBusy = false; ActivityStatus = string.Empty; }
    }

    private async Task ScheduleSyncAsync()
    {
        if (!TimeOnly.TryParse(DailySyncTime, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out var time))
        {
            Status = "Enter a valid daily time, such as 02:00 or 8:30 PM.";
            return;
        }
        try
        {
            await new WindowsSyncScheduler().ConfigureDailyAsync(time);
            Status = $"Daily sync scheduled for {time:hh\\:mm tt} using MailZen --sync.";
        }
        catch (Exception ex) { Status = $"Could not schedule daily sync: {ex.Message}"; }
    }
    private async Task RunSearchAsync()
    {
        if (IsBusy) return;
        IsBusy = true;
        ActivityStatus = "Search in progress…";
        try { await SearchAsync(); }
        catch (JsonException) { ResultSummary = "AI search failed"; Status = "The AI response was incomplete. Try again or test the AI connection."; }
        catch (HttpRequestException ex) { ResultSummary = "AI connection failed"; Status = $"AI connection error: {ex.Message}"; }
        catch (InvalidOperationException ex) { ResultSummary = "AI setup incomplete"; Status = ex.Message; }
        finally { IsBusy = false; ActivityStatus = string.Empty; }
    }

    private async Task TestAiAsync()
    {
        if (IsBusy) return;
        IsBusy = true;
        ActivityStatus = "Testing AI connection…";
        try
        {
            if (!string.IsNullOrWhiteSpace(ApiKey)) { await _credentials.SetApiKeyAsync(ApiKey); ApiKey = string.Empty; }
            using var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
            var provider = new OpenAiSearchProvider(httpClient, SearchModelProfile.Fast, SelectedModelId);
            var ok = await provider.CheckHealthAsync();
            Status = ok ? $"AI connection succeeded using {SelectedModelId}." : $"AI connection failed for {SelectedModelId}. Check the API key and model.";
        }
        catch (Exception ex) { Status = $"AI connection test failed: {ex.Message}"; }
        finally { IsBusy = false; ActivityStatus = string.Empty; }
    }

    private async Task SearchAsync()
    {
        Results.Clear();
        var ids = Accounts.Where(a => a.IsSelected).Select(a => a.AccountId).ToHashSet();
        await RefreshCoverageAsync(ids);
        if (ids.Count == 0 || string.IsNullOrWhiteSpace(Query)) { ResultSummary = "Select an account and enter a search."; return; }
        var local = new LocalSearchService(_database);
        var afterUtc = DateTime.SpecifyKind(SearchFromDate, DateTimeKind.Local).ToUniversalTime();
        var beforeUtc = DateTime.SpecifyKind(SearchToDate.AddDays(1), DateTimeKind.Local).ToUniversalTime();
        var scope = new SearchScope(ids, ReceivedAfterUtc: afterUtc, ReceivedBeforeUtc: beforeUtc);
        var request = new SearchRequest(Query, scope, 50, SelectedSearchMode.Mode);
        IReadOnlyList<SearchResult> results;
        if (SelectedSearchMode.Mode == SearchMode.ConversationalAi)
        {
            using var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
            var provider = new OpenAiSearchProvider(httpClient, SearchModelProfile.Fast, SelectedModelId);
            results = await new SearchOrchestrator(local, provider).SearchAsync(request, useAi: true);
        }
        else
        {
            results = await local.SearchAsync(request);
        }
        foreach (var result in results) Results.Add(new SearchResultRow { Subject = result.Message.Subject, Sender = $"{result.Message.SenderName} <{result.Message.SenderAddress}>", Account = result.Message.AccountId, Received = result.Message.ReceivedUtc.ToLocalTime().ToString("g"), Excerpt = result.Excerpt.Replace("<b>", "").Replace("</b>", ""), Explanation = result.Explanation ?? "Local match", Message = result.Message });
        ResultSummary = $"{Results.Count} result(s) · {SelectedSearchMode.DisplayName}";
    }
    private Task RefreshCoverageAsync() => RefreshCoverageAsync(Accounts.Where(a => a.IsSelected).Select(a => a.AccountId).ToHashSet());
    private async Task RefreshAccountCoverageAsync()
    {
        var coverage = await _database.GetIndexCoverageByAccountAsync(Accounts.Select(a => a.AccountId).ToHashSet());
        foreach (var account in Accounts)
            account.SetCoverage(coverage.GetValueOrDefault(account.AccountId));
    }
    private async Task RefreshCoverageAsync(IReadOnlySet<string> ids)
    {
        ScopeSummary = ids.Count == 0 ? "No Outlook accounts selected." : $"Searching {ids.Count} selected Outlook account(s).";
        var coverage = await _database.GetIndexCoverageAsync(ids);
        IndexSummary = coverage.MessageCount == 0
            ? "Local index: 0 messages. Click Sync Now to populate the last 30 days."
            : $"Local index: {coverage.MessageCount:N0} messages, {coverage.EarliestReceivedUtc:MMM d, yyyy}–{coverage.LatestReceivedUtc:MMM d, yyyy}.";
    }
    private void SetDateRange(DateTime minimum, DateTime maximum)
    {
        if (maximum < minimum) maximum = minimum;
        DateRangeMinimum = minimum.Date.ToOADate();
        DateRangeMaximum = maximum.Date.ToOADate();
        _searchStartValue = DateRangeMinimum;
        _searchEndValue = DateRangeMaximum;
        OnChanged(nameof(SearchStartValue)); OnChanged(nameof(SearchEndValue));
        OnChanged(nameof(SearchFromDate)); OnChanged(nameof(SearchToDate)); OnChanged(nameof(DateRangeSummary));
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new(name));
}

public sealed class ActionCommand(Action action) : ICommand { public event EventHandler? CanExecuteChanged { add { } remove { } } public bool CanExecute(object? parameter) => true; public void Execute(object? parameter) => action(); }
public sealed record SearchModeOption(SearchMode Mode, string DisplayName, string Description);

public sealed class AsyncCommand(Func<Task> action) : ICommand { public event EventHandler? CanExecuteChanged { add { } remove { } } public bool CanExecute(object? parameter) => true; public async void Execute(object? parameter) => await action(); }
public sealed class AsyncParameterCommand(Func<object?, Task> action) : ICommand { public event EventHandler? CanExecuteChanged { add { } remove { } } public bool CanExecute(object? parameter) => true; public async void Execute(object? parameter) => await action(parameter); }
