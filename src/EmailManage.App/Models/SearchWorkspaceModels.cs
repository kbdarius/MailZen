using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace EmailManage.Models;

public sealed class SearchAccountRow : INotifyPropertyChanged
{
    private bool _isSelected = true;
    private string _coverageText = "Not indexed";
    public string AccountId { get; }
    public string DisplayName { get; }
    public string EmailAddress { get; }
    public bool IsSelected { get => _isSelected; set { if (_isSelected != value) { _isSelected = value; PropertyChanged?.Invoke(this, new(nameof(IsSelected))); } } }
    public string CoverageText { get => _coverageText; private set { if (_coverageText != value) { _coverageText = value; PropertyChanged?.Invoke(this, new(nameof(CoverageText))); } } }
    public SearchAccountRow(OutlookAccount account) { AccountId = account.AccountId; DisplayName = account.DisplayName; EmailAddress = account.EmailAddress; }
    public void SetCoverage(IndexCoverage? coverage)
    {
        CoverageText = coverage is null || coverage.MessageCount == 0
            ? "Not indexed"
            : $"{coverage.EarliestReceivedUtc!.Value.ToLocalTime():MMM d, yyyy} – {coverage.LatestReceivedUtc!.Value.ToLocalTime():MMM d, yyyy}\n{coverage.MessageCount:N0} email(s)";
    }
    public event PropertyChangedEventHandler? PropertyChanged;
}

public sealed class SearchResultRow
{
    public string Subject { get; init; } = string.Empty;
    public string Sender { get; init; } = string.Empty;
    public string Account { get; init; } = string.Empty;
    public string Received { get; init; } = string.Empty;
    public string Excerpt { get; init; } = string.Empty;
    public string Explanation { get; init; } = string.Empty;
    public IndexedMessage Message { get; init; } = null!;
}
