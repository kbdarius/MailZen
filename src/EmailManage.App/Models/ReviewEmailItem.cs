using CommunityToolkit.Mvvm.ComponentModel;

namespace EmailManage.Models;

/// <summary>
/// Wraps an email from "Needs Review" with its score, matched rules, and group info.
/// </summary>
public partial class ReviewEmailItem : ObservableObject
{
    public string EntryId { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string SenderName { get; set; } = string.Empty;
    public string SenderEmail { get; set; } = string.Empty;
    public string Domain { get; set; } = string.Empty;
    public DateTime ReceivedTime { get; set; }
    public double Confidence { get; set; }
    public string ConfidenceBand { get; set; } = string.Empty; // "High", "Medium", "Low"
    public List<string> MatchedRuleIds { get; set; } = new();
    public string MatchedRulesSummary { get; set; } = string.Empty;

    [ObservableProperty]
    private bool _isSelected;

    /// <summary>
    /// For undo: original folder path before action was taken.
    /// </summary>
    public string OriginalFolderName { get; set; } = "Needs Review";
}

/// <summary>
/// Groups review items by domain or category for batch operations.
/// </summary>
public partial class ReviewGroup : ObservableObject
{
    public string GroupKey { get; set; } = string.Empty;
    public string GroupLabel { get; set; } = string.Empty;
    public int ItemCount { get; set; }
    public double AvgConfidence { get; set; }
    public List<ReviewEmailItem> Items { get; set; } = new();

    [ObservableProperty]
    private bool _isSelected;

    partial void OnIsSelectedChanged(bool value)
    {
        foreach (var item in Items)
            item.IsSelected = value;
    }
}

/// <summary>
/// Result from an improve model batch action.
/// </summary>
public class ImproveActionResult
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public string Action { get; set; } = string.Empty; // "Delete+Learn", "Delete Only", "Keep+Learn"
    public int ProcessedCount { get; set; }
    public int NewPatternsAdded { get; set; }
    public int PatternsUpdated { get; set; }
    public List<string> AffectedSenders { get; set; } = new();
}

/// <summary>
/// Tracks a batch for undo support.
/// </summary>
public class UndoBatch
{
    public DateTime Timestamp { get; set; }
    public string Action { get; set; } = string.Empty;
    public List<UndoEntry> Entries { get; set; } = new();
}

public class UndoEntry
{
    public string EntryId { get; set; } = string.Empty;
    public string OriginalFolderName { get; set; } = string.Empty;
    public string DestinationFolderName { get; set; } = string.Empty;
    public string SenderEmail { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
}
