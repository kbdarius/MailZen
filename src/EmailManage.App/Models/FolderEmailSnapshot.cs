namespace EmailManage.Models;

/// <summary>
/// Lightweight snapshot of one email in a folder — used for before/after comparison
/// when the user reviews emails in Outlook and returns to the tool.
/// </summary>
public class FolderEmailSnapshot
{
    public string EntryId { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string SenderEmail { get; set; } = string.Empty;
    public string Domain { get; set; } = string.Empty;
    public DateTime ReceivedTime { get; set; }
}
