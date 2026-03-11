namespace EmailManage.Models;

/// <summary>
/// Represents an email preview for the Color Coding Test feature.
/// </summary>
public class ColorTestEmailItem
{
    public string AccountName { get; set; } = "";
    public string StoreId { get; set; } = "";
    public string EntryId { get; set; } = "";
    public string Sender { get; set; } = "";
    public string SubjectSnippet { get; set; } = "";
    public string ReceivedDate { get; set; } = "";
    public string Category { get; set; } = "";
    public string DisplayColor { get; set; } = "#333333";
}
