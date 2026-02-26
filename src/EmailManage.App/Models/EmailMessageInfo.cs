namespace EmailManage.Models;

public class EmailMessageInfo
{
    public string EntryId { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string SenderName { get; set; } = string.Empty;
    public string SenderEmailAddress { get; set; } = string.Empty;
    public string Body { get; set; } = string.Empty;
    public DateTime ReceivedTime { get; set; }
    public bool UnRead { get; set; }
    public bool IsDeleted { get; set; }
}
