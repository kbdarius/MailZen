namespace EmailManage.Models;

/// <summary>
/// Result of attempting to connect to Outlook and enumerate accounts.
/// </summary>
public sealed class ConnectionResult
{
    public bool Success { get; init; }
    public string? ErrorMessage { get; init; }
    public string? ErrorCode { get; init; }
    public List<OutlookAccountInfo> Accounts { get; init; } = [];
    public TimeSpan Elapsed { get; init; }

    public static ConnectionResult Ok(List<OutlookAccountInfo> accounts, TimeSpan elapsed) => new()
    {
        Success = true,
        Accounts = accounts,
        Elapsed = elapsed
    };

    public static ConnectionResult Fail(string code, string message, TimeSpan elapsed) => new()
    {
        Success = false,
        ErrorCode = code,
        ErrorMessage = message,
        Elapsed = elapsed
    };
}
