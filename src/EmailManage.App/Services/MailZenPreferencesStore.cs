using System.Text.Json;
using System.IO;

namespace EmailManage.Services;

public sealed class MailZenPreferencesStore
{
    private readonly string _path = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MailZen", "preferences.json");

    public async Task<bool> GetIncludeSentItemsAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            if (!File.Exists(_path)) return false;
            var json = await File.ReadAllTextAsync(_path, cancellationToken);
            return JsonSerializer.Deserialize<MailZenPreferences>(json)?.IncludeSentItems ?? false;
        }
        catch { return false; }
    }

    public async Task SetIncludeSentItemsAsync(bool includeSentItems, CancellationToken cancellationToken = default)
    {
        var directory = Path.GetDirectoryName(_path)!;
        Directory.CreateDirectory(directory);
        var json = JsonSerializer.Serialize(new MailZenPreferences(includeSentItems), new JsonSerializerOptions { WriteIndented = true });
        await File.WriteAllTextAsync(_path, json, cancellationToken);
    }

    private sealed record MailZenPreferences(bool IncludeSentItems);
}
