using System.IO;
using System.Text;
using System.Text.Json;

namespace InvoiceMailAssistant.App;

public sealed class AppSettings
{
    public string EmailAccount { get; set; } = string.Empty;
    public string ImapHost { get; set; } = "imap.exmail.qq.com";
    public int ImapPort { get; set; } = 993;
    public string ExcelPath { get; set; } = string.Empty;
    public string WorksheetName { get; set; } = "中外运";
    public int PollSeconds { get; set; } = 60;
    public string MailboxIdentity { get; set; } = string.Empty;
    public DateTimeOffset? MonitorFromUtc { get; set; }

    public static async Task<AppSettings> LoadAsync(string path)
    {
        if (!File.Exists(path)) return new AppSettings();
        await using var stream = File.OpenRead(path);
        return await JsonSerializer.DeserializeAsync<AppSettings>(stream) ?? new AppSettings();
    }

    public async Task SaveAsync(string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (string.IsNullOrWhiteSpace(directory)) directory = Directory.GetCurrentDirectory();
        Directory.CreateDirectory(directory);

        var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
        _ = JsonSerializer.Deserialize<AppSettings>(json) ?? throw new InvalidDataException("设置序列化校验失败。");
        var tempPath = Path.Combine(directory, $".{Path.GetFileName(path)}.{Guid.NewGuid():N}.tmp");
        try
        {
            await using (var stream = new FileStream(tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.WriteThrough))
            {
                var bytes = Encoding.UTF8.GetBytes(json);
                await stream.WriteAsync(bytes);
                await stream.FlushAsync();
                stream.Flush(true);
            }

            File.Move(tempPath, path, true);
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    public static bool EnsureMonitorStart(AppSettings settings, string account, DateTimeOffset nowUtc, out DateTimeOffset monitorFromUtc)
    {
        var normalizedAccount = account.Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(normalizedAccount))
        {
            monitorFromUtc = nowUtc;
            return false;
        }

        if (!string.Equals(settings.MailboxIdentity, normalizedAccount, StringComparison.OrdinalIgnoreCase) || settings.MonitorFromUtc is null)
        {
            settings.MailboxIdentity = normalizedAccount;
            settings.MonitorFromUtc = nowUtc;
            monitorFromUtc = nowUtc;
            return true;
        }

        monitorFromUtc = settings.MonitorFromUtc.Value;
        return false;
    }
}
