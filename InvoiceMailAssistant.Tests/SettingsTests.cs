using InvoiceMailAssistant.App;
using Xunit;

namespace InvoiceMailAssistant.Tests;

public sealed class SettingsTests
{
    [Fact]
    public void MonitorStartOnlyChangesWhenMailboxChangesOrIsUnset()
    {
        var first = new AppSettings();
        var initial = new DateTimeOffset(2026, 8, 8, 1, 0, 0, TimeSpan.Zero);
        Assert.True(AppSettings.EnsureMonitorStart(first, "User@Example.com", initial, out var firstStart));
        Assert.Equal(initial, firstStart);

        var later = initial.AddHours(2);
        Assert.False(AppSettings.EnsureMonitorStart(first, "user@example.com", later, out var unchanged));
        Assert.Equal(initial, unchanged);

        first.ExcelPath = "another.xlsx";
        first.PollSeconds = 120;
        Assert.False(AppSettings.EnsureMonitorStart(first, "user@example.com", later, out var afterOrdinaryChanges));
        Assert.Equal(initial, afterOrdinaryChanges);

        Assert.True(AppSettings.EnsureMonitorStart(first, "other@example.com", later, out var switched));
        Assert.Equal(later, switched);
    }

    [Fact]
    public async Task SaveAsyncWritesValidSettingsWithoutPasswordProperty()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}");
        var path = Path.Combine(directory, "settings.json");
        try
        {
            var settings = new AppSettings { EmailAccount = "user@example.com", ExcelPath = "book.xlsx" };
            await settings.SaveAsync(path);

            var json = await File.ReadAllTextAsync(path);
            Assert.DoesNotContain("password", json, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(settings.EmailAccount, (await AppSettings.LoadAsync(path)).EmailAccount);
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, true);
        }
    }
}
