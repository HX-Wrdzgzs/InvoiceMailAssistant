using InvoiceMailAssistant.App;
using Xunit;

namespace InvoiceMailAssistant.Tests;

public sealed class UiLayoutTests
{
    [Fact]
    public async Task SavesAndLoadsIndependentColumnWidthsWithBounds()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"invoice-layout-{Guid.NewGuid():N}");
        var path = Path.Combine(directory, "ui-layout.json");
        try
        {
            using (var store = await UiLayoutStore.LoadAsync(path))
            {
                store.SetWidthDebounced("overview.company", 512, 140, 800);
                store.SetWidthDebounced("pending.error", 1200, 200, 800);
                store.SetWidthDebounced("records.company", 20, 140, 800);
                await store.FlushAsync();
            }

            using var reloaded = await UiLayoutStore.LoadAsync(path);
            Assert.Equal(512, reloaded.GetWidth("overview.company", 260, 140, 800));
            Assert.Equal(800, reloaded.GetWidth("pending.error", 360, 200, 800));
            Assert.Equal(140, reloaded.GetWidth("records.company", 260, 140, 800));
            Assert.Equal(260, reloaded.GetWidth("overview.credit", 260, 140, 800));
            Assert.DoesNotContain("password", await File.ReadAllTextAsync(path), StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }
}
