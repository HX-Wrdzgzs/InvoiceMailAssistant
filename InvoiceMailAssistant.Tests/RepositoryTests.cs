using InvoiceMailAssistant.App;
using Xunit;

namespace InvoiceMailAssistant.Tests;

public sealed class RepositoryTests
{
    [Fact]
    public async Task SourceStatisticsSeparateSinotransAndYuntianhuaAcrossStatuses()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.db");
        try
        {
            var repository = new SqliteInvoiceRepository(path);
            await repository.InitializeAsync();
            var now = DateTimeOffset.Now;

            var sinotrans = CreateApplication("stats-sino", 70, 1);
            sinotrans.Source = InvoiceSource.Sinotrans;
            sinotrans.MailReceivedAt = now.AddMinutes(-5);
            sinotrans.UpdatedAt = now.AddMinutes(-5);
            sinotrans.Id = (await repository.TryInsertAsync(sinotrans, Deduplication.CreateFallbackHash(sinotrans)))!.Value;
            await repository.UpdateStatusAsync(sinotrans.Id, ProcessingStatus.Completed);

            var yuntianhuaPending = CreateApplication("stats-yun-pending", 71, 1);
            yuntianhuaPending.Source = InvoiceSource.Yuntianhua;
            yuntianhuaPending.MailReceivedAt = now.AddMinutes(-4);
            yuntianhuaPending.Id = (await repository.TryInsertAsync(yuntianhuaPending, Deduplication.CreateFallbackHash(yuntianhuaPending)))!.Value;
            await repository.UpdateStatusAsync(yuntianhuaPending.Id, ProcessingStatus.PendingExcel);

            var yuntianhuaFailed = CreateApplication("stats-yun-failed", 72, 1);
            yuntianhuaFailed.Source = InvoiceSource.Yuntianhua;
            yuntianhuaFailed.MailReceivedAt = now.AddMinutes(-3);
            yuntianhuaFailed.ProcessingStatus = ProcessingStatus.ParseFailed;
            yuntianhuaFailed.Id = (await repository.TryInsertAsync(yuntianhuaFailed, Deduplication.CreateFallbackHash(yuntianhuaFailed)))!.Value;

            var stats = await repository.GetSourceStatisticsAsync(now);
            var sinoStats = Assert.Single(stats, x => x.Source == InvoiceSource.Sinotrans);
            var yunStats = Assert.Single(stats, x => x.Source == InvoiceSource.Yuntianhua);

            Assert.Equal(1, sinoStats.TodayReceived);
            Assert.Equal(1, sinoStats.TodayCompleted);
            Assert.Equal(0, sinoStats.Pending);
            Assert.Equal(0, sinoStats.Failed);
            Assert.Equal(2, yunStats.TodayReceived);
            Assert.Equal(0, yunStats.TodayCompleted);
            Assert.Equal(1, yunStats.Pending);
            Assert.Equal(1, yunStats.Failed);
        }
        finally
        {
            DeleteDatabase(path);
        }
    }

    [Fact]
    public async Task ConcurrentUniqueInsertReturnsOneWinnerAndOneDuplicate()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.db");
        try
        {
            var repository = new SqliteInvoiceRepository(path);
            await repository.InitializeAsync();
            var first = CreateApplication("same-message", 21, 7);
            var second = CreateApplication("same-message", 21, 7);
            var hash = Deduplication.CreateFallbackHash(first);

            var results = await Task.WhenAll(
                repository.TryInsertAsync(first, hash),
                repository.TryInsertAsync(second, hash));

            Assert.Single(results, x => x is not null);
            Assert.Single(results, x => x is null);
            Assert.Single(await repository.GetRecentAsync(10));
        }
        finally
        {
            DeleteDatabase(path);
        }
    }

    [Fact]
    public async Task UidValidityAndMailboxArePartOfUidIdentity()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.db");
        try
        {
            var repository = new SqliteInvoiceRepository(path);
            await repository.InitializeAsync();
            var first = CreateApplication("", 22, 100);
            first.UidValidity = 1;
            first.MailboxName = "INBOX";
            var second = CreateApplication("", 22, 101);
            second.UidValidity = 2;
            second.MailboxName = "INBOX";

            Assert.NotNull(await repository.TryInsertAsync(first, "hash-1"));
            Assert.NotNull(await repository.TryInsertAsync(second, "hash-2"));
            Assert.Equal(2, (await repository.GetRecentAsync(10)).Count);
        }
        finally
        {
            DeleteDatabase(path);
        }
    }

    [Fact]
    public async Task SameMessageIdInDifferentMailboxesIsNotDeduplicated()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.db");
        try
        {
            var repository = new SqliteInvoiceRepository(path);
            await repository.InitializeAsync();
            var first = CreateApplication("same-message", 30, 7, "first@example.com");
            var second = CreateApplication("same-message", 30, 7, "second@example.com");

            Assert.NotNull(await repository.TryInsertAsync(first, Deduplication.CreateFallbackHash(first)));
            Assert.NotNull(await repository.TryInsertAsync(second, Deduplication.CreateFallbackHash(second)));
            Assert.Equal(2, (await repository.GetRecentAsync(10)).Count);
        }
        finally
        {
            DeleteDatabase(path);
        }
    }

    [Fact]
    public async Task SameBusinessFieldsWithDifferentMessageBodiesRemainDistinct()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.db");
        try
        {
            var repository = new SqliteInvoiceRepository(path);
            await repository.InitializeAsync();
            var first = CreateApplication("", 31, 7, body: "body-a");
            var second = CreateApplication("", 32, 8, body: "body-b");
            second.MailReceivedAt = first.MailReceivedAt;

            Assert.NotNull(await repository.TryInsertAsync(first, Deduplication.CreateFallbackHash(first)));
            Assert.NotNull(await repository.TryInsertAsync(second, Deduplication.CreateFallbackHash(second)));
            Assert.Equal(2, (await repository.GetRecentAsync(10)).Count);
        }
        finally
        {
            DeleteDatabase(path);
        }
    }

    [Fact]
    public async Task MissingMessageIdWithUidValidityChangeUsesStableFallbackDeduplication()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.db");
        try
        {
            var repository = new SqliteInvoiceRepository(path);
            await repository.InitializeAsync();
            var first = CreateApplication("", 40, 7, body: "same-body");
            var second = CreateApplication("", 99, 8, body: "same-body");
            second.CreditCode = first.CreditCode;
            second.MailReceivedAt = first.MailReceivedAt;

            Assert.NotNull(await repository.TryInsertAsync(first, Deduplication.CreateFallbackHash(first)));
            Assert.Null(await repository.TryInsertAsync(second, Deduplication.CreateFallbackHash(second)));
            Assert.Single(await repository.GetRecentAsync(10));
        }
        finally
        {
            DeleteDatabase(path);
        }
    }

    [Fact]
    public async Task LegacyFallbackHashStillPreventsDuplicateAfterUpgrade()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.db");
        try
        {
            var repository = new SqliteInvoiceRepository(path);
            await repository.InitializeAsync();
            var first = CreateApplication("", 50, 7);
            var second = CreateApplication("", 51, 8);
            second.CreditCode = first.CreditCode;
            second.MailReceivedAt = first.MailReceivedAt;

            Assert.NotNull(await repository.TryInsertAsync(first, Deduplication.CreateLegacyFallbackHash(first)));
            Assert.Null(await repository.TryInsertAsync(second, Deduplication.CreateFallbackHash(second)));
            Assert.Single(await repository.GetRecentAsync(10));
        }
        finally
        {
            DeleteDatabase(path);
        }
    }

    [Fact]
    public async Task LegacyRowsWithoutMailboxIdentityStillPreventUpgradeDuplicates()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.db");
        try
        {
            var repository = new SqliteInvoiceRepository(path);
            await repository.InitializeAsync();
            var legacy = CreateApplication("legacy-message", 60, 9, mailbox: string.Empty);
            var upgraded = CreateApplication("legacy-message", 99, 10, mailbox: "account@example.com");

            Assert.NotNull(await repository.TryInsertAsync(legacy, Deduplication.CreateFallbackHash(legacy)));
            Assert.Null(await repository.TryInsertAsync(upgraded, Deduplication.CreateFallbackHash(upgraded)));
            Assert.Single(await repository.GetRecentAsync(10));
        }
        finally
        {
            DeleteDatabase(path);
        }
    }

    private static InvoiceApplication CreateApplication(string messageId, uint uid, uint uidValidity, string mailbox = "account@example.com", string body = "body")
        => new()
        {
            CompanyName = "测试企业",
            CreditCode = $"CODE-{uid}-{uidValidity}",
            Amount = 100m,
            ApplyTime = new DateTime(2026, 8, 8, 9, 30, 0),
            Email = "finance@example.com",
            MessageId = messageId,
            ImapUid = uid,
            UidValidity = uidValidity,
            MailboxName = "INBOX",
            MailboxIdentity = mailbox,
            MailReceivedAt = new DateTimeOffset(2026, 8, 8, 2, 0, 0, TimeSpan.Zero),
            MailSubject = "subject",
            MailFrom = "sino-esign@sinotrans.com",
            NormalizedBody = body,
            ProcessingStatus = ProcessingStatus.Parsed
        };

    private static void DeleteDatabase(string path)
    {
        foreach (var suffix in new[] { "", "-wal", "-shm" })
        {
            var candidate = path + suffix;
            if (File.Exists(candidate)) File.Delete(candidate);
        }
    }
}
