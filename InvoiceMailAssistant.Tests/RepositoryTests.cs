using InvoiceMailAssistant.App;
using Xunit;

namespace InvoiceMailAssistant.Tests;

public sealed class RepositoryTests
{
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
