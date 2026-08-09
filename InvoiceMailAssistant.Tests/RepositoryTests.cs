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

    private static InvoiceApplication CreateApplication(string messageId, uint uid, uint uidValidity)
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
            MailboxIdentity = "account@example.com",
            MailReceivedAt = DateTimeOffset.UtcNow,
            MailSubject = "subject",
            MailFrom = "sino-esign@sinotrans.com",
            NormalizedBody = "body",
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
