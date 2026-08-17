using InvoiceMailAssistant.App;
using Xunit;

namespace InvoiceMailAssistant.Tests;

public sealed class MailboxServiceTests
{
    [Fact]
    public async Task FetchesMoreThanTwoHundredCandidatesWithoutTruncation()
    {
        var monitorFrom = new DateTimeOffset(2026, 8, 8, 0, 0, 0, TimeSpan.Zero);
        var messages = Enumerable.Range(1, 250)
            .Select(index => new MailboxMessage(
                (uint)index,
                7,
                "INBOX",
                monitorFrom.AddMinutes(index),
                ["sino-esign@sinotrans.com"],
                "中外运向您提交了开票申请",
                $"message-{index}",
                $"公司名称：企业{index}"))
            .ToArray();
        var factory = new FakeMailboxSessionFactory(() => new FakeMailboxSession(messages));
        using var service = new MailboxService(factory);

        var result = await service.FetchCandidateMessagesAsync("account@example.com", "password", "imap.example.com", 993, monitorFrom, 200, CancellationToken.None);

        Assert.Equal(250, result.Count);
        Assert.Equal(1u, result[0].Uid);
        Assert.Equal(250u, result[^1].Uid);
        Assert.Equal(1, factory.ConnectCount);
    }

    [Fact]
    public async Task ReusesAuthenticatedSessionBetweenPolls()
    {
        var monitorFrom = DateTimeOffset.UtcNow.AddMinutes(-5);
        var messages = new[] { CreateMessage(1, monitorFrom.AddMinutes(1)) };
        var factory = new FakeMailboxSessionFactory(() => new FakeMailboxSession(messages));
        using var service = new MailboxService(factory);

        await service.FetchCandidateMessagesAsync("account@example.com", "password", "imap.example.com", 993, monitorFrom, 200, CancellationToken.None);
        await service.FetchCandidateMessagesAsync("account@example.com", "password", "imap.example.com", 993, monitorFrom, 200, CancellationToken.None);

        Assert.Equal(1, factory.ConnectCount);
    }

    [Fact]
    public async Task ReconnectsAndRetriesWhenTheExistingImapSessionDrops()
    {
        var monitorFrom = DateTimeOffset.UtcNow.AddMinutes(-5);
        var messages = new[] { CreateMessage(1, monitorFrom.AddMinutes(1)) };
        var first = new FakeMailboxSession(messages, new IOException("IMAP server unexpectedly disconnected"));
        var second = new FakeMailboxSession(messages);
        var sessions = new Queue<IMailboxSession>([first, second]);
        var factory = new FakeMailboxSessionFactory(() => sessions.Dequeue());
        using var service = new MailboxService(factory);

        var result = await service.FetchCandidateMessagesAsync("account@example.com", "password", "imap.example.com", 993, monitorFrom, 200, CancellationToken.None);

        Assert.Single(result);
        Assert.Equal(2, factory.ConnectCount);
        Assert.False(first.IsConnected);

        await service.FetchCandidateMessagesAsync("account@example.com", "password", "imap.example.com", 993, monitorFrom, 200, CancellationToken.None);
        Assert.Equal(2, factory.ConnectCount);
    }

    [Fact]
    public async Task LimitsReconnectToOneAttemptAndBacksOffWhenTheSessionKeepsFailing()
    {
        var monitorFrom = DateTimeOffset.UtcNow.AddMinutes(-5);
        var failure = new IOException("IMAP server unexpectedly disconnected");
        var sessions = new Queue<IMailboxSession>([
            new FakeMailboxSession([], failure),
            new FakeMailboxSession([], failure)
        ]);
        var factory = new FakeMailboxSessionFactory(() => sessions.Dequeue());
        using var service = new MailboxService(factory);

        await Assert.ThrowsAsync<IOException>(() => service.FetchCandidateMessagesAsync("account@example.com", "password", "imap.example.com", 993, monitorFrom, 200, CancellationToken.None));
        Assert.Equal(2, factory.ConnectCount);

        await Assert.ThrowsAsync<MailboxBackoffException>(() => service.FetchCandidateMessagesAsync("account@example.com", "password", "imap.example.com", 993, monitorFrom, 200, CancellationToken.None));
        Assert.Equal(2, factory.ConnectCount);
    }

    [Fact]
    public async Task FailedConnectionEntersBackoffInsteadOfReloggingEveryPoll()
    {
        var factory = new FakeMailboxSessionFactory(() => throw new IOException("network down"));
        using var service = new MailboxService(factory);

        await Assert.ThrowsAsync<IOException>(() => service.FetchCandidateMessagesAsync("account@example.com", "password", "imap.example.com", 993, DateTimeOffset.UtcNow.AddMinutes(-1), 200, CancellationToken.None));
        await Assert.ThrowsAsync<MailboxBackoffException>(() => service.FetchCandidateMessagesAsync("account@example.com", "password", "imap.example.com", 993, DateTimeOffset.UtcNow.AddMinutes(-1), 200, CancellationToken.None));

        Assert.Equal(1, factory.ConnectCount);
    }

    [Fact]
    public async Task RejectsMultipleFromAddressesAndKeepsFetchFailuresForReview()
    {
        var monitorFrom = DateTimeOffset.UtcNow.AddMinutes(-5);
        var messages = new[]
        {
            new MailboxMessage(1, 1, "INBOX", monitorFrom.AddMinutes(1), ["sino-esign@sinotrans.com", "other@example.com"], "中外运向您提交了开票申请", "message-1", "body"),
            new MailboxMessage(2, 1, "INBOX", monitorFrom.AddMinutes(2), ["sino-esign@sinotrans.com"], "中外运向您提交了开票申请", "message-2", string.Empty, "MIME parse failed")
        };
        using var service = new MailboxService(new FakeMailboxSessionFactory(() => new FakeMailboxSession(messages)));

        var result = await service.FetchCandidateMessagesAsync("account@example.com", "password", "imap.example.com", 993, monitorFrom, 200, CancellationToken.None);

        var failure = Assert.Single(result);
        Assert.Equal(2u, failure.Uid);
        Assert.Equal("MIME parse failed", failure.FetchError);
    }

    private static MailboxMessage CreateMessage(uint uid, DateTimeOffset receivedAt)
        => new(uid, 1, "INBOX", receivedAt, ["sino-esign@sinotrans.com"], "中外运向您提交了开票申请", $"message-{uid}", "body");

    private sealed class FakeMailboxSessionFactory(Func<IMailboxSession> create)
        : IMailboxSessionFactory
    {
        public int ConnectCount { get; private set; }

        public Task<IMailboxSession> ConnectAsync(string account, string password, string host, int port, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ConnectCount++;
            return Task.FromResult(create());
        }
    }

    private sealed class FakeMailboxSession(IReadOnlyList<MailboxMessage> messages, Exception? fetchException = null) : IMailboxSession
    {
        public bool IsConnected { get; private set; } = true;

        public Task<IReadOnlyList<MailboxMessage>> FetchCandidateMessagesAsync(DateTimeOffset monitorFromUtc, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (fetchException is not null) throw fetchException;
            return Task.FromResult(messages);
        }

        public ValueTask DisposeAsync()
        {
            IsConnected = false;
            return ValueTask.CompletedTask;
        }
    }
}
