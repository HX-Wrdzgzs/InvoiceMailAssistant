using System.IO;
using System.IO.Compression;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Security;
using System.Text;
using System.Xml.Linq;
using ClosedXML.Excel;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using Microsoft.Data.Sqlite;
using Microsoft.Win32;

namespace InvoiceMailAssistant.App;

public sealed class MailboxService : IDisposable
{
    private const string Sender = "sino-esign@sinotrans.com";
    private const string SubjectPrefix = "ä¸­å¤–è¿å‘æ‚¨æäº¤äº†å¼€ç¥¨ç”³è¯·";
    private readonly IMailboxSessionFactory _sessionFactory;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private IMailboxSession? _session;
    private string? _connectionKey;
    private DateTimeOffset _retryAfterUtc;
    private int _failureCount;

    public MailboxService(IMailboxSessionFactory? sessionFactory = null)
        => _sessionFactory = sessionFactory ?? new MailKitMailboxSessionFactory();

    public async Task TestConnectionAsync(string account, string password, string host, int port, CancellationToken cancellationToken)
    {
        await _gate.WaitAsync(cancellationToken);
        try
        {
            _retryAfterUtc = DateTimeOffset.MinValue;
            await DisconnectCoreAsync();
            await ConnectCoreAsync(account, password, host, port, cancellationToken);
        }
        catch (Exception ex) when (IsTransient(ex))
        {
            await DisconnectCoreAsync();
            RegisterFailure();
            throw;
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task<IReadOnlyList<MailEnvelope>> FetchCandidateMessagesAsync(string account, string password, string host, int port, DateTimeOffset monitorFromUtc, int maxMessages, CancellationToken cancellationToken)
    {
        _ = maxMessages; // Kept for source compatibility; the search is intentionally not capped.
        await _gate.WaitAsync(cancellationToken);
        try
        {
            await EnsureConnectedAsync(account, password, host, port, cancellationToken);
            var messages = await _session!.FetchCandidateMessagesAsync(monitorFromUtc, cancellationToken);
            var result = new List<MailEnvelope>(messages.Count);
            foreach (var message in messages)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (message.InternalDate < monitorFromUtc && string.IsNullOrWhiteSpace(message.FetchError))
                    continue;
                var fromAddresses = message.FromAddresses
                    .Select(x => x.Trim().ToLowerInvariant())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .ToArray();
                var from = fromAddresses.Length == 1 ? fromAddresses[0] : string.Empty;
                var subject = message.Subject.Trim();
                var matches = string.Equals(from, Sender, StringComparison.OrdinalIgnoreCase)
                    && subject.StartsWith(SubjectPrefix, StringComparison.Ordinal);
                if (!matches && string.IsNullOrWhiteSpace(message.FetchError)) continue;

                result.Add(new MailEnvelope(
                    message.Uid,
                    message.MessageId,
                    from,
                    subject,
                    message.InternalDate,
                    message.BodyText,
                    message.UidValidity,
                    message.MailboxName,
                    message.FetchError));
            }

            _failureCount = 0;
            _retryAfterUtc = DateTimeOffset.MinValue;
            return result;
        }
        catch (Exception ex) when (IsTransient(ex))
        {
            await DisconnectCoreAsync();
            RegisterFailure();
            throw;
        }
        finally
        {
            _gate.Release();
        }
    }

    private async Task EnsureConnectedAsync(string account, string password, string host, int port, CancellationToken cancellationToken)
    {
        var key = $"{account.Trim().ToLowerInvariant()}|{host.Trim().ToLowerInvariant()}|{port}";
        if (_session?.IsConnected == true && string.Equals(_connectionKey, key, StringComparison.Ordinal))
            return;

        if (_retryAfterUtc > DateTimeOffset.UtcNow)
            throw new MailboxBackoffException($"é‚®ç®±è¿žæŽ¥å¤±è´¥ï¼Œæ­£åœ¨é€€é¿é‡è¯•ï¼ˆä¸‹æ¬¡é‡è¯•æ—¶é—´ï¼š{_retryAfterUtc.ToLocalTime():HH:mm:ss}ï¼‰ã€‚");

        await DisconnectCoreAsync();
        await ConnectCoreAsync(account, password, host, port, cancellationToken);
    }

    private async Task ConnectCoreAsync(string account, string password, string host, int port, CancellationToken cancellationToken)
    {
        var session = await _sessionFactory.ConnectAsync(account, password, host, port, cancellationToken);
        _session = session;
        _connectionKey = $"{account.Trim().ToLowerInvariant()}|{host.Trim().ToLowerInvariant()}|{port}";
        _failureCount = 0;
        _retryAfterUtc = DateTimeOffset.MinValue;
    }

    private async Task DisconnectCoreAsync()
    {
        var session = _session;
        _session = null;
        _connectionKey = null;
        if (session is not null) await session.DisposeAsync();
    }

    private void RegisterFailure()
    {
        var seconds = Math.Min(600, 30 * Math.Pow(2, Math.Min(_failureCount, 5)));
        _failureCount++;
        _retryAfterUtc = DateTimeOffset.UtcNow.AddSeconds(seconds);
    }

    private static bool IsTransient(Exception exception)
        => exception is IOException or SocketException or MailKit.Security.AuthenticationException or SslHandshakeException or MailKit.ServiceNotConnectedException or MailKit.ProtocolException;

    public void Dispose()
    {
        try
        {
            _gate.Wait();
            try { DisconnectCoreAsync().GetAwaiter().GetResult(); }
            catch { }
            finally { _gate.Release(); }
        }
        finally { _gate.Dispose(); }
    }
}

public sealed class MailKitMailboxSessionFactory : IMailboxSessionFactory
{
    public async Task<IMailboxSession> ConnectAsync(string account, string password, string host, int port, CancellationToken cancellationToken)
    {
        var client = new ImapClient { Timeout = 30_000 };
        try
        {
            await client.ConnectAsync(host, port, true, cancellationToken);
            await client.AuthenticateAsync(account, password, cancellationToken);
            return new MailKitMailboxSession(client);
        }
        catch
        {
            client.Dispose();
            throw;
        }
    }
}

public sealed class MailKitMailboxSession(ImapClient client) : IMailboxSession
{
    private const string Sender = "sino-esign@sinotrans.com";
    private const string SubjectPrefix = "ä¸­å¤–è¿å‘æ‚¨æäº¤äº†å¼€ç¥¨ç”³è¯·";

    public bool IsConnected => client.IsConnected && client.IsAuthenticated;

    public async Task<IReadOnlyList<MailboxMessage>> FetchCandidateMessagesAsync(DateTimeOffset monitorFromUtc, CancellationToken cancellationToken)
    {
        var inbox = client.Inbox;
        if (!inbox.IsOpen) await inbox.OpenAsync(FolderAccess.ReadOnly, cancellationToken);

        var imapDateFloor = monitorFromUtc.UtcDateTime.Date.AddDays(-1);
        var query = SearchQuery.FromContains(Sender)
            .And(SearchQuery.SubjectContains(SubjectPrefix))
            .And(SearchQuery.DeliveredAfter(imapDateFloor));
        var uids = await inbox.SearchAsync(query, cancellationToken);
        var summaries = await inbox.FetchAsync(uids, MessageSummaryItems.Envelope | MessageSummaryItems.InternalDate, cancellationToken);
        var result = new List<MailboxMessage>(summaries.Count);
        foreach (var summary in summaries.OrderBy(x => x.UniqueId.Id))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var receivedAt = summary.InternalDate ?? DateTimeOffset.MinValue;
            if (receivedAt < monitorFromUtc) continue;
            try
            {
                var message = await inbox.GetMessageAsync(summary.UniqueId, cancellationToken);
                var from = message.From.Mailboxes.Select(x => x.Address ?? string.Empty).ToArray();
                result.Add(new MailboxMessage(summary.UniqueId.Id, inbox.UidValidity, inbox.FullName, receivedAt, from, message.Subject ?? string.Empty, message.MessageId ?? string.Empty, message.TextBody ?? message.HtmlBody ?? string.Empty));
            }
            catch (MimeKit.ParseException ex)
            {
                result.Add(new MailboxMessage(
                    summary.UniqueId.Id,
                    inbox.UidValidity,
                    inbox.FullName,
                    receivedAt,
                    summary.Envelope?.From.Mailboxes.Select(x => x.Address ?? string.Empty).ToArray() ?? [],
                    summary.Envelope?.Subject ?? string.Empty,
                    summary.Envelope?.MessageId ?? string.Empty,
                    string.Empty,
                    $"é‚®ä»¶ MIME è§£æžå¤±è´¥ï¼š{ex.Message}"));
            }
        }

        return result;
    }

    public async ValueTask DisposeAsync()
    {
        try
        {
            if (client.IsConnected) await client.DisconnectAsync(true);
        }
        finally
        {
            client.Dispose();
        }
    }
}

public sealed class MailboxBackoffException(string message) : IOException(message);

public sealed class SqliteInvoiceRepository(string databasePath)
{
    private readonly string _connectionString = new SqliteConnectionStringBuilder { DataSource = databasePath, Pooling = false, DefaultTimeout = 30 }.ToString();

    public async Task InitializeAsync(CancellationToken cancellationToken = default)
    {
        var directory = Path.GetDirectoryName(databasePath);
        if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
        await using var connection = await OpenConnectionAsync(cancellationToken);
        var command = connection.CreateCommand();
        command.CommandText = "PRAGMA journal_mode=WAL; PRAGMA busy_timeout=10000;";
        await command.ExecuteNonQueryAsync(cancellationToken);
        command = connection.CreateCommand();
        command.CommandText = """
            CREATE TABLE IF NOT EXISTS invoice_applications (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                company_name TEXT NOT NULL,
                credit_code TEXT NOT NULL,
                amount TEXT NOT NULL,
                apply_time TEXT NOT NULL,
                invoice_type TEXT NOT NULL,
                recipient TEXT NOT NULL,
                phone TEXT NOT NULL,
                address TEXT NOT NULL,
                email TEXT NOT NULL,
                remark TEXT NOT NULL,
                message_id TEXT NOT NULL,
                imap_uid INTEGER NOT NULL,
                uid_validity INTEGER NOT NULL DEFAULT 0,
                mailbox_name TEXT NOT NULL DEFAULT 'INBOX',
                mailbox_identity TEXT NOT NULL,
                mail_received_at TEXT NOT NULL,
                mail_subject TEXT NOT NULL,
                mail_from TEXT NOT NULL,
                normalized_body TEXT NOT NULL,
                processing_status INTEGER NOT NULL,
                excel_row INTEGER NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                error_message TEXT NOT NULL,
                fallback_hash TEXT NOT NULL
            );
            CREATE UNIQUE INDEX IF NOT EXISTS ux_invoice_message_id ON invoice_applications(message_id) WHERE message_id <> '';
            CREATE UNIQUE INDEX IF NOT EXISTS ux_invoice_fallback_hash ON invoice_applications(fallback_hash);
            """;
        await command.ExecuteNonQueryAsync(cancellationToken);

        await EnsureColumnAsync(connection, "uid_validity", "INTEGER NOT NULL DEFAULT 0", cancellationToken);
        await EnsureColumnAsync(connection, "mailbox_name", "TEXT NOT NULL DEFAULT 'INBOX'", cancellationToken);
        await EnsureColumnAsync(connection, "mailbox_identity", "TEXT NOT NULL DEFAULT ''", cancellationToken);
        command = connection.CreateCommand();
        command.CommandText = "DROP INDEX IF EXISTS ux_invoice_message_id; DROP INDEX IF EXISTS ux_invoice_fallback_hash; DROP INDEX IF EXISTS ux_invoice_uid; CREATE UNIQUE INDEX IF NOT EXISTS ux_invoice_message_id ON invoice_applications(mailbox_identity, message_id) WHERE message_id <> ''; CREATE UNIQUE INDEX IF NOT EXISTS ux_invoice_fallback_hash ON invoice_applications(mailbox_identity, fallback_hash); CREATE UNIQUE INDEX IF NOT EXISTS ux_invoice_uid ON invoice_applications(mailbox_identity, mailbox_name, uid_validity, imap_uid);";
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<bool> ExistsAsync(string messageId, uint uid, string mailboxIdentity, string fallbackHash, CancellationToken cancellationToken = default)
        => await ExistsAsync(messageId, uid, mailboxIdentity, "INBOX", 0, fallbackHash, cancellationToken);

    public async Task<bool> ExistsAsync(string messageId, uint uid, string mailboxIdentity, string mailboxName, uint uidValidity, string fallbackHash, CancellationToken cancellationToken = default)
        => await ExistsAnyAsync(messageId, uid, mailboxIdentity, mailboxName, uidValidity, fallbackHash, null, cancellationToken);

    private async Task<bool> ExistsAnyAsync(string messageId, uint uid, string mailboxIdentity, string mailboxName, uint uidValidity, string fallbackHash, string? legacyFallbackHash, CancellationToken cancellationToken)
    {
        await using var connection = await OpenConnectionAsync(cancellationToken);
        var command = connection.CreateCommand();
        command.CommandText = "SELECT 1 FROM invoice_applications WHERE (message_id <> '' AND message_id = $messageId AND mailbox_identity = $mailbox) OR (mailbox_identity = $mailbox AND mailbox_name = $mailboxName AND uid_validity = $uidValidity AND imap_uid = $uid) OR (mailbox_identity = $mailbox AND (fallback_hash = $hash OR ($legacyHash IS NOT NULL AND fallback_hash = $legacyHash))) LIMIT 1;";
        command.Parameters.AddWithValue("$messageId", messageId);
        command.Parameters.AddWithValue("$mailbox", mailboxIdentity);
        command.Parameters.AddWithValue("$mailboxName", mailboxName);
        command.Parameters.AddWithValue("$uidValidity", (long)uidValidity);
        command.Parameters.AddWithValue("$uid", (long)uid);
        command.Parameters.AddWithValue("$hash", fallbackHash);
        command.Parameters.AddWithValue("$legacyHash", (object?)legacyFallbackHash ?? DBNull.Value);
        return await command.ExecuteScalarAsync(cancellationToken) is not null;
    }

    public async Task<long?> TryInsertAsync(InvoiceApplication app, string fallbackHash, CancellationToken cancellationToken = default)
    {
        var legacyFallbackHash = app.ProãÏu¶‰žËkºwµç`¹•Ñ¥±•9…µ”¡™¥±•A…Ñ ¥ô¹íÕ¥¹9•ÝÕ¥ ¤é9ô¹ÑµÀ¹á±Íàˆ¤ì((€€€€€€€ÑÉä(€€€€€€€ì(€€€€€€€€€€€¹ÍÕÉ•]É¥Ñ…‰±”¡™¥±•A…Ñ ¤ì(€€€€€€€€€€€Ù…È½É¥¥¹…±]½É­Í¡••ÑA…ÉÑÌ€ôI•…‘AÉ•Í•ÉÙ•‘]½É­Í¡••ÑA…ÉÑÌ¡™¥±•A…Ñ °Ý½É­Í¡••Ñ9…µ”¤ì(€€€€€€€€€€€¥±”¹½Áä¡™¥±•A…Ñ °Ñ•µÁA…Ñ °ÑÉÕ”¤ì(€€€€€€€€€€€Ù…ÈÉ½Ü€ô…ÁÁ±¥…Ñ¥½¸¹á•±I½Ü€üüÑ¡É½Ü¹•Ü%¹Ù…±¥‘=Á•É…Ñ¥½¹á•ÁÑ¥½¸ ‹–Âkšr«’âèá•°ƒ–g–—¦ŠžVgžn»š‚¢†3Žˆ¤ì(€€€€€€€€€€€ÕÍ¥¹œ€¡Ù…ÈÝ½É­‰½½¬€ô¹•Üa1]½É­‰½½¬¡Ñ•µÁA…Ñ ¤¤(€€€€€€€€€€€ì(€€€€€€€€€€€€€€€¥˜€ …Ý½É­‰½½¬¹]½É­Í¡••ÑÌ¹QÉå•Ñ]½É­Í¡••Ð¡Ý½É­Í¡••Ñ9…µ”°½ÕÐÙ…ÈÝ½É­Í¡••Ð¤¤(€€€€€€€€€€€€€€€€€€€Ñ¡É½Ü¹•Ü%¹Ù…±¥‘=Á•É…Ñ¥½¹á•ÁÑ¥½¸ ‹–Þ—’ös¢†£’â7–¶c–r£¾òiíÝ½É­Í¡••Ñ9…µ•ôˆ¤ì((€€€€€€€€€€€€€€€¥˜€¡I½Ý5…Ñ¡•Ì¡Ý½É­Í¡••Ð°É½Ü°…ÁÁ±¥…Ñ¥½¸¤¤(€€€€€€€€€€€€€€€€€€€É•ÑÕÉ¸É½Üì((€€€€€€€€€€€€€€€¥˜€ …I½Ý%ÍµÁÑä¡Ý½É­Í¡••Ð°É½Ü¤¤(€€€€€€€€€€€€€€€ì(€€€€€€€€€€€€€€€€€€€Ù…È•á¥ÍÑ¥¹I½Ü€ô¥¹‘5…Ñ¡¥¹I½Ü¡Ý½É­Í¡••Ð°…ÁÁ±¥…Ñ¥½¸°5…Ñ ¹5…à¡É½Ü°Ý½É­Í¡••Ð¹1…ÍÑI½ÝUÍ• ¤ü¹I½Ý9Õµ‰•È ¤€üüÉ½Ü¤¤ì(€€€€€€€€€€€€€€€€€€€¥˜€¡•á¥ÍÑ¥¹I½Ü¥Ì¹Õ±°¤(€€€€€€€€€€€€€€€€€€€€€€€Ñ¡É½Ü¹•Üá•±I½Ý=ÕÁ¥•‘á•ÁÑ¥½¸ ‰á•°ƒž²°íÉ½Ýôƒ¢†3–ÞË¢Š¯–Û’î[šVÃš6»–6ƒžR£¾ò3¢¾ß¦7šZÃ¢ž–"Kžn»š‚¢†3Žˆ¤ì(€€€€€€€€€€€€€€€€€€€É½Ü€ô•á¥ÍÑ¥¹I½Ü¹Y…±Õ”ì(€€€€€€€€€€€€€€€€€€€…ÁÁ±¥…Ñ¥½¸¹á•±I½Ü€ôÉ½Üì(€€€€€€€€€€€€€€€ô((€€€€€€€€€€€€€€€¥˜€¡I½Ý5…Ñ¡•Ì¡Ý½É­Í¡••Ð°É½Ü°…ÁÁ±¥…Ñ¥½¸¤¤(€€€€€€€€€€€€€€€€€€€É•ÑÕÉ¸É½Üì((€€€€€€€€€€€€€€€¥˜€¡É½Ü€ø€È¤½ÁåI½ÝMÑå±”¡Ý½É­Í¡••Ð°É½Ü€´€Ä°É½Ü¤ì((€€€€€€€€€€€€€€€Ù…ÈÁÉ•Ù¥½ÕÍ…Ñ”€ô¥¹‘AÉ•Ù¥½ÕÍÁÁ±¥…Ñ¥½¹…Ñ”¡Ý½É­Í¡••Ð°É½Ü€´€Ä°…ÁÁ±¥…Ñ¥½¸¹ÁÁ±åQ¥µ”¹e•…È¤ì(€€€€€€€€€€€€€€€¥˜€¡ÁÉ•Ù¥½ÕÍ…Ñ”ü¹…Ñ”€„ô…ÁÁ±¥…Ñ¥½¸¹ÁÁ±åQ¥µ”¹…Ñ”¤(€€€€€€€€€€€€€€€€€€€Ý½É­Í¡••Ð¹•±°¡É½Ü°€Ä¤¹Y…±Õ”€ô€‰í…ÁÁ±¥…Ñ¥½¸¹ÁÁ±åQ¥µ”¹5½¹Ñ¡ô¹í…ÁÁ±¥…Ñ¥½¸¹ÁÁ±åQ¥µ”¹…åôˆì(€€€€€€€€€€€€€€€•±Í”(€€€€€€€€€€€€€€€€€€€Ý½É­Í¡••Ð¹•±°¡É½Ü°€Ä¤¹±•…È¡a1±•…É=ÁÑ¥½¹Ì¹½¹Ñ•¹ÑÌ¤ì((€€€€€€€€€€€€€€€Ý½É­Í¡••Ð¹•±°¡É½Ü°€È¤¹Y…±Õ”€ô…ÁÁ±¥…Ñ¥½¸¹½µÁ…¹å9…µ”ì(€€€€€€€€€€€€€€€Ý½É­Í¡••Ð¹•±°¡É½Ü°€Ì¤¹M•ÑY…±Õ”¡…ÁÁ±¥…Ñ¥½¸¹É•‘¥Ñ½‘”¤ì(€€€€€€€€€€€€€€€Ý½É­Í¡••Ð¹•±°¡É½Ü°€Ì¤¹MÑå±”¹9Õµ‰•É½Éµ…Ð¹½Éµ…Ð€ô€‰ ˆì(€€€€€€€€€€€€€€€Ý½É­Í¡••Ð¹•±°¡É½Ü°€Ð¤¹Y…±Õ”€ô…ÁÁ±¥…Ñ¥½¸¹µ½Õ¹Ðì(€€€€€€€€€€€€€€€Ý½É­Í¡••Ð¹•±°¡É½Ü°€Ø¤¹M•ÑY…±Õ”¡…ÁÁ±¥…Ñ¥½¸¹µ…¥°¤ì(€€€€€€€€€€€€€€€Ý½É­Í¡••Ð¹•±°¡É½Ü°€Ø¤¹MÑå±”¹9Õµ‰•É½Éµ…Ð¹½Éµ…Ð€ô€‰ ˆì(€€€€€€€€€€€€€€€Ý½É­‰½½¬¹M…Ù•Ì¡Ñ•µÁA…Ñ ¤ì(€€€€€€€€€€€ô((€€€€€€€€€€€I•ÍÑ½É•AÉ•Í•ÉÙ•‘]½É­Í¡••ÑA…ÉÑÌ¡Ñ•µÁA…Ñ °Ý½É­Í¡••Ñ9…µ”°½É¥¥¹…±]½É­Í¡••ÑA…ÉÑÌ¤ì((€€€€€€€€€€€ÕÍ¥¹œ€¡Ù…ÈÙ•É¥™ä€ô¹•Üa1]½É­‰½½¬¡Ñ•µÁA…Ñ ¤¤(€€€€€€€€€€€ì(€€€€€€€€€€€€€€€Ù…ÈÙ•É¥™åM¡••Ð€ôÙ•É¥™ä¹]½É­Í¡••Ð¡Ý½É­Í¡••Ñ9…µ”¤ì(€€€€€€€€€€€€€€€¥˜€ …I½Ý5…Ñ¡•Ì¡Ù•É¥™åM¡••Ð°É½Ü°…ÁÁ±¥…Ñ¥½¸¤¤(€€€€€€€€€€€€€€€€€€€Ñ¡É½Ü¹•Ü%¹Ù…±¥‘…Ñ…á•ÁÑ¥½¸ ‰á•°ƒ’âÓš^ÛšZ’îÛš‚‡¦ª3–’Ç¢Ò—¾òkžn»š‚¢†3––ºç’â7’â¢ÓŽˆ¤ì(€€€€€€€€€€€ô((€€€€€€€€€€€ÕÍ¥¹œ€¡Ù…È‘ÕÉ…‰±”€ô¹•Ü¥±•MÑÉ•…´¡Ñ•µÁA…Ñ °¥±•5½‘”¹=Á•¸°¥±••ÍÌ¹I•…‘]É¥Ñ”°¥±•M¡…É”¹I•…¤¤(€€€€€€€€€€€€€€€‘ÕÉ…‰±”¹±ÕÍ ¡ÑÉÕ”¤ì((€€€€€€€€€€€¥±”¹I•Á±…”¡Ñ•µÁA…Ñ °™¥±•A…Ñ °¹Õ±°°ÑÉÕ”¤ì(€€€€€€€€€€€É•ÑÕÉ¸É½Üì(€€€€€€€ô(€€€€€€€™¥¹…±±ä(€€€€€€€ì(€€€€€€€€€€€¥˜€¡¥±”¹á¥ÍÑÌ¡Ñ•µÁA…Ñ ¤¤¥±”¹•±•Ñ”¡Ñ•µÁA…Ñ ¤ì(€€€€€€€€€€€½Ý¹•‘1½¬ü¹¥ÍÁ½Í” ¤ì(€€€€€€€ô(€€€ô((€€€ÁÉ¥Ù…Ñ”ÍÑ…Ñ¥ŒÙ½¥¹ÍÕÉ•]É¥Ñ…‰±”¡ÍÑÉ¥¹œ™¥±•A…Ñ ¤(€€€ì(€€€€€€€¥˜€ …¥±”¹á¥ÍÑÌ¡™¥±•A…Ñ ¤¤Ñ¡É½Ü¹•Ü¥±•9½Ñ½Õ¹‘á•ÁÑ¥½¸ ‰á•°ƒžfï¢ºÃ¢†£’â7–¶c–r£Žˆ°™¥±•A…Ñ ¤ì(€€€€€€€ÑÉä(€€€€€€€ì(€€€€€€€€€€€ÕÍ¥¹œÙ…ÈÍÑÉ•…´€ô¹•Ü¥±•MÑÉ•…´¡™¥±•A…Ñ °¥±•5½‘”¹=Á•¸°¥±••ÍÌ¹I•…‘]É¥Ñ”°¥±•M¡…É”¹9½¹”¤ì(€€€€€€€ô(€€€€€€€…Ñ €¡%=á•ÁÑ¥½¸•à¤(€€€€€€€ì(€€€€€€€€€€€Ñ¡É½Ü¹•Ü%=á•ÁÑ¥½¸ ‰á•°ƒžfï¢ºÃ¢†£š¶–r£¢Š¯–6ƒžR£¾ò3žRÏ¢¾ß–Â’þwžVg–r£ž¶'–ú¦b–"_Žˆ°•à¤ì(€€€€€€€ô(€€€ô((€€€ÁÉ¥Ù…Ñ”ÍÑ…Ñ¥Œ‰½½°I½Ý%ÍµÁÑä¡%a1]½É­Í¡••ÐÝ½É­Í¡••Ð°¥¹ÐÉ½Ü¤(€€€ì(€€€€€€€Ù…È±…ÍÑ½±Õµ¸€ô5…Ñ ¹5…à à°Ý½É­Í¡••Ð¹1…ÍÑ½±Õµ¹UÍ• ¤ü¹½±Õµ¹9Õµ‰•È ¤€üü€à¤ì(€€€€€€€É•ÑÕÉ¸¹Õµ•É…‰±”¹I…¹” Ä°±…ÍÑ½±Õµ¸¤¹±°¡½±Õµ¸€ôøÝ½É­Í¡••Ð¹•±°¡É½Ü°½±Õµ¸¤¹%ÍµÁÑä ¤¤ì(€€€ô((€€€ÁÉ¥Ù…Ñ”ÍÑ…Ñ¥Œ¥¹Ðü¥¹‘5…Ñ¡¥¹I½Ü¡%a1]½É­Í¡••ÐÝ½É­Í¡••Ð°%¹Ù½¥•ÁÁ±¥…Ñ¥½¸…ÁÁ±¥…Ñ¥½¸°¥¹Ð±…ÍÑI½Ü¤(€€€ì(€€€€€€€¥¹Ðüµ…Ñ €ô¹Õ±°ì(€€€€€€€™½È€¡Ù…ÈÉ½Ü€ô€ÈìÉ½Ü€ðô±…ÍÑI½ÜìÉ½Ü¬¬¤(€€€€€€€ì(€€€€€€€€€€€¥˜€ …I½Ý5…Ñ¡•Ì¡Ý½É­Í¡••Ð°É½Ü°…ÁÁ±¥…Ñ¥½¸¤¤½¹Ñ¥¹Õ”ì(€€€€€€€€€€€¥˜€¡µ…Ñ ¥Ì¹½Ð¹Õ±°¤É•ÑÕÉ¸¹Õ±°ì(€€€€€€€€€€€µ…Ñ €ôÉ½Üì(€€€€€€€ô(€€€€€€€É•ÑÕÉ¸µ…Ñ ì(€€€ô((€€€ÁÉ¥Ù…Ñ”ÍÑ…Ñ¥Œ‰½½°I½Ý5…Ñ¡•Ì¡%a1]½É­Í¡••ÐÝ½É­Í¡••Ð°¥¹ÐÉ½Ü°%¹Ù½¥•ÁÁ±¥…Ñ¥½¸…ÁÁ±¥…Ñ¥½¸¤(€€€ì(€€€€€€€¥˜€ …ÍÑÉ¥¹œ¹ÅÕ…±Ì¡Ý½É­Í¡••Ð¹•±°¡É½Ü°€È¤¹•ÑMÑÉ¥¹œ ¤¹QÉ¥´ ¤°…ÁÁ±¥…Ñ¥½¸¹½µÁ…¹å9…µ”¹QÉ¥´ ¤°MÑÉ¥¹½µÁ…É¥Í½¸¹=É‘¥¹…°¤¤É•ÑÕÉ¸™…±Í”ì(€€€€€€€¥˜€ …ÍÑÉ¥¹œ¹ÅÕ…±Ì¡Ý½É­Í¡••Ð¹•±°¡É½Ü°€Ì¤¹•ÑMÑÉ¥¹œ ¤¹QÉ¥´ ¤°…ÁÁ±¥…Ñ¥½¸¹É•‘¥Ñ½‘”¹QÉ¥´ ¤°MÑÉ¥¹½µÁ…É¥Í½¸¹=É‘¥¹…°¤¤É•ÑÕÉ¸™…±Í”ì(€€€€€€€¥˜€ …‘•¥µ…°¹QÉåA…ÉÍ”¡Ý½É­Í¡••Ð¹•±°¡É½Ü°€Ð¤¹•ÑMÑÉ¥¹œ ¤°½ÕÐÙ…È…µ½Õ¹Ð¤ñð…µ½Õ¹Ð€„ô…ÁÁ±¥…Ñ¥½¸¹µ½Õ¹Ð¤É•ÑÕÉ¸™…±Í”ì(€€€€€€€É•ÑÕÉ¸ÍÑÉ¥¹œ¹ÅÕ…±Ì¡Ý½É­Í¡••Ð¹•±°¡É½Ü°€Ø¤¹•ÑMÑÉ¥¹œ ¤¹QÉ¥´ ¤°…ÁÁ±¥…Ñ¥½¸¹µ…¥°¹QÉ¥´ ¤°MÑÉ¥¹½µÁ…É¥Í½¸¹=É‘¥¹…±%¹½É•…Í”¤ì(€€€ô((€€€ÁÉ¥Ù…Ñ”ÍÑ…Ñ¥ŒÙ½¥½ÁåI½ÝMÑå±”¡%a1]½É­Í¡••ÐÝÌ°¥¹ÐÍ½ÕÉ•I½Ü°¥¹ÐÑ…É•ÑI½Ü¤(€€€ì(€€€€€€€™½È€¡Ù…È½°€ô€Äì½°€ðô€àì½°¬¬¤(€€€€€€€€€€€ÝÌ¹•±°¡Ñ…É•ÑI½Ü°½°¤¹MÑå±”€ôÝÌ¹•±°¡Í½ÕÉ•I½Ü°½°¤¹MÑå±”ì(€€€€€€€ÝÌ¹I½Ü¡Ñ…É•ÑI½Ü¤¹!•¥¡Ð€ôÝÌ¹I½Ü¡Í½ÕÉ•I½Ü¤¹!•¥¡Ðì(€€€ô((€€€ÁÉ¥Ù…Ñ”ÍÑ…Ñ¥Œ…Ñ•Q¥µ”ü¥¹‘AÉ•Ù¥½ÕÍÁÁ±¥…Ñ¥½¹…Ñ”¡%a1]½É­Í¡••ÐÝÌ°¥¹ÐÉ½Ü°¥¹Ð…ÍÍÕµ•‘e•…È¤(€€€ì(€€€€€€€™½È€¡Ù…ÈÕÉÉ•¹Ð€ôÉ½ÜìÕÉÉ•¹Ð€øô€ÈìÕÉÉ•¹Ð´´¤(€€€€€€€ì(€€€€€€€€€€€Ù…ÈÑ•áÐ€ôÝÌ¹•±°¡ÕÉÉ•¹Ð°€Ä¤¹•ÑMÑÉ¥¹œ ¤¹QÉ¥´ ¤ì(€€€€€€€€€€€¥˜€¡ÍÑÉ¥¹œ¹%Í9Õ±±=É]¡¥Ñ•MÁ…”¡Ñ•áÐ¤¤½¹Ñ¥¹Õ”ì(€€€€€€€€€€€Ù…ÈÁ…ÉÑÌ€ôÑ•áÐ¹MÁ±¥Ð œ¸œ¤ì(€€€€€€€€€€€¥˜€¡Á…ÉÑÌ¹1•¹Ñ €ôô€È€˜˜¥¹Ð¹QÉåA…ÉÍ”¡Á…ÉÑÍlÁt°½ÕÐÙ…Èµ½¹Ñ ¤€˜˜¥¹Ð¹QÉåA…ÉÍ”¡Á…ÉÑÍlÅt°½ÕÐÙ…È‘…ä¤¤(€€€€€€€€€€€ì(€€€€€€€€€€€€€€€ÑÉäìÉ•ÑÕÉ¸¹•Ü…Ñ•Q¥µ”¡…ÍÍÕµ•‘e•…È°µ½¹Ñ °‘…ä¤ìô(€€€€€€€€€€€€€€€…Ñ €¡ÉÕµ•¹Ñ=ÕÑ=™I…¹•á•ÁÑ¥½¸¤ìÉ•ÑÕÉ¸¹Õ±°ìô(€€€€€€€€€€€ô(€€€€€€€ô(€€€€€€€É•ÑÕÉ¸¹Õ±°ì(€€€ô((€€€ÁÉ¥Ù…Ñ”Í•…±•É•½ÉAÉ•Í•ÉÙ•‘]½É­Í¡••ÑA…ÉÑÌ (€€€€€€€a±•µ•¹ÐüM¡••ÑAÉ½Á•ÉÑ¥•Ì°(€€€€€€€a±•µ•¹ÐüM¡••Ñ½Éµ…ÑAÉ½Á•ÉÑ¥•Ì°(€€€€€€€a±•µ•¹ÐüÉ••é•A…¹”°(€€€€€€€a±•µ•¹ÐüAÉ¥¹Ñ=ÁÑ¥½¹Ì°(€€€€€€€a±•µ•¹ÐüA…•5…É¥¹Ì°(€€€€€€€a±•µ•¹ÐüA…•M•ÑÕÀ°(€€€€€€€a±•µ•¹Ðü!•…‘•É½½Ñ•È°(€€€€€€€a±•µ•¹ÐüI½Ý	É•…­Ì°(€€€€€€€a±•µ•¹Ðü½±Õµ¹	É•…­Ì¤ì((€€€ÁÉ¥Ù…Ñ”ÍÑ…Ñ¥ŒAÉ•Í•ÉÙ•‘]½É­Í¡••ÑA…ÉÑÌI•…‘AÉ•Í•ÉÙ•‘]½É­Í¡••ÑA…ÉÑÌ¡ÍÑÉ¥¹œ™¥±•A…Ñ °ÍÑÉ¥¹œÝ½É­Í¡••Ñ9…µ”¤(€€€ì(€€€€€€€ÕÍ¥¹œÙ…È…É¡¥Ù”€ôi¥Á¥±”¹=Á•¹I•…¡™¥±•A…Ñ ¤ì(€€€€€€€Ù…ÈÝ½É­Í¡••ÑA…Ñ €ôI•Í½±Ù•]½É­Í¡••ÑA…Ñ ¡…É¡¥Ù”°Ý½É­Í¡••Ñ9…µ”¤ì(€€€€€€€ÕÍ¥¹œÙ…ÈÍÑÉ•…´€ô…É¡¥Ù”¹•Ñ¹ÑÉä¡Ý½É­Í¡••ÑA…Ñ ¤ü¹=Á•¸ ¤ì(€€€€€€€¥˜€¡ÍÑÉ•…´¥Ì¹Õ±°¤Ñ¡É½Ü¹•Ü%¹Ù…±¥‘…Ñ…á•ÁÑ¥½¸ ‹–Þ—’ös¢† a50ƒ’â7–¶c–r£¾òiíÝ½É­Í¡••Ñ9…µ•ôˆ¤ì(€€€€€€€Ù…È‘½Õµ•¹Ð€ôa½Õµ•¹Ð¹1½…¡ÍÑÉ•…´¤ì(€€€€€€€a9…µ•ÍÁ…”ÍÁÉ•…‘Í¡••Ð€ô€‰¡ÑÑÀè¼½Í¡•µ…Ì¹½Á•¹áµ±™½Éµ…ÑÌ¹½Éœ½ÍÁÉ•…‘Í¡••Ñµ°¼ÈÀÀØ½µ…¥¸ˆì(€€€€€€€Ù…ÈÉ½½Ð€ô‘½Õµ•¹Ð¹I½½Ð€üüÑ¡É½Ü¹•Ü%¹Ù…±¥‘…Ñ…á•ÁÑ¥½¸ ‹–Þ—’ös¢† a50ƒš^ƒš‚ç¢*ž
ç¾òiíÝ½É­Í¡••Ñ9…µ•ôˆ¤ì(€€€€€€€Ù…ÈÍ¡••ÑY¥•Ü€ôÉ½½Ð¹±•µ•¹Ð¡ÍÁÉ•…‘Í¡••Ð€¬€‰Í¡••ÑY¥•ÝÌˆ¤ü¹±•µ•¹ÑÌ¡ÍÁÉ•…‘Í¡••Ð€¬€‰Í¡••ÑY¥•Üˆ¤¹¥ÉÍÑ=É•™…Õ±Ð ¤ì(€€€€€€€É•ÑÕÉ¸¹•ÜAÉ•Í•ÉÙ•‘]½É­Í¡••ÑA…ÉÑÌ (€€€€€€€€€€€±½¹•±•µ•¹Ð¡É½½Ð¹±•µ•¹Ð¡ÍÁÉ•…‘Í¡••Ð€¬€‰Í¡••ÑAÈˆ¤¤°(€€€€€€€€€€€±½¹•±•µ•¹Ð¡É½½Ð¹±•µ•¹Ð¡ÍÁÉ•…‘Í¡••Ð€¬€‰Í¡••Ñ½Éµ…ÑAÈˆ¤¤°(€€€€€€€€€€€±½¹•±•µ•¹Ð¡Í¡••ÑY¥•Üü¹±•µ•¹Ð¡ÍÁÉ•…‘Í¡••Ð€¬€‰Á…¹”ˆ¤¤°(€€€€€€€€€€€±½¹•±•µ•¹Ð¡É½½Ð¹±•µ•¹Ð¡ÍÁÉ•…‘Í¡••Ð€¬€‰ÁÉ¥¹Ñ=ÁÑ¥½¹Ìˆ¤¤°(€€€€€€€€€€€±½¹•±•µ•¹Ð¡É½½Ð¹±•µ•¹Ð¡ÍÁÉ•…‘Í¡••Ð€¬€‰Á…•5…É¥¹Ìˆ¤¤°(€€€€€€€€€€€±½¹•±•µ•¹Ð¡É½½Ð¹±•µ•¹Ð¡ÍÁÉ•…‘Í¡••Ð€¬€‰Á…•M•ÑÕÀˆ¤¤°(€€€€€€€€€€€±½¹•±•µ•¹Ð¡É½½Ð¹±•µ•¹Ð¡ÍÁÉ•…‘Í¡••Ð€¬€‰¡•…‘•É½½Ñ•Èˆ¤¤°(€€€€€€€€€€€±½¹•±•µ•¹Ð¡É½½Ð¹±•µ•¹Ð¡ÍÁÉ•…‘Í¡••Ð€¬€‰É½Ý	É•…­Ìˆ¤¤°(€€€€€€€€€€€±½¹•±•µ•¹Ð¡É½½Ð¹±•µ•¹Ð¡ÍÁÉ•…‘Í¡••Ð€¬€‰½±	É•…­Ìˆ¤¤¤ì(€€€ô((€€€ÁÉ¥Ù…Ñ”ÍÑ…Ñ¥ŒÙ½¥I•ÍÑ½É•AÉ•Í•ÉÙ•‘]½É­Í¡••ÑA…ÉÑÌ¡ÍÑÉ¥¹œ™¥±•A…Ñ °ÍÑÉ¥¹œÝ½É­Í¡••Ñ9…µ”°AÉ•Í•ÉÙ•‘]½É­Í¡••ÑA…ÉÑÌ½É¥¥¹…±A…ÉÑÌ¤(€€€ì(€€€€€€€ÕÍ¥¹œÙ…È…É¡¥Ù”€ôi¥Á¥±”¹=Á•¸¡™¥±•A…Ñ °i¥ÁÉ¡¥Ù•5½‘”¹UÁ‘…Ñ”¤ì(€€€€€€€Ù…ÈÝ½É­Í¡••ÑA…Ñ €ôI•Í½±Ù•]½É­Í¡••ÑA…Ñ ¡…É¡¥Ù”°Ý½É­Í¡••Ñ9…µ”¤ì(€€€€€€€Ù…È•¹ÑÉä€ô…É¡¥Ù”¹•Ñ¹ÑÉä¡Ý½É­Í¡••ÑA…Ñ ¤€üüÑ¡É½Ü¹•Ü%¹Ù…±¥‘…Ñ…á•ÁÑ¥½¸ ‹–Þ—’ös¢† a50ƒ’â7–¶c–r£¾òiíÝ½É­Í¡••Ñ9…µ•ôˆ¤ì(€€€€€€€a½Õµ•¹Ð‘½Õµ•¹Ðì(€€€€€€€ÕÍ¥¹œ€¡Ù…ÈÍÑÉ•…´€ô•¹ÑÉä¹=Á•¸ ¤¤‘½Õµ•¹Ð€ôa½Õµ•¹Ð¹1½…¡ÍÑÉ•…´¤ì((€€€€€€€a9…µ•ÍÁ…”ÍÁÉ•…‘Í¡••Ð€ô€‰¡ÑÑÀè¼½Í¡•µ…Ì¹½Á•¹áµ±™½Éµ…ÑÌ¹½Éœ½ÍÁÉ•…‘Í¡••Ñµ°¼ÈÀÀØ½µ…¥¸ˆì(€€€€€€€Ù…ÈÉ½½Ð€ô‘½Õµ•¹Ð¹I½½Ð€üüÑ¡É½Ü¹•Ü%¹Ù…±¥‘…Ñ…á•ÁÑ¥½¸ ‹–Þ—’ös¢† a50ƒš^ƒš‚ç¢*ž
ç¾òiíÝ½É­Í¡••Ñ9…µ•ôˆ¤ì(€€€€€€€I•Á±…•¡¥±¡É½½Ð°ÍÁÉ•…‘Í¡••Ð€¬€‰Í¡••ÑAÈˆ°½É¥¥¹…±A…ÉÑÌ¹M¡••ÑAÉ½Á•ÉÑ¥•Ì¤ì(€€€€€€€I•Á±…•¡¥±¡É½½Ð°ÍÁÉ•…‘Í¡••Ð€¬€‰Í¡••Ñ½Éµ…ÑAÈˆ°½É¥¥¹…±A…ÉÑÌ¹M¡••Ñ½Éµ…ÑAÉ½Á•ÉÑ¥•Ì¤ì(€€€€€€€Ù…ÈÍ¡••ÑY¥•Ü€ô‘½Õµ•¹Ð¹I½½Ðü¹±•µ•¹Ð¡ÍÁÉ•…‘Í¡••Ð€¬€‰Í¡••ÑY¥•ÝÌˆ¤ü¹±•µ•¹ÑÌ¡ÍÁÉ•…‘Í¡••Ð€¬€‰Í¡••ÑY¥•Üˆ¤¹¥ÉÍÑ=É•™…Õ±Ð ¤ì(€€€€€€€¥˜€¡Í¡••ÑY¥•Ü¥Ì¹½Ð¹Õ±°¤(€€€€€€€ì(€€€€€€€€€€€I•Á±…•¡¥±¡Í¡••ÑY¥•Ü°ÍÁÉ•…‘Í¡••Ð€¬€‰Á…¹”ˆ°½É¥¥¹…±A…ÉÑÌ¹É••é•A…¹”¤ì(€€€€€€€ô((€€€€€€€I•Á±…•¡¥±¡É½½Ð°ÍÁÉ•…‘Í¡••Ð€¬€‰ÁÉ¥¹Ñ=ÁÑ¥½¹Ìˆ°½É¥¥¹…±A…ÉÑÌ¹AÉ¥¹Ñ=ÁÑ¥½¹Ì¤ì(€€€€€€€I•Á±…•¡¥±¡É½½Ð°ÍÁÉ•…‘Í¡••Ð€¬€‰Á…•5…É¥¹Ìˆ°½É¥¥¹…±A…ÉÑÌ¹A…•5…É¥¹Ì¤ì(€€€€€€€I•Á±…•¡¥±¡É½½Ð°ÍÁÉ•…‘Í¡••Ð€¬€‰Á…•M•ÑÕÀˆ°½É¥¥¹…±A…ÉÑÌ¹A…•M•ÑÕÀ¤ì(€€€€€€€I•Á±…•¡¥±¡É½½Ð°ÍÁÉ•…‘Í¡••Ð€¬€‰¡•…‘•É½½Ñ•Èˆ°½É¥¥¹…±A…ÉÑÌ¹!•…‘•É½½Ñ•È¤ì(€€€€€€€I•Á±…•¡¥±¡É½½Ð°ÍÁÉ•…‘Í¡••Ð€¬€‰É½Ý	É•…­Ìˆ°½É¥¥¹…±A…ÉÑÌ¹I½Ý	É•…­Ì¤ì(€€€€€€€I•Á±…•¡¥±¡É½½Ð°ÍÁÉ•…‘Í¡••Ð€¬€‰½±	É•…­Ìˆ°½É¥¥¹…±A…ÉÑÌ¹½±Õµ¹	É•…­Ì¤ì((€€€€€€€•¹ÑÉä¹•±•Ñ” ¤ì(€€€€€€€Ù…ÈÉ•Á±…•µ•¹Ð€ô…É¡¥Ù”¹É•…Ñ•¹ÑÉä¡Ý½É­Í¡••ÑA…Ñ °½µÁÉ•ÍÍ¥½¹1•Ù•°¹=ÁÑ¥µ…°¤ì(€€€€€€€ÕÍ¥¹œÙ…È½ÕÑÁÕÐ€ôÉ•Á±…•µ•¹Ð¹=Á•¸ ¤ì(€€€€€€€‘½Õµ•¹Ð¹M…Ù”¡½ÕÑÁÕÐ°MåÍÑ•´¹aµ°¹1¥¹Ä¹M…Ù•=ÁÑ¥½¹Ì¹¥Í…‰±•½Éµ…ÑÑ¥¹œ¤ì(€€€ô((€€€ÁÉ¥Ù…Ñ”ÍÑ…Ñ¥Œa±•µ•¹Ðü±½¹•±•µ•¹Ð¡a±•µ•¹Ðü•±•µ•¹Ð¤€ôø•±•µ•¹Ð¥Ì¹Õ±°€ü¹Õ±°€è¹•Üa±•µ•¹Ð¡•±•µ•¹Ð¤ì((€€€ÁÉ¥Ù…Ñ”ÍÑ…Ñ¥ŒÙ½¥I•Á±…•¡¥±¡a±•µ•¹ÐÁ…É•¹Ð°a9…µ”¹…µ”°a±•µ•¹Ðü½É¥¥¹…°¤(€€€ì(€€€€€€€Ù…ÈÕÉÉ•¹Ð€ôÁ…É•¹Ð¹±•µ•¹Ð¡¹…µ”¤ì(€€€€€€€¥˜€¡ÕÉÉ•¹Ð¥Ì¹½Ð¹Õ±°¤(€€€€€€€ì(€€€€€€€€€€€¥˜€¡½É¥¥¹…°¥Ì¹Õ±°¤ÕÉÉ•¹Ð¹I•µ½Ù” ¤ì(€€€€€€€€€€€•±Í”ÕÉÉ•¹Ð¹I•Á±…•]¥Ñ ¡¹•Üa±•µ•¹Ð¡½É¥¥¹…°¤¤ì(€€€€€€€ô(€€€€€€€•±Í”¥˜€¡½É¥¥¹…°¥Ì¹½Ð¹Õ±°¤(€€€€€€€ì(€€€€€€€€€€€Á…É•¹Ð¹‘¡¹•Üa±•µ•¹Ð¡½É¥¥¹…°¤¤ì(€€€€€€€ô(€€€ô((€€€ÁÉ¥Ù…Ñ”ÍÑ…Ñ¥ŒÍÑÉ¥¹œI•Í½±Ù•]½É­Í¡••ÑA…Ñ ¡i¥ÁÉ¡¥Ù”…É¡¥Ù”°ÍÑÉ¥¹œÝ½É­Í¡••Ñ9…µ”¤(€€€ì(€€€€€€€a9…µ•ÍÁ…”ÍÁÉ•…‘Í¡••Ð€ô€‰¡ÑÑÀè¼½Í¡•µ…Ì¹½Á•¹áµ±™½Éµ…ÑÌ¹½Éœ½ÍÁÉ•…‘Í¡••Ñµ°¼ÈÀÀØ½µ…¥¸ˆì(€€€€€€€a9…µ•ÍÁ…”½™™¥•I•±…Ñ¥½¹Í¡¥ÁÌ€ô€‰¡ÑÑÀè¼½Í¡•µ…Ì¹½Á•¹áµ±™½Éµ…ÑÌ¹½Éœ½½™™¥•½Õµ•¹Ð¼ÈÀÀØ½É•±…Ñ¥½¹Í¡¥ÁÌˆì(€€€€€€€a9…µ•ÍÁ…”Á…­…•I•±…Ñ¥½¹Í¡¥ÁÌ€ô€‰¡ÑÑÀè¼½Í¡•µ…Ì¹½Á•¹áµ±™½Éµ…ÑÌ¹½Éœ½Á…­…”¼ÈÀÀØ½É•±…Ñ¥½¹Í¡¥ÁÌˆì(€€€€€€€Ù…ÈÝ½É­‰½½¬€ô1½…‘aµ±¹ÑÉä¡…É¡¥Ù”°€‰á°½Ý½É­‰½½¬¹áµ°ˆ¤ì(€€€€€€€Ù…ÈÉ•±…Ñ¥½¹Í¡¥ÁÌ€ô1½…‘aµ±¹ÑÉä¡…É¡¥Ù”°€‰á°½}É•±Ì½Ý½É­‰½½¬¹áµ°¹É•±Ìˆ¤ì(€€€€€€€Ù…ÈÍ¡••Ð€ôÝ½É­‰½½¬¹I½½Ðü¹±•µ•¹Ð¡ÍÁÉ•…‘Í¡••Ð€¬€‰Í¡••ÑÌˆ¤ü¹±•µ•¹ÑÌ¡ÍÁÉ•…‘Í¡••Ð€¬€‰Í¡••Ðˆ¤¹M¥¹±•=É•™…Õ±Ð¡à€ôøà¹ÑÑÉ¥‰ÕÑ” ‰¹…µ”ˆ¤ü¹Y…±Õ”€ôôÝ½É­Í¡••Ñ9…µ”¤(€€€€€€€€€€€€üüÑ¡É½Ü¹•Ü%¹Ù…±¥‘=Á•É…Ñ¥½¹á•ÁÑ¥½¸ ‹–Þ—’ös¢†£’â7–¶c–r£¾òiíÝ½É­Í¡••Ñ9…µ•ôˆ¤ì(€€€€€€€Ù…ÈÉ•±…Ñ¥½¹Í¡¥Á%€ôÍ¡••Ð¹ÑÑÉ¥‰ÕÑ”¡½™™¥•I•±…Ñ¥½¹Í¡¥ÁÌ€¬€‰¥ˆ¤ü¹Y…±Õ”(€€€€€€€€€€€€üüÑ¡É½Ü¹•Ü%¹Ù…±¥‘…Ñ…á•ÁÑ¥½¸ ‹–Þ—’ös¢†£–ÏžÎï’â7–¶c–r£¾òiíÝ½É­Í¡••Ñ9…µ•ôˆ¤ì(€€€€€€€Ù…ÈÑ…É•Ð€ôÉ•±…Ñ¥½¹Í¡¥ÁÌ¹I½½Ðü¹±•µ•¹ÑÌ¡Á…­…•I•±…Ñ¥½¹Í¡¥ÁÌ€¬€‰I•±…Ñ¥½¹Í¡¥Àˆ¤¹M¥¹±•=É•™…Õ±Ð¡à€ôøà¹ÑÑÉ¥‰ÕÑ” ‰%ˆ¤ü¹Y…±Õ”€ôôÉ•±…Ñ¥½¹Í¡¥Á%¤ü¹ÑÑÉ¥‰ÕÑ” ‰Q…É•Ðˆ¤ü¹Y…±Õ”(€€€€€€€€€€€€üüÑ¡É½Ü¹•Ü%¹Ù…±¥‘…Ñ…á•ÁÑ¥½¸ ‹–Þ—’ös¢†£žn»š‚’â7–¶c–r£¾òiíÝ½É­Í¡••Ñ9…µ•ôˆ¤ì(€€€€€€€Ñ…É•Ð€ôÑ…É•Ð¹I•Á±…” qpœ°€œ¼œ¤ì(€€€€€€€¥˜€¡Ñ…É•Ð¹MÑ…ÉÑÍ]¥Ñ  ˆ¼ˆ°MÑÉ¥¹½µÁ…É¥Í½¸¹=É‘¥¹…°¤¤É•ÑÕÉ¸Ñ…É•Ð¹QÉ¥µMÑ…ÉÐ œ¼œ¤ì(€€€€€€€¥˜€¡Ñ…É•Ð¹MÑ…ÉÑÍ]¥Ñ  ˆ¸¸¼ˆ°MÑÉ¥¹½µÁ…É¥Í½¸¹=É‘¥¹…°¤¤É•ÑÕÉ¸€‰á°¼ˆ€¬Ñ…É•ÑlÌ¸¹tì(€€€€€€€É•ÑÕÉ¸€‰á°¼ˆ€¬Ñ…É•Ð¹QÉ¥µMÑ…ÉÐ œ¼œ¤ì(€€€ô((€€€ÁÉ¥Ù…Ñ”ÍÑ…Ñ¥Œa½Õµ•¹Ð1½…‘aµ±¹ÑÉä¡i¥ÁÉ¡¥Ù”…É¡¥Ù”°ÍÑÉ¥¹œÁ…Ñ ¤(€€€ì(€€€€€€€Ù…È•¹ÑÉä€ô…É¡¥Ù”¹•Ñ¹ÑÉä¡Á…Ñ ¤€üüÑ¡É½Ü¹•Ü%¹Ù…±¥‘…Ñ…á•ÁÑ¥½¸ ‰á•°ƒšZ’îÛžòë–ÂDa50ƒ¦£’îÛ¾òiíÁ…Ñ¡ôˆ¤ì(€€€€€€€ÕÍ¥¹œÙ…ÈÍÑÉ•…´€ô•¹ÑÉä¹=Á•¸ ¤ì(€€€€€€€É•ÑÕÉ¸a½Õµ•¹Ð¹1½…¡ÍÑÉ•…´¤ì(€€€ô)ô()ÁÕ‰±¥ŒÍ•…±•±…ÍÌÁ…Á¥É•‘•¹Ñ¥…±MÑ½É”¡ÍÑÉ¥¹œ‘¥É•Ñ½Éä¤)ì(€€€ÁÕ‰±¥Œ…Íå¹ŒQ…Í¬M…Ù•Íå¹Œ¡ÍÑÉ¥¹œ…½Õ¹Ð°ÍÑÉ¥¹œÁ…ÍÍÝ½É°…¹•±±…Ñ¥½¹Q½­•¸…¹•±±…Ñ¥½¹Q½­•¸€ô‘•™…Õ±Ð¤(€€€ì(€€€€€€€¥É•Ñ½Éä¹É•…Ñ•¥É•Ñ½Éä¡‘¥É•Ñ½Éä¤ì(€€€€€€€Ù…È¥Á¡•È€ôAÉ½Ñ•Ñ•‘…Ñ„¹AÉ½Ñ•Ð¡¹½‘¥¹œ¹UQà¹•Ñ	åÑ•Ì¡Á…ÍÍÝ½É¤°¹ÑÉ½Áä¡…½Õ¹Ð¤°…Ñ…AÉ½Ñ•Ñ¥½¹M½Á”¹ÕÉÉ•¹ÑUÍ•È¤ì(€€€€€€€Ù…ÈÁ…Ñ €ôA…Ñ¡½È¡…½Õ¹Ð¤ì(€€€€€€€Ù…ÈÑ•µÁA…Ñ €ôA…Ñ ¹½µ‰¥¹”¡‘¥É•Ñ½Éä°€ˆ¹íA…Ñ ¹•Ñ¥±•9…µ”¡Á…Ñ ¥ô¹íÕ¥¹9•ÝÕ¥ ¤é9ô¹ÑµÀˆ¤ì(€€€€€€€ÑÉä(€€€€€€€ì(€€€€€€€€€€€…Ý…¥Ð¥±”¹]É¥Ñ•±±	åÑ•ÍÍå¹Œ¡Ñ•µÁA…Ñ °¥Á¡•È°…¹•±±…Ñ¥½¹Q½­•¸¤ì(€€€€€€€€€€€ÕÍ¥¹œ€¡Ù…È‘ÕÉ…‰±”€ô¹•Ü¥±•MÑÉ•…´¡Ñ•µÁA…Ñ °¥±•5½‘”¹=Á•¸°¥±••ÍÌ¹I•…‘]É¥Ñ”°¥±•M¡…É”¹I•…¤¤(€€€€€€€€€€€€€€€‘ÕÉ…‰±”¹±ÕÍ ¡ÑÉÕ”¤ì((€€€€€€€€€€€¥˜€¡¥±”¹á¥ÍÑÌ¡Á…Ñ ¤¤¥±”¹I•Á±…”¡Ñ•µÁA…Ñ °Á…Ñ °¹Õ±°°ÑÉÕ”¤ì(€€€€€€€€€€€•±Í”¥±”¹5½Ù”¡Ñ•µÁA…Ñ °Á…Ñ ¤ì(€€€€€€€ô(€€€€€€€™¥¹…±±ä(€€€€€€€ì(€€€€€€€€€€€¥˜€¡¥±”¹á¥ÍÑÌ¡Ñ•µÁA…Ñ ¤¤¥±”¹•±•Ñ”¡Ñ•µÁA…Ñ ¤ì(€€€€€€€ô(€€€ô((€€€ÁÕ‰±¥Œ…Íå¹ŒQ…Í¬ñÍÑÉ¥¹œüø1½…‘Íå¹Œ¡ÍÑÉ¥¹œ…½Õ¹Ð°…¹•±±…Ñ¥½¹Q½­•¸…¹•±±…Ñ¥½¹Q½­•¸€ô‘•™…Õ±Ð¤(€€€ì(€€€€€€€Ù…ÈÁ…Ñ €ôA…Ñ¡½È¡…½Õ¹Ð¤ì(€€€€€€€¥˜€ …¥±”¹á¥ÍÑÌ¡Á…Ñ ¤¤É•ÑÕÉ¸¹Õ±°ì(€€€€€€€Ù…È¥Á¡•È€ô…Ý…¥Ð¥±”¹I•…‘±±	åÑ•ÍÍå¹Œ¡Á…Ñ °…¹•±±…Ñ¥½¹Q½­•¸¤ì(€€€€€€€Ù…È±•…È€ôAÉ½Ñ•Ñ•‘…Ñ„¹U¹ÁÉ½Ñ•Ð¡¥Á¡•È°¹ÑÉ½Áä¡…½Õ¹Ð¤°…Ñ…AÉ½Ñ•Ñ¥½¹M½Á”¹ÕÉÉ•¹ÑUÍ•È¤ì(€€€€€€€É•ÑÕÉ¸¹½‘¥¹œ¹UQà¹•ÑMÑÉ¥¹œ¡±•…È¤ì(€€€ô((€€€ÁÕ‰±¥ŒQ…Í¬•±•Ñ•Íå¹Œ¡ÍÑÉ¥¹œ…½Õ¹Ð°…¹•±±…Ñ¥½¹Q½­•¸…¹•±±…Ñ¥½¹Q½­•¸€ô‘•™…Õ±Ð¤(€€€ì(€€€€€€€…¹•±±…Ñ¥½¹Q½­•¸¹Q¡É½Ý%™…¹•±±…Ñ¥½¹I•ÅÕ•ÍÑ• ¤ì(€€€€€€€Ù…ÈÁ…Ñ €ôA…Ñ¡½È¡…½Õ¹Ð¤ì(€€€€€€€¥˜€¡¥±”¹á¥ÍÑÌ¡Á…Ñ ¤¤¥±”¹•±•Ñ”¡Á…Ñ ¤ì(€€€€€€€É•ÑÕÉ¸Q…Í¬¹½µÁ±•Ñ•‘Q…Í¬ì(€€€ô((€€€ÁÉ¥Ù…Ñ”ÍÑ…Ñ¥Œ‰åÑ•mt¹ÑÉ½Áä¡ÍÑÉ¥¹œ…½Õ¹Ð¤€ôøM!ÈÔØ¹!…Í¡…Ñ„¡¹½‘¥¹œ¹UQà¹•Ñ	åÑ•Ì ‰%¹Ù½¥•5…¥±ÍÍ¥ÍÑ…¹Ññí…½Õ¹Ð¹QÉ¥´ ¤¹Q½1½Ý•É%¹Ù…É¥…¹Ð ¥ôˆ¤¤ì(€€€ÁÉ¥Ù…Ñ”ÍÑÉ¥¹œA…Ñ¡½È¡ÍÑÉ¥¹œ…½Õ¹Ð¤€ôøA…Ñ ¹½µ‰¥¹”¡‘¥É•Ñ½Éä°½¹Ù•ÉÐ¹Q½!•áMÑÉ¥¹œ¡M!ÈÔØ¹!…Í¡…Ñ„¡¹½‘¥¹œ¹UQà¹•Ñ	åÑ•Ì¡…½Õ¹Ð¹QÉ¥´ ¤¹Q½1½Ý•É%¹Ù…É¥…¹Ð ¤¤¤¤€¬€ˆ¹É•ˆ¤ì)ô()ÁÕ‰±¥ŒÍÑ…Ñ¥Œ±…ÍÌMÑ…ÉÑÕÁ5…¹…•È)ì(€€€ÁÉ¥Ù…Ñ”½¹ÍÐÍÑÉ¥¹œIÕ¹-•ä€ô€‰M½™ÑÝ…É•qq5¥É½Í½™Ñqq]¥¹‘½ÝÍqqÕÉÉ•¹ÑY•ÉÍ¥½¹qqIÕ¸ˆì(€€€ÁÉ¥Ù…Ñ”½¹ÍÐÍÑÉ¥¹œY…±Õ•9…µ”€ô€‰%¹Ù½¥•5…¥±ÍÍ¥ÍÑ…¹Ðˆì((€€€ÁÕ‰±¥ŒÍÑ…Ñ¥ŒÙ½¥M•Ñ¹…‰±•¡‰½½°•¹…‰±•¤(€€€ì(€€€€€€€ÑÉä(€€€€€€€ì(€€€€€€€€€€€ÕÍ¥¹œÙ…È­•ä€ôI•¥ÍÑÉä¹ÕÉÉ•¹ÑUÍ•È¹É•…Ñ•MÕ‰-•ä¡IÕ¹-•ä°ÑÉÕ”¤(€€€€€€€€€€€€€€€€üüÑ¡É½Ü¹•Ü%¹Ù…±¥‘=Á•É…Ñ¥½¹á•ÁÑ¥½¸ ‹š^ƒšÎWš&O–ò–öO–&7žR£š"ßžj–òšrë–B¿–*£šÎ£–3¢†£¦†çŽˆ¤ì(€€€€€€€€€€€¥˜€ …•¹…‰±•¤(€€€€€€€€€€€ì(€€€€€€€€€€€€€€€­•ä¹•±•Ñ•Y…±Õ”¡Y…±Õ•9…µ”°™…±Í”¤ì(€€€€€€€€€€€€€€€É•ÑÕÉ¸ì(€€€€€€€€€€€ô((€€€€€€€€€€€Ù…ÈÁÉ½•ÍÍA…Ñ €ô¹Ù¥É½¹µ•¹Ð¹AÉ½•ÍÍA…Ñ ì(€€€€€€€€€€€¥˜€¡ÍÑÉ¥¹œ¹%Í9Õ±±=É]¡¥Ñ•MÁ…”¡ÁÉ½•ÍÍA…Ñ ¤¤Ñ¡É½Ü¹•Ü%¹Ù…±¥‘=Á•É…Ñ¥½¹á•ÁÑ¥½¸ ‹š^ƒšÎWž†»–ºkž¢/–ê?¢Þ¿–úŽˆ¤ì(€€€€€€€€€€€­•ä¹M•ÑY…±Õ”¡Y…±Õ•9…µ”°€‰p‰íÁÉ½•ÍÍA…Ñ¡õpˆˆ¤ì(€€€€€€€ô(€€€€€€€…Ñ €¡á•ÁÑ¥½¸•à¤Ý¡•¸€¡•à¥ÌU¹…ÕÑ¡½É¥é•‘•ÍÍá•ÁÑ¥½¸½ÈM•ÕÉ¥Ñåá•ÁÑ¥½¸½È%=á•ÁÑ¥½¸¤(€€€€€€€ì(€€€€€€€€€€€Ñ¡É½Ü¹•Ü%¹Ù…±¥‘=Á•É…Ñ¥½¹á•ÁÑ¥½¸ ‹–òšrë–B¿–*£¢ºûžö»–’Ç¢Ò—Žˆ°•à¤ì(€€€€€€€ô(€€€ô)ô