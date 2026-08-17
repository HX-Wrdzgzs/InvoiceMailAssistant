using System.IO;
using System.IO.Compression;
using System.Globalization;
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
    private const string SubjectPrefix = "中外运向您提交了开票申请";
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
            var connectionKey = BuildConnectionKey(account, host, port);
            for (var attempt = 0; ; attempt++)
            {
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
                catch (Exception ex) when (IsTransient(ex) && attempt == 0 && HasSessionForKey(connectionKey))
                {
                    // A long-lived IMAP socket can look connected locally after
                    // the server or a network device has already closed it. Do
                    // one controlled reconnect in the same poll so the user does
                    // not have to press "重新连接" manually. A single retry is
                    // intentional: persistent failures still enter backoff.
                    await DisconnectCoreAsync();
                    await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken);
                }
                catch (Exception ex) when (IsTransient(ex))
                {
                    await DisconnectCoreAsync();
                    RegisterFailure();
                    throw;
                }
            }
        }
        finally
        {
            _gate.Release();
        }
    }

    private async Task EnsureConnectedAsync(string account, string password, string host, int port, CancellationToken cancellationToken)
    {
        var key = BuildConnectionKey(account, host, port);
        if (_session?.IsConnected == true && string.Equals(_connectionKey, key, StringComparison.Ordinal))
            return;

        if (_retryAfterUtc > DateTimeOffset.UtcNow)
            throw new MailboxBackoffException($"邮箱连接失败，正在退避重试（下次重试时间：{_retryAfterUtc.ToLocalTime():HH:mm:ss}）。");

        await DisconnectCoreAsync();
        await ConnectCoreAsync(account, password, host, port, cancellationToken);
    }

    private async Task ConnectCoreAsync(string account, string password, string host, int port, CancellationToken cancellationToken)
    {
        var session = await _sessionFactory.ConnectAsync(account, password, host, port, cancellationToken);
        _session = session;
        _connectionKey = BuildConnectionKey(account, host, port);
        _failureCount = 0;
        _retryAfterUtc = DateTimeOffset.MinValue;
    }

    private bool HasSessionForKey(string connectionKey)
        => _session is not null && string.Equals(_connectionKey, connectionKey, StringComparison.Ordinal);

    private static string BuildConnectionKey(string account, string host, int port)
        => $"{account.Trim().ToLowerInvariant()}|{host.Trim().ToLowerInvariant()}|{port}";

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
    private const string SubjectPrefix = "中外运向您提交了开票申请";

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
                    $"邮件 MIME 解析失败：{ex.Message}"));
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
        command.CommandText = "SELECT 1 FROM invoice_applications WHERE (message_id <> '' AND message_id = $messageId AND (mailbox_identity = $mailbox OR mailbox_identity = '')) OR ((mailbox_identity = $mailbox OR mailbox_identity = '') AND mailbox_name = $mailboxName AND uid_validity = $uidValidity AND imap_uid = $uid) OR ((mailbox_identity = $mailbox OR mailbox_identity = '') AND (fallback_hash = $hash OR ($legacyHash IS NOT NULL AND fallback_hash = $legacyHash))) LIMIT 1;";
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
        var legacyFallbackHash = app.ProcessingStatus == ProcessingStatus.Parsed
            ? Deduplication.CreateLegacyFallbackHash(app)
            : null;
        if (await ExistsAnyAsync(app.MessageId, app.ImapUid, app.MailboxIdentity, app.MailboxName, app.UidValidity, fallbackHash, legacyFallbackHash, cancellationToken))
            return null;

        try
        {
            return await InsertAsync(app, fallbackHash, cancellationToken);
        }
        catch (SqliteException ex) when (IsUniqueConstraint(ex))
        {
            return null;
        }
    }

    public async Task<long> InsertAsync(InvoiceApplication app, string fallbackHash, CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenConnectionAsync(cancellationToken);
        var command = connection.CreateCommand();
        command.CommandText = """
            INSERT INTO invoice_applications (
                company_name, credit_code, amount, apply_time, invoice_type, recipient, phone, address, email, remark,
                message_id, imap_uid, uid_validity, mailbox_name, mailbox_identity, mail_received_at, mail_subject, mail_from, normalized_body,
                processing_status, excel_row, created_at, updated_at, error_message, fallback_hash)
            VALUES ($company, $credit, $amount, $apply, $type, $recipient, $phone, $address, $email, $remark,
                $messageId, $uid, $uidValidity, $mailboxName, $mailbox, $received, $subject, $from, $body, $status, $excelRow, $created, $updated, $error, $hash);
            SELECT last_insert_rowid();
            """;
        AddParameters(command, app, fallbackHash);
        var id = (long)(await command.ExecuteScalarAsync(cancellationToken) ?? 0L);
        app.Id = id;
        return id;
    }

    public async Task UpdateStatusAsync(long id, ProcessingStatus status, string? error = null, int? excelRow = null, CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenConnectionAsync(cancellationToken);
        var command = connection.CreateCommand();
        command.CommandText = "UPDATE invoice_applications SET processing_status=$status, error_message=$error, excel_row=COALESCE($row, excel_row), updated_at=$updated WHERE id=$id;";
        command.Parameters.AddWithValue("$status", (int)status);
        command.Parameters.AddWithValue("$error", error ?? string.Empty);
        command.Parameters.AddWithValue("$row", (object?)excelRow ?? DBNull.Value);
        command.Parameters.AddWithValue("$updated", DateTimeOffset.Now.ToString("O"));
        command.Parameters.AddWithValue("$id", id);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<IReadOnlyList<InvoiceApplication>> GetRecentAsync(int limit, CancellationToken cancellationToken = default)
    {
        var result = new List<InvoiceApplication>();
        await using var connection = await OpenConnectionAsync(cancellationToken);
        var command = connection.CreateCommand();
        command.CommandText = "SELECT * FROM invoice_applications ORDER BY id DESC LIMIT $limit;";
        command.Parameters.AddWithValue("$limit", limit);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken)) result.Add(Map(reader));
        return result;
    }

    public async Task<IReadOnlyList<InvoiceApplication>> GetPendingExcelAsync(CancellationToken cancellationToken = default)
    {
        var result = new List<InvoiceApplication>();
        await using var connection = await OpenConnectionAsync(cancellationToken);
        var command = connection.CreateCommand();
        command.CommandText = "SELECT * FROM invoice_applications WHERE processing_status IN (2, 5) ORDER BY id;";
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken)) result.Add(Map(reader));
        return result;
    }

    public async Task<IReadOnlyList<InvoiceApplication>> GetParseFailedAsync(CancellationToken cancellationToken = default)
    {
        var result = new List<InvoiceApplication>();
        await using var connection = await OpenConnectionAsync(cancellationToken);
        var command = connection.CreateCommand();
        command.CommandText = "SELECT * FROM invoice_applications WHERE processing_status = $status ORDER BY id;";
        command.Parameters.AddWithValue("$status", (int)ProcessingStatus.ParseFailed);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken)) result.Add(Map(reader));
        return result;
    }

    public async Task<IReadOnlyList<InvoiceApplication>> GetCompletedRepeatedFormRecordsAsync(CancellationToken cancellationToken = default)
    {
        var result = new List<InvoiceApplication>();
        await using var connection = await OpenConnectionAsync(cancellationToken);
        var command = connection.CreateCommand();
        command.CommandText = "SELECT * FROM invoice_applications WHERE processing_status = $status AND normalized_body LIKE '%公司名称%公司名称%' ORDER BY id;";
        command.Parameters.AddWithValue("$status", (int)ProcessingStatus.Completed);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken)) result.Add(Map(reader));
        return result;
    }

    public async Task UpdateParsedAsync(long id, InvoiceApplication app, string fallbackHash, CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenConnectionAsync(cancellationToken);
        var command = connection.CreateCommand();
        command.CommandText = """
            UPDATE invoice_applications SET
                company_name=$company, credit_code=$credit, amount=$amount, apply_time=$apply,
                invoice_type=$type, recipient=$recipient, phone=$phone, address=$address, email=$email, remark=$remark,
                message_id=$messageId, imap_uid=$uid, uid_validity=$uidValidity, mailbox_name=$mailboxName,
                mailbox_identity=$mailbox, mail_received_at=$received, mail_subject=$subject, mail_from=$from,
                normalized_body=$body, processing_status=$status, excel_row=$excelRow, updated_at=$updated,
                error_message=$error, fallback_hash=$hash
            WHERE id=$id;
            """;
        app.Id = id;
        app.ProcessingStatus = ProcessingStatus.PendingExcel;
        app.ErrorMessage = string.Empty;
        app.UpdatedAt = DateTimeOffset.Now;
        AddParameters(command, app, fallbackHash);
        command.Parameters.AddWithValue("$id", id);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static void AddParameters(SqliteCommand command, InvoiceApplication app, string hash)
    {
        command.Parameters.AddWithValue("$company", app.CompanyName);
        command.Parameters.AddWithValue("$credit", app.CreditCode);
        command.Parameters.AddWithValue("$amount", app.Amount.ToString(System.Globalization.CultureInfo.InvariantCulture));
        command.Parameters.AddWithValue("$apply", app.ApplyTime.ToString("O"));
        command.Parameters.AddWithValue("$type", app.InvoiceType);
        command.Parameters.AddWithValue("$recipient", app.Recipient);
        command.Parameters.AddWithValue("$phone", app.Phone);
        command.Parameters.AddWithValue("$address", app.Address);
        command.Parameters.AddWithValue("$email", app.Email);
        command.Parameters.AddWithValue("$remark", app.Remark);
        command.Parameters.AddWithValue("$messageId", app.MessageId);
        command.Parameters.AddWithValue("$uid", (long)app.ImapUid);
        command.Parameters.AddWithValue("$uidValidity", (long)app.UidValidity);
        command.Parameters.AddWithValue("$mailboxName", app.MailboxName);
        command.Parameters.AddWithValue("$mailbox", app.MailboxIdentity);
        command.Parameters.AddWithValue("$received", app.MailReceivedAt.ToString("O"));
        command.Parameters.AddWithValue("$subject", app.MailSubject);
        command.Parameters.AddWithValue("$from", app.MailFrom);
        command.Parameters.AddWithValue("$body", app.NormalizedBody);
        command.Parameters.AddWithValue("$status", (int)app.ProcessingStatus);
        command.Parameters.AddWithValue("$excelRow", (object?)app.ExcelRow ?? DBNull.Value);
        command.Parameters.AddWithValue("$created", app.CreatedAt.ToString("O"));
        command.Parameters.AddWithValue("$updated", app.UpdatedAt.ToString("O"));
        command.Parameters.AddWithValue("$error", app.ErrorMessage);
        command.Parameters.AddWithValue("$hash", hash);
    }

    private async Task<SqliteConnection> OpenConnectionAsync(CancellationToken cancellationToken)
    {
        var connection = new SqliteConnection(_connectionString);
        try
        {
            await connection.OpenAsync(cancellationToken);
            var pragma = connection.CreateCommand();
            pragma.CommandText = "PRAGMA busy_timeout=10000;";
            await pragma.ExecuteNonQueryAsync(cancellationToken);
            return connection;
        }
        catch
        {
            await connection.DisposeAsync();
            throw;
        }
    }

    private static InvoiceApplication Map(SqliteDataReader r) => new()
    {
        Id = r.GetInt64(r.GetOrdinal("id")),
        CompanyName = r.GetString(r.GetOrdinal("company_name")),
        CreditCode = r.GetString(r.GetOrdinal("credit_code")),
        Amount = decimal.Parse(r.GetString(r.GetOrdinal("amount")), System.Globalization.CultureInfo.InvariantCulture),
        ApplyTime = DateTime.Parse(r.GetString(r.GetOrdinal("apply_time")), null, System.Globalization.DateTimeStyles.RoundtripKind),
        InvoiceType = r.GetString(r.GetOrdinal("invoice_type")),
        Recipient = r.GetString(r.GetOrdinal("recipient")),
        Phone = r.GetString(r.GetOrdinal("phone")),
        Address = r.GetString(r.GetOrdinal("address")),
        Email = r.GetString(r.GetOrdinal("email")),
        Remark = r.GetString(r.GetOrdinal("remark")),
        MessageId = r.GetString(r.GetOrdinal("message_id")),
        ImapUid = checked((uint)r.GetInt64(r.GetOrdinal("imap_uid"))),
        UidValidity = checked((uint)r.GetInt64(r.GetOrdinal("uid_validity"))),
        MailboxName = r.GetString(r.GetOrdinal("mailbox_name")),
        MailboxIdentity = r.GetString(r.GetOrdinal("mailbox_identity")),
        MailReceivedAt = DateTimeOffset.Parse(r.GetString(r.GetOrdinal("mail_received_at"))),
        MailSubject = r.GetString(r.GetOrdinal("mail_subject")),
        MailFrom = r.GetString(r.GetOrdinal("mail_from")),
        NormalizedBody = r.GetString(r.GetOrdinal("normalized_body")),
        ProcessingStatus = (ProcessingStatus)r.GetInt32(r.GetOrdinal("processing_status")),
        ExcelRow = r.IsDBNull(r.GetOrdinal("excel_row")) ? null : r.GetInt32(r.GetOrdinal("excel_row")),
        CreatedAt = DateTimeOffset.Parse(r.GetString(r.GetOrdinal("created_at"))),
        UpdatedAt = DateTimeOffset.Parse(r.GetString(r.GetOrdinal("updated_at"))),
        ErrorMessage = r.GetString(r.GetOrdinal("error_message"))
    };

    private static async Task EnsureColumnAsync(SqliteConnection connection, string name, string definition, CancellationToken cancellationToken)
    {
        var check = connection.CreateCommand();
        check.CommandText = "PRAGMA table_info(invoice_applications);";
        var exists = false;
        await using (var reader = await check.ExecuteReaderAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                if (string.Equals(reader.GetString(1), name, StringComparison.OrdinalIgnoreCase))
                {
                    exists = true;
                    break;
                }
            }
        }

        if (!exists)
        {
            var alter = connection.CreateCommand();
            alter.CommandText = $"ALTER TABLE invoice_applications ADD COLUMN {name} {definition};";
            await alter.ExecuteNonQueryAsync(cancellationToken);
        }
    }

    private static bool IsUniqueConstraint(SqliteException exception)
        => exception.SqliteErrorCode == 19 || exception.Message.Contains("UNIQUE", StringComparison.OrdinalIgnoreCase);
}

public sealed class ExcelRowOccupiedException(string message) : Exception(message);

public sealed class ExcelWriterLock : IDisposable
{
    private readonly Semaphore _semaphore;
    private bool _released;

    internal ExcelWriterLock(Semaphore semaphore) => _semaphore = semaphore;

    public void Dispose()
    {
        if (_released) return;
        _released = true;
        _semaphore.Release();
        _semaphore.Dispose();
    }
}

public sealed class ExcelWriter
{
    private const string LockName = "Local\\InvoiceMailAssistant.ExcelWrite";
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string OfficeRelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string HyperlinkRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

    private sealed record OutputStyles(int BoldTextStyleId, int HyperlinkStyleId);

    public async Task<ExcelWriterLock> AcquireWriteLockAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var semaphore = new Semaphore(1, 1, LockName);
        var acquired = false;
        try
        {
            acquired = await Task.Run(() => semaphore.WaitOne(TimeSpan.FromSeconds(30)), CancellationToken.None);
            if (!acquired) throw new IOException("Excel 写入锁等待超时，申请将保留在等待队列。");
            cancellationToken.ThrowIfCancellationRequested();
            return new ExcelWriterLock(semaphore);
        }
        catch
        {
            if (acquired) semaphore.Release();
            semaphore.Dispose();
            throw;
        }
    }

    public int ResolveTargetRow(InvoiceApplication application, string filePath, string worksheetName)
    {
        EnsureWritable(filePath);
        using var workbook = new XLWorkbook(filePath);
        if (!workbook.Worksheets.TryGetWorksheet(worksheetName, out var worksheet))
            throw new InvalidOperationException($"工作表不存在：{worksheetName}");

        var lastUsed = worksheet.LastRowUsed()?.RowNumber() ?? 1;
        if (application.ExcelRow is int plannedRow && plannedRow >= 2)
        {
            if (RowMatches(worksheet, plannedRow, application))
                return plannedRow;
            if (RowIsEmpty(worksheet, plannedRow))
                return plannedRow;

            // A user may have inserted rows above a previously planned row.
            // Reuse a matching row only when it is unambiguous; otherwise keep
            // the record pending instead of risking a different application.
            var existingRow = FindMatchingRow(worksheet, application, Math.Max(lastUsed, plannedRow));
            if (existingRow is not null) return existingRow.Value;
        }
        else
        {
            // A new or rebuilt SQLite database may not know about rows that
            // already exist in the customer's workbook. Reuse an unambiguous
            // matching row before appending, otherwise a historical scan can
            // duplicate an entire existing block.
            var existingRow = FindMatchingRow(worksheet, application, lastUsed);
            if (existingRow is not null) return existingRow.Value;
        }

        return Math.Max(lastUsed + 1, 2);
    }

    public async Task RepairExistingRowAsync(
        InvoiceApplication original,
        InvoiceApplication corrected,
        string filePath,
        string worksheetName,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (original.ExcelRow is not int targetRow || targetRow < 2)
            throw new InvalidOperationException("历史记录没有可验证的 Excel 行号。");

        var ownedLock = await AcquireWriteLockAsync(cancellationToken);

        var directory = Path.GetDirectoryName(filePath)!;
        var tempPath = Path.Combine(directory, $".{Path.GetFileName(filePath)}.{Guid.NewGuid():N}.tmp.xlsx");
        try
        {
            EnsureWritable(filePath);
            File.Copy(filePath, tempPath, true);

            DateTime? previousDate;
            using (var workbook = new XLWorkbook(tempPath))
            {
                if (!workbook.Worksheets.TryGetWorksheet(worksheetName, out var worksheet))
                    throw new InvalidOperationException($"工作表不存在：{worksheetName}");

                if (RowMatches(worksheet, targetRow, corrected)) return;
                if (!RowMatches(worksheet, targetRow, original))
                    throw new ExcelRowOccupiedException($"Excel 第 {targetRow} 行与历史记录不一致，已停止自动修复以保护人工数据。");

                previousDate = FindPreviousApplicationDate(worksheet, targetRow - 1, corrected.ApplyTime.Year);
            }

            PatchWorksheetXml(tempPath, worksheetName, corrected, targetRow, previousDate);

            using (var verify = new XLWorkbook(tempPath))
            {
                var verifySheet = verify.Worksheet(worksheetName);
                if (!RowMatches(verifySheet, targetRow, corrected))
                    throw new InvalidDataException("Excel 历史记录修复校验失败：目标行内容不一致。");
            }

            using (var durable = new FileStream(tempPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                durable.Flush(true);

            File.Replace(tempPath, filePath, null, true);
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
            ownedLock?.Dispose();
        }
    }

    public async Task<int> WriteAsync(InvoiceApplication application, string filePath, string worksheetName, CancellationToken cancellationToken = default, ExcelWriterLock? writeLock = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ExcelWriterLock? ownedLock = null;
        if (writeLock is null) ownedLock = await AcquireWriteLockAsync(cancellationToken);

        var directory = Path.GetDirectoryName(filePath)!;
        var tempPath = Path.Combine(directory, $".{Path.GetFileName(filePath)}.{Guid.NewGuid():N}.tmp.xlsx");

        try
        {
            EnsureWritable(filePath);
            File.Copy(filePath, tempPath, true);
            var row = application.ExcelRow ?? throw new InvalidOperationException("尚未为 Excel 写入预留目标行。");
            DateTime? previousDate;
            using (var workbook = new XLWorkbook(tempPath))
            {
                if (!workbook.Worksheets.TryGetWorksheet(worksheetName, out var worksheet))
                    throw new InvalidOperationException($"工作表不存在：{worksheetName}");

                if (RowMatches(worksheet, row, application))
                    return row;

                if (!RowIsEmpty(worksheet, row))
                {
                    var existingRow = FindMatchingRow(worksheet, application, Math.Max(row, worksheet.LastRowUsed()?.RowNumber() ?? row));
                    if (existingRow is null)
                        throw new ExcelRowOccupiedException($"Excel 第 {row} 行已被其他数据占用，请重新规划目标行。");
                    row = existingRow.Value;
                    application.ExcelRow = row;
                }

                if (RowMatches(worksheet, row, application))
                    return row;

                previousDate = FindPreviousApplicationDate(worksheet, row - 1, application.ApplyTime.Year);
            }

            // ClosedXML is used above only for read-only row planning. Saving the
            // whole workbook through it rewrites styles.xml and can make real
            // customer workbooks trigger Excel's repair dialog. Patch only the
            // target worksheet package part so every other workbook part remains
            // byte-for-byte unchanged.
            PatchWorksheetXml(tempPath, worksheetName, application, row, previousDate);

            using (var verify = new XLWorkbook(tempPath))
            {
                var verifySheet = verify.Worksheet(worksheetName);
                if (!RowMatches(verifySheet, row, application))
                    throw new InvalidDataException("Excel 临时文件校验失败：目标行内容不一致。");
            }

            using (var durable = new FileStream(tempPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                durable.Flush(true);

            File.Replace(tempPath, filePath, null, true);
            return row;
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
            ownedLock?.Dispose();
        }
    }

    private static void EnsureWritable(string filePath)
    {
        if (!File.Exists(filePath)) throw new FileNotFoundException("Excel 登记表不存在。", filePath);
        if (File.GetAttributes(filePath).HasFlag(FileAttributes.ReadOnly))
            throw new IOException("Excel 登记表处于只读属性，请取消只读后重试。");
        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new IOException("Excel 登记表没有写入权限，请检查文件权限后重试。", ex);
        }
        catch (IOException ex)
        {
            throw new IOException("Excel 登记表正在被占用，申请将保留在等待队列。", ex);
        }
    }

    private static bool RowIsEmpty(IXLWorksheet worksheet, int row)
    {
        var lastColumn = Math.Max(8, worksheet.LastColumnUsed()?.ColumnNumber() ?? 8);
        return Enumerable.Range(1, lastColumn).All(column => worksheet.Cell(row, column).IsEmpty());
    }

    private static int? FindMatchingRow(IXLWorksheet worksheet, InvoiceApplication application, int lastRow)
    {
        int? match = null;
        for (var row = 2; row <= lastRow; row++)
        {
            if (!RowMatches(worksheet, row, application)) continue;
            if (match is not null)
                throw new ExcelRowOccupiedException($"Excel 中存在多个与申请匹配的历史行（第 {match} 行和第 {row} 行），已停止自动追加以避免重复数据。");
            match = row;
        }
        return match;
    }

    private static bool RowMatches(IXLWorksheet worksheet, int row, InvoiceApplication application)
    {
        if (!string.Equals(worksheet.Cell(row, 2).GetString().Trim(), application.CompanyName.Trim(), StringComparison.Ordinal)) return false;
        if (!string.Equals(worksheet.Cell(row, 3).GetString().Trim(), application.CreditCode.Trim(), StringComparison.Ordinal)) return false;
        if (!decimal.TryParse(worksheet.Cell(row, 4).GetString(), out var amount) || amount != application.Amount) return false;
        if (!string.Equals(worksheet.Cell(row, 6).GetString().Trim(), application.Email.Trim(), StringComparison.OrdinalIgnoreCase)) return false;
        return ApplicationDateMatches(worksheet, row, application.ApplyTime.Date);
    }

    private static bool ApplicationDateMatches(IXLWorksheet worksheet, int row, DateTime expectedDate)
    {
        var rowDate = ParseSheetDate(worksheet.Cell(row, 1).GetString(), expectedDate.Year)
            ?? FindPreviousApplicationDate(worksheet, row - 1, expectedDate.Year);
        // A completely unlabelled legacy block has no date context to compare.
        // Its business fields can still safely identify an existing row.
        return rowDate is null || rowDate.Value.Date == expectedDate;
    }

    private static DateTime? FindPreviousApplicationDate(IXLWorksheet ws, int row, int assumedYear)
    {
        for (var current = row; current >= 2; current--)
        {
            var date = ParseSheetDate(ws.Cell(current, 1).GetString(), assumedYear);
            if (date is not null) return date;
        }
        return null;
    }

    private static DateTime? ParseSheetDate(string? text, int assumedYear)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        var parts = text.Trim().Split(['.', '/'], StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 2 || !int.TryParse(parts[0], out var month) || !int.TryParse(parts[1], out var day)) return null;
        try { return new DateTime(assumedYear, month, day); }
        catch (ArgumentOutOfRangeException) { return null; }
    }

    private static void PatchWorksheetXml(string filePath, string worksheetName, InvoiceApplication application, int targetRow, DateTime? previousDate)
    {
        using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
        var worksheetPath = ResolveWorksheetPath(archive, worksheetName);
        var entry = archive.GetEntry(worksheetPath) ?? throw new InvalidDataException($"工作表 XML 不存在：{worksheetName}");
        XDocument document;
        using (var stream = entry.Open())
            document = XDocument.Load(stream, System.Xml.Linq.LoadOptions.PreserveWhitespace);

        XNamespace spreadsheet = SpreadsheetNamespace;
        var root = document.Root ?? throw new InvalidDataException($"工作表 XML 无根节点：{worksheetName}");
        var sheetData = root.Element(spreadsheet + "sheetData")
            ?? throw new InvalidDataException($"工作表 XML 缺少 sheetData：{worksheetName}");
        var rows = sheetData.Elements(spreadsheet + "row").ToArray();
        var target = rows.SingleOrDefault(x => RowNumber(x) == targetRow);
        var previous = rows.Where(x => RowNumber(x) < targetRow).OrderBy(x => RowNumber(x)).LastOrDefault();

        if (target is null)
        {
            target = new XElement(spreadsheet + "row", new XAttribute("r", targetRow));
            CopyRowFormatting(previous, target, spreadsheet, targetRow);
            var next = rows.FirstOrDefault(x => RowNumber(x) > targetRow);
            if (next is null) sheetData.Add(target);
            else next.AddBeforeSelf(target);
        }

        var outputStyles = EnsureOutputStyles(
            archive,
            GetStyleId(FindCell(target, spreadsheet, 2)) ?? GetStyleId(FindCell(previous, spreadsheet, 2)),
            GetStyleId(FindCell(target, spreadsheet, 6)) ?? GetStyleId(FindCell(previous, spreadsheet, 6)));

        if (previousDate?.Date != application.ApplyTime.Date)
            SetInlineString(target, spreadsheet, 1, $"{application.ApplyTime.Month}.{application.ApplyTime.Day}", previous, targetRow);
        else
            ClearCell(target, spreadsheet, 1, previous, targetRow);

        SetInlineString(target, spreadsheet, 2, application.CompanyName, previous, targetRow, outputStyles.BoldTextStyleId);
        SetInlineString(target, spreadsheet, 3, application.CreditCode, previous, targetRow, outputStyles.BoldTextStyleId);
        SetNumeric(target, spreadsheet, 4, application.Amount, previous, targetRow);
        SetInlineString(target, spreadsheet, 6, application.Email, previous, targetRow, outputStyles.HyperlinkStyleId);
        UpsertEmailHyperlink(archive, worksheetPath, root, targetRow, application.Email);
        ExpandDimension(root, spreadsheet, targetRow, 8);

        entry.Delete();
        var replacement = archive.CreateEntry(worksheetPath, CompressionLevel.Optimal);
        using var output = replacement.Open();
        document.Save(output, System.Xml.Linq.SaveOptions.DisableFormatting);
    }

    private static void CopyRowFormatting(XElement? source, XElement target, XNamespace spreadsheet, int targetRow)
    {
        if (source is null) return;
        foreach (var attribute in source.Attributes().Where(x => x.Name.LocalName is "ht" or "customHeight" or "s" or "customFormat"))
            target.SetAttributeValue(attribute.Name, attribute.Value);

        for (var column = 1; column <= 8; column++)
        {
            var sourceCell = FindCell(source, spreadsheet, column);
            if (sourceCell is null) continue;
            var cell = new XElement(sourceCell);
            cell.SetAttributeValue("r", CellReference(column, targetRow));
            cell.SetAttributeValue("t", null);
            cell.RemoveNodes();
            target.Add(cell);
        }
    }

    private static void SetInlineString(XElement row, XNamespace spreadsheet, int column, string value, XElement? styleSource, int rowNumber, int? styleId = null)
    {
        var cell = GetOrCreateCell(row, spreadsheet, column, styleSource, rowNumber);
        ClearCellContent(cell, spreadsheet);
        if (styleId is int outputStyleId) cell.SetAttributeValue("s", outputStyleId);
        cell.SetAttributeValue("t", "inlineStr");
        var text = new XElement(spreadsheet + "t", value);
        if (value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1])))
            text.SetAttributeValue(XNamespace.Xml + "space", "preserve");
        AddCellContent(cell, spreadsheet, new XElement(spreadsheet + "is", text));
    }

    private static OutputStyles EnsureOutputStyles(ZipArchive archive, int? baseTextStyleId, int? baseEmailStyleId)
    {
        var document = LoadXmlEntry(archive, "xl/styles.xml");
        XNamespace spreadsheet = SpreadsheetNamespace;
        var root = document.Root ?? throw new InvalidDataException("Excel 样式 XML 无根节点。");
        var fonts = root.Element(spreadsheet + "fonts") ?? throw new InvalidDataException("Excel 样式 XML 缺少 fonts。");
        var cellXfs = root.Element(spreadsheet + "cellXfs") ?? throw new InvalidDataException("Excel 样式 XML 缺少 cellXfs。");

        var boldTextStyleId = EnsureStyleVariant(
            fonts,
            cellXfs,
            baseTextStyleId ?? 0,
            font => EnsureBoldFont(font, spreadsheet));
        var hyperlinkStyleId = EnsureStyleVariant(
            fonts,
            cellXfs,
            baseEmailStyleId ?? boldTextStyleId,
            font => EnsureHyperlinkFont(font, spreadsheet));

        fonts.SetAttributeValue("count", fonts.Elements(spreadsheet + "font").Count());
        cellXfs.SetAttributeValue("count", cellXfs.Elements(spreadsheet + "xf").Count());
        SaveXmlEntry(archive, "xl/styles.xml", document);
        return new OutputStyles(boldTextStyleId, hyperlinkStyleId);
    }

    private static int EnsureStyleVariant(XElement fonts, XElement cellXfs, int baseStyleId, Action<XElement> mutateFont)
    {
        XNamespace spreadsheet = SpreadsheetNamespace;
        var styles = cellXfs.Elements(spreadsheet + "xf").ToList();
        var baseStyle = baseStyleId >= 0 && baseStyleId < styles.Count ? styles[baseStyleId] : styles[0];
        var fontId = int.TryParse(baseStyle.Attribute("fontId")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedFontId)
            ? parsedFontId
            : 0;
        var sourceFonts = fonts.Elements(spreadsheet + "font").ToList();
        var font = fontId >= 0 && fontId < sourceFonts.Count
            ? new XElement(sourceFonts[fontId])
            : new XElement(spreadsheet + "font");
        mutateFont(font);

        var existingFont = sourceFonts.FindIndex(candidate => XNode.DeepEquals(candidate, font));
        if (existingFont < 0)
        {
            existingFont = sourceFonts.Count;
            fonts.Add(font);
        }

        var variant = new XElement(baseStyle);
        variant.SetAttributeValue("fontId", existingFont);
        var existingStyle = styles.FindIndex(candidate => XNode.DeepEquals(candidate, variant));
        if (existingStyle >= 0) return existingStyle;

        var styleId = styles.Count;
        cellXfs.Add(variant);
        return styleId;
    }

    private static void EnsureBoldFont(XElement font, XNamespace spreadsheet)
    {
        if (font.Element(spreadsheet + "b") is null)
            font.AddFirst(new XElement(spreadsheet + "b"));
    }

    private static void EnsureHyperlinkFont(XElement font, XNamespace spreadsheet)
    {
        EnsureBoldFont(font, spreadsheet);
        if (font.Element(spreadsheet + "u") is null)
        {
            var anchor = font.Elements().FirstOrDefault(x => x.Name == spreadsheet + "sz" || x.Name == spreadsheet + "color" || x.Name == spreadsheet + "name");
            if (anchor is null) font.Add(new XElement(spreadsheet + "u"));
            else anchor.AddBeforeSelf(new XElement(spreadsheet + "u"));
        }

        var color = font.Element(spreadsheet + "color");
        if (color is null)
        {
            var size = font.Element(spreadsheet + "sz");
            color = new XElement(spreadsheet + "color");
            if (size is null) font.Add(color);
            else size.AddAfterSelf(color);
        }
        color.RemoveAttributes();
        color.SetAttributeValue("rgb", "FF267EF0");
    }

    private static int? GetStyleId(XElement? cell)
        => int.TryParse(cell?.Attribute("s")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var styleId)
            ? styleId
            : null;

    private static void UpsertEmailHyperlink(ZipArchive archive, string worksheetPath, XElement worksheetRoot, int rowNumber, string email)
    {
        XNamespace spreadsheet = SpreadsheetNamespace;
        XNamespace officeRelationships = OfficeRelationshipsNamespace;
        XNamespace packageRelationships = PackageRelationshipsNamespace;
        var cellReference = CellReference(6, rowNumber);
        var hyperlinks = worksheetRoot.Element(spreadsheet + "hyperlinks");
        if (hyperlinks is null)
        {
            hyperlinks = new XElement(spreadsheet + "hyperlinks");
            var anchor = worksheetRoot.Elements().FirstOrDefault(element => element.Name.LocalName is
                "printOptions" or "pageMargins" or "pageSetup" or "headerFooter" or "rowBreaks" or "colBreaks" or
                "drawing" or "legacyDrawing" or "legacyDrawingHF" or "picture" or "oleObjects" or "controls" or
                "webPublishItems" or "tableParts" or "extLst");
            if (anchor is null) worksheetRoot.Add(hyperlinks);
            else anchor.AddBeforeSelf(hyperlinks);
        }

        var oldHyperlink = hyperlinks.Elements(spreadsheet + "hyperlink")
            .SingleOrDefault(element => string.Equals(element.Attribute("ref")?.Value, cellReference, StringComparison.OrdinalIgnoreCase));
        var oldRelationshipId = oldHyperlink?.Attribute(officeRelationships + "id")?.Value;
        oldHyperlink?.Remove();

        var relationshipPath = RelationshipPath(worksheetPath);
        var relationshipEntry = archive.GetEntry(relationshipPath);
        var relationships = relationshipEntry is null
            ? new XDocument(new XElement(packageRelationships + "Relationships"))
            : LoadXmlEntry(archive, relationshipPath);
        var relationshipRoot = relationships.Root ?? throw new InvalidDataException("Excel 工作表关系 XML 无根节点。");

        if (!string.IsNullOrWhiteSpace(oldRelationshipId)
            && !hyperlinks.Elements(spreadsheet + "hyperlink").Any(element => string.Equals(element.Attribute(officeRelationships + "id")?.Value, oldRelationshipId, StringComparison.Ordinal)))
        {
            relationshipRoot.Elements(packageRelationships + "Relationship")
                .Where(element => string.Equals(element.Attribute("Id")?.Value, oldRelationshipId, StringComparison.Ordinal))
                .Remove();
        }

        if (!string.IsNullOrWhiteSpace(email))
        {
            var relationshipId = NextRelationshipId(relationshipRoot, packageRelationships);
            relationshipRoot.Add(new XElement(
                packageRelationships + "Relationship",
                new XAttribute("Id", relationshipId),
                new XAttribute("Type", HyperlinkRelationshipType),
                new XAttribute("Target", "mailto:" + email.Trim()),
                new XAttribute("TargetMode", "External")));
            hyperlinks.Add(new XElement(
                spreadsheet + "hyperlink",
                new XAttribute("ref", cellReference),
                new XAttribute(officeRelationships + "id", relationshipId),
                new XAttribute("display", email.Trim()),
                new XAttribute("tooltip", "mailto:" + email.Trim())));
        }

        SaveXmlEntry(archive, relationshipPath, relationships);
    }

    private static string NextRelationshipId(XElement relationshipRoot, XNamespace packageRelationships)
    {
        var max = relationshipRoot.Elements(packageRelationships + "Relationship")
            .Select(element => element.Attribute("Id")?.Value)
            .Select(value => value is not null && value.StartsWith("rId", StringComparison.Ordinal)
                && int.TryParse(value[3..], NumberStyles.Integer, CultureInfo.InvariantCulture, out var number) ? number : 0)
            .DefaultIfEmpty(0)
            .Max();
        return $"rId{max + 1}";
    }

    private static string RelationshipPath(string worksheetPath)
    {
        var separator = worksheetPath.LastIndexOf('/');
        return separator < 0
            ? "_rels/" + worksheetPath + ".rels"
            : worksheetPath[..(separator + 1)] + "_rels/" + worksheetPath[(separator + 1)..] + ".rels";
    }

    private static void SaveXmlEntry(ZipArchive archive, string path, XDocument document)
    {
        archive.GetEntry(path)?.Delete();
        var replacement = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var output = replacement.Open();
        document.Save(output, System.Xml.Linq.SaveOptions.DisableFormatting);
    }

    private static void SetNumeric(XElement row, XNamespace spreadsheet, int column, decimal value, XElement? styleSource, int rowNumber)
    {
        var cell = GetOrCreateCell(row, spreadsheet, column, styleSource, rowNumber);
        ClearCellContent(cell, spreadsheet);
        cell.SetAttributeValue("t", null);
        AddCellContent(cell, spreadsheet, new XElement(spreadsheet + "v", value.ToString("0.############################", CultureInfo.InvariantCulture)));
    }

    private static void ClearCell(XElement row, XNamespace spreadsheet, int column, XElement? styleSource, int rowNumber)
    {
        var cell = GetOrCreateCell(row, spreadsheet, column, styleSource, rowNumber);
        ClearCellContent(cell, spreadsheet);
        cell.SetAttributeValue("t", null);
    }

    private static void ClearCellContent(XElement cell, XNamespace spreadsheet)
    {
        cell.SetAttributeValue("t", null);
        cell.Element(spreadsheet + "f")?.Remove();
        cell.Element(spreadsheet + "v")?.Remove();
        cell.Element(spreadsheet + "is")?.Remove();
    }

    private static void AddCellContent(XElement cell, XNamespace spreadsheet, XElement content)
    {
        var extensionList = cell.Element(spreadsheet + "extLst");
        if (extensionList is null) cell.Add(content);
        else extensionList.AddBeforeSelf(content);
    }

    private static XElement GetOrCreateCell(XElement row, XNamespace spreadsheet, int column, XElement? styleSource, int rowNumber)
    {
        var existing = FindCell(row, spreadsheet, column);
        if (existing is not null) return existing;

        var source = styleSource is null ? null : FindCell(styleSource, spreadsheet, column);
        var cell = source is null ? new XElement(spreadsheet + "c") : new XElement(source);
        cell.SetAttributeValue("r", CellReference(column, rowNumber));
        cell.SetAttributeValue("t", null);
        cell.RemoveNodes();

        var next = row.Elements(spreadsheet + "c").FirstOrDefault(x => ColumnNumber(x) > column);
        if (next is null) row.Add(cell);
        else next.AddBeforeSelf(cell);
        return cell;
    }

    private static XElement? FindCell(XElement? row, XNamespace spreadsheet, int column)
        => row?.Elements(spreadsheet + "c").FirstOrDefault(x => ColumnNumber(x) == column);

    private static int RowNumber(XElement row)
        => int.TryParse(row.Attribute("r")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number) ? number : 0;

    private static int ColumnNumber(XElement cell)
    {
        var reference = cell.Attribute("r")?.Value ?? string.Empty;
        var number = 0;
        foreach (var character in reference.TakeWhile(char.IsLetter))
            number = number * 26 + (char.ToUpperInvariant(character) - 'A' + 1);
        return number;
    }

    private static string CellReference(int column, int row)
    {
        var letters = string.Empty;
        for (var current = column; current > 0; current = (current - 1) / 26)
            letters = (char)('A' + (current - 1) % 26) + letters;
        return $"{letters}{row}";
    }

    private static void ExpandDimension(XElement root, XNamespace spreadsheet, int targetRow, int targetColumn)
    {
        var dimension = root.Element(spreadsheet + "dimension");
        if (dimension is null)
        {
            dimension = new XElement(spreadsheet + "dimension");
            var before = root.Elements().FirstOrDefault(x => x.Name == spreadsheet + "sheetViews" || x.Name == spreadsheet + "sheetData");
            if (before is null) root.AddFirst(dimension);
            else before.AddBeforeSelf(dimension);
            dimension.SetAttributeValue("ref", $"A1:{CellReference(targetColumn, targetRow)}");
            return;
        }

        var range = (dimension.Attribute("ref")?.Value ?? "A1").Split(':');
        var start = ParseCellReference(range[0]);
        var end = ParseCellReference(range.Length > 1 ? range[1] : range[0]);
        end = (Math.Max(end.Column, targetColumn), Math.Max(end.Row, targetRow));
        dimension.SetAttributeValue("ref", $"{CellReference(start.Column, start.Row)}:{CellReference(end.Column, end.Row)}");
    }

    private static (int Column, int Row) ParseCellReference(string reference)
    {
        var column = 0;
        var index = 0;
        while (index < reference.Length && char.IsLetter(reference[index]))
        {
            column = column * 26 + (char.ToUpperInvariant(reference[index]) - 'A' + 1);
            index++;
        }
        return (column, int.TryParse(reference[index..], NumberStyles.Integer, CultureInfo.InvariantCulture, out var row) ? row : 1);
    }

    private static string ResolveWorksheetPath(ZipArchive archive, string worksheetName)
    {
        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace officeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        var workbook = LoadXmlEntry(archive, "xl/workbook.xml");
        var relationships = LoadXmlEntry(archive, "xl/_rels/workbook.xml.rels");
        var sheet = workbook.Root?.Element(spreadsheet + "sheets")?.Elements(spreadsheet + "sheet").SingleOrDefault(x => x.Attribute("name")?.Value == worksheetName)
            ?? throw new InvalidOperationException($"工作表不存在：{worksheetName}");
        var relationshipId = sheet.Attribute(officeRelationships + "id")?.Value
            ?? throw new InvalidDataException($"工作表关系不存在：{worksheetName}");
        var target = relationships.Root?.Elements(packageRelationships + "Relationship").SingleOrDefault(x => x.Attribute("Id")?.Value == relationshipId)?.Attribute("Target")?.Value
            ?? throw new InvalidDataException($"工作表目标不存在：{worksheetName}");
        target = target.Replace('\\', '/');
        if (target.StartsWith("/", StringComparison.Ordinal)) return target.TrimStart('/');
        if (target.StartsWith("../", StringComparison.Ordinal)) return "xl/" + target[3..];
        return "xl/" + target.TrimStart('/');
    }

    private static XDocument LoadXmlEntry(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path) ?? throw new InvalidDataException($"Excel 文件缺少 XML 部件：{path}");
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }
}

public sealed class DpapiCredentialStore(string directory)
{
    public async Task SaveAsync(string account, string password, CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(directory);
        var cipher = ProtectedData.Protect(Encoding.UTF8.GetBytes(password), Entropy(account), DataProtectionScope.CurrentUser);
        var path = PathFor(account);
        var tempPath = Path.Combine(directory, $".{Path.GetFileName(path)}.{Guid.NewGuid():N}.tmp");
        try
        {
            await File.WriteAllBytesAsync(tempPath, cipher, cancellationToken);
            using (var durable = new FileStream(tempPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                durable.Flush(true);

            if (File.Exists(path)) File.Replace(tempPath, path, null, true);
            else File.Move(tempPath, path);
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    public async Task<string?> LoadAsync(string account, CancellationToken cancellationToken = default)
    {
        var path = PathFor(account);
        if (!File.Exists(path)) return null;
        var cipher = await File.ReadAllBytesAsync(path, cancellationToken);
        var clear = ProtectedData.Unprotect(cipher, Entropy(account), DataProtectionScope.CurrentUser);
        return Encoding.UTF8.GetString(clear);
    }

    public Task DeleteAsync(string account, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var path = PathFor(account);
        if (File.Exists(path)) File.Delete(path);
        return Task.CompletedTask;
    }

    private static byte[] Entropy(string account) => SHA256.HashData(Encoding.UTF8.GetBytes($"InvoiceMailAssistant|{account.Trim().ToLowerInvariant()}"));
    private string PathFor(string account) => Path.Combine(directory, Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(account.Trim().ToLowerInvariant()))) + ".cred");
}

public static class StartupManager
{
    private const string RunKey = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
    private const string ValueName = "InvoiceMailAssistant";

    public static void SetEnabled(bool enabled)
    {
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(RunKey, true)
                ?? throw new InvalidOperationException("无法打开当前用户的开机启动注册表项。");
            if (!enabled)
            {
                key.DeleteValue(ValueName, false);
                return;
            }

            var processPath = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(processPath)) throw new InvalidOperationException("无法确定程序路径。");
            key.SetValue(ValueName, $"\"{processPath}\"");
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or SecurityException or IOException)
        {
            throw new InvalidOperationException("开机启动设置失败。", ex);
        }
    }
}
