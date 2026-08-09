using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using ClosedXML.Excel;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using Microsoft.Data.Sqlite;

namespace InvoiceMailAssistant.App;

public sealed class MailboxService
{
    private const string Sender = "sino-esign@sinotrans.com";
    private const string SubjectPrefix = "中外运向您提交了开票申请";

    public async Task TestConnectionAsync(string account, string password, string host, int port, CancellationToken cancellationToken)
    {
        using var client = new ImapClient();
        await client.ConnectAsync(host, port, true, cancellationToken);
        await client.AuthenticateAsync(account, password, cancellationToken);
        await client.DisconnectAsync(true, cancellationToken);
    }

    public async Task<IReadOnlyList<MailEnvelope>> FetchCandidateMessagesAsync(string account, string password, string host, int port, DateTimeOffset monitorFromUtc, int maxMessages, CancellationToken cancellationToken)
    {
        using var client = new ImapClient();
        await client.ConnectAsync(host, port, true, cancellationToken);
        await client.AuthenticateAsync(account, password, cancellationToken);
        var inbox = client.Inbox;
        await inbox.OpenAsync(FolderAccess.ReadOnly, cancellationToken);

        var imapDateFloor = monitorFromUtc.UtcDateTime.Date.AddDays(-1);
        var query = SearchQuery.FromContains(Sender)
            .And(SearchQuery.SubjectContains(SubjectPrefix))
            .And(SearchQuery.DeliveredAfter(imapDateFloor));
        var uids = await inbox.SearchAsync(query, cancellationToken);
        var selected = uids.OrderByDescending(x => x.Id).Take(Math.Clamp(maxMessages, 1, 500)).OrderBy(x => x.Id).ToArray();
        var uidValidity = inbox.UidValidity;
        var result = new List<MailEnvelope>(selected.Length);

        foreach (var uid in selected)
        {
            var message = await inbox.GetMessageAsync(uid, cancellationToken);
            var from = message.From.Mailboxes.FirstOrDefault()?.Address?.Trim().ToLowerInvariant() ?? string.Empty;
            var subject = message.Subject?.Trim() ?? string.Empty;
            if (!string.Equals(from, Sender, StringComparison.OrdinalIgnoreCase) || !subject.StartsWith(SubjectPrefix, StringComparison.Ordinal))
                continue;
            if (message.Date.ToUniversalTime() < monitorFromUtc.ToUniversalTime())
                continue;

            result.Add(new MailEnvelope(
                uid.Id,
                message.MessageId ?? string.Empty,
                from,
                subject,
                message.Date,
                message.TextBody ?? message.HtmlBody ?? string.Empty,
                uidValidity,
                inbox.FullName));
        }

        await client.DisconnectAsync(true, cancellationToken);
        return result;
    }
}

public sealed class SqliteInvoiceRepository(string databasePath)
{
    private readonly string _connectionString = new SqliteConnectionStringBuilder { DataSource = databasePath, Pooling = false }.ToString();

    public async Task InitializeAsync(CancellationToken cancellationToken = default)
    {
        var directory = Path.GetDirectoryName(databasePath);
        if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
        await using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync(cancellationToken);
        var command = connection.CreateCommand();
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
        command = connection.CreateCommand();
        command.CommandText = "DROP INDEX IF EXISTS ux_invoice_uid; CREATE UNIQUE INDEX IF NOT EXISTS ux_invoice_uid ON invoice_applications(mailbox_identity, mailbox_name, uid_validity, imap_uid);";
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<bool> ExistsAsync(string messageId, uint uid, string mailboxIdentity, string fallbackHash, CancellationToken cancellationToken = default)
        => await ExistsAsync(messageId, uid, mailboxIdentity, "INBOX", 0, fallbackHash, cancellationToken);

    public async Task<bool> ExistsAsync(string messageId, uint uid, string mailboxIdentity, string mailboxName, uint uidValidity, string fallbackHash, CancellationToken cancellationToken = default)
    {
        await using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync(cancellationToken);
        var command = connection.CreateCommand();
        command.CommandText = "SELECT 1 FROM invoice_applications WHERE (message_id <> '' AND message_id = $messageId) OR (mailbox_identity = $mailbox AND mailbox_name = $mailboxName AND uid_validity = $uidValidity AND imap_uid = $uid) OR fallback_hash = $hash LIMIT 1;";
        command.Parameters.AddWithValue("$messageId", messageId);
        command.Parameters.AddWithValue("$mailbox", mailboxIdentity);
        command.Parameters.AddWithValue("$mailboxName", mailboxName);
        command.Parameters.AddWithValue("$uidValidity", (long)uidValidity);
        command.Parameters.AddWithValue("$uid", (long)uid);
        command.Parameters.AddWithValue("$hash", fallbackHash);
        return await command.ExecuteScalarAsync(cancellationToken) is not null;
    }

    public async Task<long?> TryInsertAsync(InvoiceApplication app, string fallbackHash, CancellationToken cancellationToken = default)
    {
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
        await using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync(cancellationToken);
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
        await using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync(cancellationToken);
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
        await using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync(cancellationToken);
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
        await using var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync(cancellationToken);
        var command = connection.CreateCommand();
        command.CommandText = "SELECT * FROM invoice_applications WHERE processing_status IN (2, 5) ORDER BY id;";
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken)) result.Add(Map(reader));
        return result;
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

public sealed class ExcelWriter
{
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
        }

        return Math.Max(lastUsed + 1, 2);
    }

    public Task<int> WriteAsync(InvoiceApplication application, string filePath, string worksheetName, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        EnsureWritable(filePath);

        var directory = Path.GetDirectoryName(filePath)!;
        var tempPath = Path.Combine(directory, $".{Path.GetFileName(filePath)}.{Guid.NewGuid():N}.tmp.xlsx");
        var originalWorksheetParts = ReadPreservedWorksheetParts(filePath, worksheetName);
        File.Copy(filePath, tempPath, true);

        try
        {
            var row = application.ExcelRow ?? throw new InvalidOperationException("尚未为 Excel 写入预留目标行。");
            using (var workbook = new XLWorkbook(tempPath))
            {
                if (!workbook.Worksheets.TryGetWorksheet(worksheetName, out var worksheet))
                    throw new InvalidOperationException($"工作表不存在：{worksheetName}");

                if (RowMatches(worksheet, row, application))
                    return Task.FromResult(row);

                if (!RowIsEmpty(worksheet, row))
                    throw new ExcelRowOccupiedException($"Excel 第 {row} 行已被其他数据占用，请重新规划目标行。");

                if (row > 2) CopyRowStyle(worksheet, row - 1, row);

                var previousDate = FindPreviousApplicationDate(worksheet, row - 1, application.ApplyTime.Year);
                if (previousDate?.Date != application.ApplyTime.Date)
                    worksheet.Cell(row, 1).Value = $"{application.ApplyTime.Month}.{application.ApplyTime.Day}";
                else
                    worksheet.Cell(row, 1).Clear(XLClearOptions.Contents);

                worksheet.Cell(row, 2).Value = application.CompanyName;
                worksheet.Cell(row, 3).SetValue(application.CreditCode);
                worksheet.Cell(row, 3).Style.NumberFormat.Format = "@";
                worksheet.Cell(row, 4).Value = application.Amount;
                worksheet.Cell(row, 6).SetValue(application.Email);
                worksheet.Cell(row, 6).Style.NumberFormat.Format = "@";
                workbook.SaveAs(tempPath);
            }

            RestorePreservedWorksheetParts(tempPath, worksheetName, originalWorksheetParts);

            using (var verify = new XLWorkbook(tempPath))
            {
                var verifySheet = verify.Worksheet(worksheetName);
                if (!RowMatches(verifySheet, row, application))
                    throw new InvalidDataException("Excel 临时文件校验失败：目标行内容不一致。");
            }

            File.Move(tempPath, filePath, true);
            return Task.FromResult(row);
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    private static void EnsureWritable(string filePath)
    {
        if (!File.Exists(filePath)) throw new FileNotFoundException("Excel 登记表不存在。", filePath);
        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        catch (IOException ex)
        {
            throw new IOException("Excel 登记表正在被占用，申请将保留在等待队列。", ex);
        }
    }

    private static bool RowIsEmpty(IXLWorksheet worksheet, int row)
        => Enumerable.Range(1, 8).All(column => worksheet.Cell(row, column).IsEmpty());

    private static bool RowMatches(IXLWorksheet worksheet, int row, InvoiceApplication application)
    {
        if (!string.Equals(worksheet.Cell(row, 2).GetString().Trim(), application.CompanyName.Trim(), StringComparison.Ordinal)) return false;
        if (!string.Equals(worksheet.Cell(row, 3).GetString().Trim(), application.CreditCode.Trim(), StringComparison.Ordinal)) return false;
        if (!decimal.TryParse(worksheet.Cell(row, 4).GetString(), out var amount) || amount != application.Amount) return false;
        return string.Equals(worksheet.Cell(row, 6).GetString().Trim(), application.Email.Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private static void CopyRowStyle(IXLWorksheet ws, int sourceRow, int targetRow)
    {
        for (var col = 1; col <= 8; col++)
            ws.Cell(targetRow, col).Style = ws.Cell(sourceRow, col).Style;
        ws.Row(targetRow).Height = ws.Row(sourceRow).Height;
    }

    private static DateTime? FindPreviousApplicationDate(IXLWorksheet ws, int row, int assumedYear)
    {
        for (var current = row; current >= 2; current--)
        {
            var text = ws.Cell(current, 1).GetString().Trim();
            if (string.IsNullOrWhiteSpace(text)) continue;
            var parts = text.Split('.');
            if (parts.Length == 2 && int.TryParse(parts[0], out var month) && int.TryParse(parts[1], out var day))
            {
                try { return new DateTime(assumedYear, month, day); }
                catch (ArgumentOutOfRangeException) { return null; }
            }
        }
        return null;
    }

    private sealed record PreservedWorksheetParts(
        XElement? FreezePane,
        XElement? PrintOptions,
        XElement? PageMargins,
        XElement? PageSetup,
        XElement? HeaderFooter,
        XElement? RowBreaks,
        XElement? ColumnBreaks);

    private static PreservedWorksheetParts ReadPreservedWorksheetParts(string filePath, string worksheetName)
    {
        using var archive = ZipFile.OpenRead(filePath);
        var worksheetPath = ResolveWorksheetPath(archive, worksheetName);
        using var stream = archive.GetEntry(worksheetPath)?.Open();
        if (stream is null) throw new InvalidDataException($"工作表 XML 不存在：{worksheetName}");
        var document = XDocument.Load(stream);
        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var root = document.Root ?? throw new InvalidDataException($"工作表 XML 无根节点：{worksheetName}");
        var sheetView = root.Element(spreadsheet + "sheetViews")?.Elements(spreadsheet + "sheetView").FirstOrDefault();
        return new PreservedWorksheetParts(
            CloneElement(sheetView?.Element(spreadsheet + "pane")),
            CloneElement(root.Element(spreadsheet + "printOptions")),
            CloneElement(root.Element(spreadsheet + "pageMargins")),
            CloneElement(root.Element(spreadsheet + "pageSetup")),
            CloneElement(root.Element(spreadsheet + "headerFooter")),
            CloneElement(root.Element(spreadsheet + "rowBreaks")),
            CloneElement(root.Element(spreadsheet + "colBreaks")));
    }

    private static void RestorePreservedWorksheetParts(string filePath, string worksheetName, PreservedWorksheetParts originalParts)
    {
        using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
        var worksheetPath = ResolveWorksheetPath(archive, worksheetName);
        var entry = archive.GetEntry(worksheetPath) ?? throw new InvalidDataException($"工作表 XML 不存在：{worksheetName}");
        XDocument document;
        using (var stream = entry.Open()) document = XDocument.Load(stream);

        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var sheetView = document.Root?.Element(spreadsheet + "sheetViews")?.Elements(spreadsheet + "sheetView").FirstOrDefault();
        if (sheetView is not null)
        {
            ReplaceChild(sheetView, spreadsheet + "pane", originalParts.FreezePane);
        }

        var root = document.Root ?? throw new InvalidDataException($"工作表 XML 无根节点：{worksheetName}");
        ReplaceChild(root, spreadsheet + "printOptions", originalParts.PrintOptions);
        ReplaceChild(root, spreadsheet + "pageMargins", originalParts.PageMargins);
        ReplaceChild(root, spreadsheet + "pageSetup", originalParts.PageSetup);
        ReplaceChild(root, spreadsheet + "headerFooter", originalParts.HeaderFooter);
        ReplaceChild(root, spreadsheet + "rowBreaks", originalParts.RowBreaks);
        ReplaceChild(root, spreadsheet + "colBreaks", originalParts.ColumnBreaks);

        entry.Delete();
        var replacement = archive.CreateEntry(worksheetPath, CompressionLevel.Optimal);
        using var output = replacement.Open();
        document.Save(output, System.Xml.Linq.SaveOptions.DisableFormatting);
    }

    private static XElement? CloneElement(XElement? element) => element is null ? null : new XElement(element);

    private static void ReplaceChild(XElement parent, XName name, XElement? original)
    {
        var current = parent.Element(name);
        if (current is not null)
        {
            if (original is null) current.Remove();
            else current.ReplaceWith(new XElement(original));
        }
        else if (original is not null)
        {
            parent.Add(new XElement(original));
        }
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
        await File.WriteAllBytesAsync(PathFor(account), cipher, cancellationToken);
    }

    public async Task<string?> LoadAsync(string account, CancellationToken cancellationToken = default)
    {
        var path = PathFor(account);
        if (!File.Exists(path)) return null;
        var cipher = await File.ReadAllBytesAsync(path, cancellationToken);
        var clear = ProtectedData.Unprotect(cipher, Entropy(account), DataProtectionScope.CurrentUser);
        return Encoding.UTF8.GetString(clear);
    }

    private static byte[] Entropy(string account) => SHA256.HashData(Encoding.UTF8.GetBytes($"InvoiceMailAssistant|{account.Trim().ToLowerInvariant()}"));
    private string PathFor(string account) => Path.Combine(directory, Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(account.Trim().ToLowerInvariant()))) + ".cred");
}
