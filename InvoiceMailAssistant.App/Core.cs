using System.Globalization;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace InvoiceMailAssistant.App;

public enum ProcessingStatus
{
    Discovered,
    Parsed,
    PendingExcel,
    Completed,
    ParseFailed,
    ExcelFailed,
    MailFailed
}

public sealed class InvoiceApplication
{
    public long Id { get; set; }
    public string CompanyName { get; set; } = string.Empty;
    public string CreditCode { get; set; } = string.Empty;
    public decimal Amount { get; set; }
    public DateTime ApplyTime { get; set; }
    public string InvoiceType { get; set; } = string.Empty;
    public string Recipient { get; set; } = string.Empty;
    public string Phone { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public string Remark { get; set; } = string.Empty;
    public string MessageId { get; set; } = string.Empty;
    public uint ImapUid { get; set; }
    public uint UidValidity { get; set; }
    public string MailboxName { get; set; } = "INBOX";
    public string MailboxIdentity { get; set; } = string.Empty;
    public DateTimeOffset MailReceivedAt { get; set; }
    public string MailSubject { get; set; } = string.Empty;
    public string MailFrom { get; set; } = string.Empty;
    public string NormalizedBody { get; set; } = string.Empty;
    public ProcessingStatus ProcessingStatus { get; set; } = ProcessingStatus.Discovered;
    public int? ExcelRow { get; set; }
    public DateTimeOffset CreatedAt { get; set; } = DateTimeOffset.Now;
    public DateTimeOffset UpdatedAt { get; set; } = DateTimeOffset.Now;
    public string ErrorMessage { get; set; } = string.Empty;
}

public sealed record MailEnvelope(
    uint Uid,
    string MessageId,
    string FromAddress,
    string Subject,
    DateTimeOffset ReceivedAt,
    string BodyText,
    uint UidValidity = 0,
    string MailboxName = "INBOX",
    string? FetchError = null);

public sealed record MailboxMessage(
    uint Uid,
    uint UidValidity,
    string MailboxName,
    DateTimeOffset InternalDate,
    IReadOnlyList<string> FromAddresses,
    string Subject,
    string MessageId,
    string BodyText,
    string? FetchError = null);

public interface IMailboxSession : IAsyncDisposable
{
    bool IsConnected { get; }
    Task<IReadOnlyList<MailboxMessage>> FetchCandidateMessagesAsync(DateTimeOffset monitorFromUtc, CancellationToken cancellationToken);
}

public interface IMailboxSessionFactory
{
    Task<IMailboxSession> ConnectAsync(string account, string password, string host, int port, CancellationToken cancellationToken);
}

public sealed record ParseResult(bool Success, InvoiceApplication? Application, IReadOnlyList<string> MissingFields, string? Error)
{
    public static ParseResult Ok(InvoiceApplication application) => new(true, application, Array.Empty<string>(), null);
    public static ParseResult Fail(IReadOnlyList<string> missingFields, string error) => new(false, null, missingFields, error);
}

public sealed partial class InvoiceParser
{
    private static readonly string[] RequiredFields = ["公司名称", "信用代码", "申请金额", "申请时间", "邮箱"];

    public ParseResult Parse(MailEnvelope mail, string mailboxIdentity)
    {
        var normalized = NormalizeBody(mail.BodyText);
        var values = ParseLabelValues(normalized);
        var missing = RequiredFields.Where(x => !values.TryGetValue(x, out var value) || string.IsNullOrWhiteSpace(value)).ToArray();
        if (missing.Length > 0)
            return ParseResult.Fail(missing, $"缺少必填字段：{string.Join("、", missing)}");

        if (!TryParseAmount(values["申请金额"], out var amount))
            return ParseResult.Fail(["申请金额"], $"无法解析申请金额：{values["申请金额"]}");

        if (!DateTime.TryParseExact(values["申请时间"].Trim(), ["yyyy-MM-dd HH:mm", "yyyy-M-d H:mm", "yyyy-MM-dd HH:mm:ss"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var applyTime))
            return ParseResult.Fail(["申请时间"], $"无法解析申请时间：{values["申请时间"]}");

        var email = NormalizeEmail(values["邮箱"]);
        if (!EmailRegex().IsMatch(email))
            return ParseResult.Fail(["邮箱"], $"无法解析邮箱地址：{values["邮箱"]}");

        return ParseResult.Ok(new InvoiceApplication
        {
            CompanyName = values["公司名称"].Trim(),
            CreditCode = values["信用代码"].Trim(),
            Amount = amount,
            ApplyTime = applyTime,
            InvoiceType = Get(values, "开票方式"),
            Recipient = Get(values, "收件人"),
            Phone = Get(values, "联系电话"),
            Address = Get(values, "寄送地址"),
            Email = email,
            Remark = Get(values, "开票备注"),
            MessageId = mail.MessageId.Trim(),
            ImapUid = mail.Uid,
            UidValidity = mail.UidValidity,
            MailboxName = mail.MailboxName,
            MailboxIdentity = mailboxIdentity,
            MailReceivedAt = mail.ReceivedAt,
            MailSubject = mail.Subject,
            MailFrom = mail.FromAddress,
            NormalizedBody = normalized,
            ProcessingStatus = ProcessingStatus.Parsed
        });
    }

    public static string NormalizeBody(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) return string.Empty;
        var decoded = WebUtility.HtmlDecode(input).Replace("\r\n", "\n").Replace('\r', '\n');
        decoded = BlockTagRegex().Replace(decoded, "\n");
        decoded = BrRegex().Replace(decoded, "\n");
        decoded = TagRegex().Replace(decoded, string.Empty);
        return string.Join('\n', decoded.Split('\n').Select(x => x.Trim()).Where(x => x.Length > 0));
    }

    private static Dictionary<string, string> ParseLabelValues(string body)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var line in body.Split('\n'))
        {
            var index = line.IndexOfAny(['：', ':']);
            if (index <= 0) continue;
            var label = line[..index].Trim();
            var value = line[(index + 1)..].Trim();
            if (!result.ContainsKey(label)) result[label] = value;
        }
        return result;
    }

    private static string Get(Dictionary<string, string> values, string key) => values.TryGetValue(key, out var value) ? value.Trim() : string.Empty;

    private static bool TryParseAmount(string raw, out decimal amount)
    {
        var cleaned = raw.Replace("￥", string.Empty).Replace("¥", string.Empty).Replace("元", string.Empty).Replace(",", string.Empty).Trim();
        return decimal.TryParse(cleaned, NumberStyles.Number, CultureInfo.InvariantCulture, out amount);
    }

    private static string NormalizeEmail(string raw)
    {
        var value = raw.Trim();
        if (value.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)) value = value[7..];
        var match = EmailRegex().Match(value);
        return match.Success ? match.Value : value;
    }

    [GeneratedRegex("<br\\s*/?>", RegexOptions.IgnoreCase)]
    private static partial Regex BrRegex();

    [GeneratedRegex("</?(?:div|p|li|tr|td|h[1-6])\\b[^>]*>", RegexOptions.IgnoreCase)]
    private static partial Regex BlockTagRegex();

    [GeneratedRegex("<[^>]+>")]
    private static partial Regex TagRegex();

    [GeneratedRegex("[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,}", RegexOptions.IgnoreCase)]
    private static partial Regex EmailRegex();
}

public static class Deduplication
{
    public static string CreateFallbackHash(InvoiceApplication app)
    {
        var source = string.Join("|", [
            app.MailboxIdentity.Trim().ToLowerInvariant(),
            app.MailboxName.Trim().ToLowerInvariant(),
            app.MailFrom.Trim().ToLowerInvariant(),
            app.MailSubject.Trim(),
            app.MailReceivedAt.ToUniversalTime().ToString("O"),
            app.CreditCode.Trim().ToUpperInvariant(),
            app.ApplyTime.ToString("O"),
            app.Amount.ToString("0.####", CultureInfo.InvariantCulture),
            app.NormalizedBody.Trim()
        ]);
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(source)));
    }

    public static string CreateLegacyFallbackHash(InvoiceApplication app)
    {
        var source = $"{app.CreditCode.Trim().ToUpperInvariant()}|{app.ApplyTime:O}|{app.Amount:0.####}|{app.MailFrom.Trim().ToLowerInvariant()}";
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(source)));
    }
}
