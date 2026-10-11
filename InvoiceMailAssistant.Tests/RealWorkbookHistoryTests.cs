using System.Security.Cryptography;
using ClosedXML.Excel;
using InvoiceMailAssistant.App;
using Xunit;

namespace InvoiceMailAssistant.Tests;

public sealed class RealWorkbookHistoryTests
{
    [RealWorkbookHistoryFact]
    public async Task ExistingTianjinRowsAreClaimedOneToOneWithoutAppending()
    {
        var workbookPath = Environment.GetEnvironmentVariable("INVOICE_REAL_XLSX_HISTORY_COPY");
        Assert.False(string.IsNullOrWhiteSpace(workbookPath));
        Assert.True(File.Exists(workbookPath), $"Real workbook history copy not found: {workbookPath}");

        var beforeHash = Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(workbookPath!)));
        var dataDirectory = Path.Combine(Path.GetTempPath(), $"invoice-history-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dataDirectory);
        try
        {
            var repository = new SqliteInvoiceRepository(Path.Combine(dataDirectory, "invoice-mail.db"));
            await repository.InitializeAsync();
            foreach (var index in Enumerable.Range(0, 3))
            {
                var application = new InvoiceApplication
                {
                    CompanyName = "天津安路佳供应链管理有限公司",
                    CreditCode = "91120111MA7LKBW877",
                    Amount = 300m,
                    ApplyTime = new DateTime(2026, 8, 27, 10, 15, 0),
                    Email = "xiriteng@163.com",
                    MessageId = $"<real-tianjin-{index}@example.test>",
                    ImapUid = (uint)(700 + index),
                    UidValidity = 11,
                    MailboxName = "INBOX",
                    MailboxIdentity = "real-test@example.test",
                    MailReceivedAt = new DateTimeOffset(2026, 8, 27, 10, 15 + index, 0, TimeSpan.FromHours(8)),
                    MailSubject = $"天津历史测试 {index}",
                    MailFrom = "sino-esign@sinotrans.com",
                    NormalizedBody = $"公司名称：天津安路佳供应链管理有限公司\n信用代码：91120111MA7LKBW877\n申请金额：300\n申请时间：2026-08-27 10:15\n邮箱：xiriteng@163.com\n历史测试：{index}",
                    ProcessingStatus = ProcessingStatus.Completed,
                    ExcelRow = 5997
                };
                application.Id = (await repository.TryInsertAsync(application, Deduplication.CreateFallbackHash(application)))!.Value;
            }

            var result = await new ExcelAssignmentCoordinator(repository, new ExcelWriter(), dataDirectory)
                .ReconcileAsync(workbookPath!, "中外运");

            Assert.Equal(3, result.MappedExistingRows);
            Assert.Equal(0, result.WrittenRows);
            Assert.Equal(0, result.PendingRows);
            var stored = (await repository.GetRecentAsync(10)).OrderBy(item => item.Id).ToArray();
            Assert.Equal(new int?[] { 5997, 5998, 5999 }, stored.Select(item => item.ExcelRow).ToArray());
            Assert.All(stored, item => Assert.Equal(ProcessingStatus.Completed, item.ProcessingStatus));

            using var workbook = new XLWorkbook(workbookPath!);
            var sheet = workbook.Worksheet("中外运");
            Assert.Equal("天津安路佳供应链管理有限公司", sheet.Cell(5997, 2).GetString());
            Assert.Equal("天津安路佳供应链管理有限公司", sheet.Cell(5998, 2).GetString());
            Assert.Equal("天津安路佳供应链管理有限公司", sheet.Cell(5999, 2).GetString());
            var afterHash = Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(workbookPath!)));
            Assert.Equal(beforeHash, afterHash);
        }
        finally
        {
            if (Directory.Exists(dataDirectory)) Directory.Delete(dataDirectory, recursive: true);
        }
    }
}

public sealed class RealWorkbookHistoryFactAttribute : FactAttribute
{
    public RealWorkbookHistoryFactAttribute()
    {
        if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("INVOICE_REAL_XLSX_HISTORY_COPY")))
            Skip = "Set INVOICE_REAL_XLSX_HISTORY_COPY to run the isolated history regression.";
    }
}
