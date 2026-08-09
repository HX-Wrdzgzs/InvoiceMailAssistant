using ClosedXML.Excel;
using InvoiceMailAssistant.App;
using Xunit;

namespace InvoiceMailAssistant.Tests;

public sealed class CrashRecoveryTests
{
    [Fact]
    public async Task PendingRowIsReusedAfterRestartBeforeExcelWrite()
    {
        var files = TestFiles.Create();
        try
        {
            TestFiles.CreateWorkbook(files.ExcelPath);
            var repository = new SqliteInvoiceRepository(files.DatabasePath);
            await repository.InitializeAsync();
            var application = TestFiles.CreateApplication(10);
            var id = await repository.TryInsertAsync(application, "recovery-a");
            Assert.NotNull(id);
            await repository.UpdateStatusAsync(id!.Value, ProcessingStatus.PendingExcel, excelRow: 10);

            var recovered = Assert.Single(await repository.GetPendingExcelAsync());
            var writer = new ExcelWriter();
            recovered.ExcelRow = writer.ResolveTargetRow(recovered, files.ExcelPath, "中外运");
            Assert.Equal(10, recovered.ExcelRow);
            var row = await writer.WriteAsync(recovered, files.ExcelPath, "中外运");
            await repository.UpdateStatusAsync(recovered.Id, ProcessingStatus.Completed, excelRow: row);

            Assert.Equal(ProcessingStatus.Completed, Assert.Single(await repository.GetRecentAsync(10)).ProcessingStatus);
            using var workbook = new XLWorkbook(files.ExcelPath);
            Assert.Equal("测试企业", workbook.Worksheet("中外运").Cell(10, 2).GetString());
        }
        finally
        {
            files.Delete();
        }
    }

    [Fact]
    public async Task PendingRowAlreadyWrittenIsCompletedWithoutAppending()
    {
        var files = TestFiles.Create();
        try
        {
            var application = TestFiles.CreateApplication(10);
            TestFiles.CreateWorkbook(files.ExcelPath, application);
            var repository = new SqliteInvoiceRepository(files.DatabasePath);
            await repository.InitializeAsync();
            var id = await repository.TryInsertAsync(application, "recovery-b");
            Assert.NotNull(id);
            await repository.UpdateStatusAsync(id!.Value, ProcessingStatus.PendingExcel, excelRow: 10);

            var recovered = Assert.Single(await repository.GetPendingExcelAsync());
            var writer = new ExcelWriter();
            recovered.ExcelRow = writer.ResolveTargetRow(recovered, files.ExcelPath, "中外运");
            var row = await writer.WriteAsync(recovered, files.ExcelPath, "中外运");
            await repository.UpdateStatusAsync(recovered.Id, ProcessingStatus.Completed, excelRow: row);

            using var workbook = new XLWorkbook(files.ExcelPath);
            Assert.Equal(10, row);
            Assert.True(workbook.Worksheet("中外运").Cell(11, 2).IsEmpty());
        }
        finally
        {
            files.Delete();
        }
    }

    [Fact]
    public async Task OccupiedPendingRowIsReplannedWithoutOverwrite()
    {
        var files = TestFiles.Create();
        try
        {
            TestFiles.CreateWorkbook(files.ExcelPath, manualRow10: true);
            var repository = new SqliteInvoiceRepository(files.DatabasePath);
            await repository.InitializeAsync();
            var application = TestFiles.CreateApplication(10);
            var id = await repository.TryInsertAsync(application, "recovery-c");
            Assert.NotNull(id);
            await repository.UpdateStatusAsync(id!.Value, ProcessingStatus.PendingExcel, excelRow: 10);

            var recovered = Assert.Single(await repository.GetPendingExcelAsync());
            var writer = new ExcelWriter();
            recovered.ExcelRow = writer.ResolveTargetRow(recovered, files.ExcelPath, "中外运");
            Assert.Equal(11, recovered.ExcelRow);
            var row = await writer.WriteAsync(recovered, files.ExcelPath, "中外运");
            await repository.UpdateStatusAsync(recovered.Id, ProcessingStatus.Completed, excelRow: row);

            using var workbook = new XLWorkbook(files.ExcelPath);
            var sheet = workbook.Worksheet("中外运");
            Assert.Equal("用户手工数据", sheet.Cell(10, 2).GetString());
            Assert.Equal("测试企业", sheet.Cell(11, 2).GetString());
        }
        finally
        {
            files.Delete();
        }
    }

    private sealed class TestFiles(string excelPath, string databasePath)
    {
        public string ExcelPath { get; } = excelPath;
        public string DatabasePath { get; } = databasePath;

        public static TestFiles Create()
        {
            var stem = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}");
            return new TestFiles(stem + ".xlsx", stem + ".db");
        }

        public void Delete()
        {
            foreach (var path in new[] { ExcelPath, DatabasePath, DatabasePath + "-wal", DatabasePath + "-shm" })
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        public static void CreateWorkbook(string path, InvoiceApplication? application = null, bool manualRow10 = false)
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.AddWorksheet("中外运");
            sheet.Cell(1, 1).Value = "日期";
            sheet.Cell(1, 2).Value = "企业名称";
            sheet.Cell(1, 3).Value = "信用代码";
            sheet.Cell(1, 4).Value = "开票金额";
            sheet.Cell(1, 5).Value = "到账时间";
            sheet.Cell(1, 6).Value = "邮箱";
            sheet.Cell(1, 7).Value = "是否开票";
            sheet.Cell(1, 8).Value = "发票是否已发";
            if (application is not null)
            {
                sheet.Cell(10, 2).Value = application.CompanyName;
                sheet.Cell(10, 3).Value = application.CreditCode;
                sheet.Cell(10, 4).Value = application.Amount;
                sheet.Cell(10, 6).Value = application.Email;
            }
            else if (manualRow10)
            {
                sheet.Cell(10, 2).Value = "用户手工数据";
            }
            workbook.SaveAs(path);
        }

        public static InvoiceApplication CreateApplication(int row)
            => new()
            {
                CompanyName = "测试企业",
                CreditCode = "CODE-RECOVERY",
                Amount = 100m,
                ApplyTime = new DateTime(2026, 8, 8, 9, 0, 0),
                Email = "finance@example.com",
                ExcelRow = row,
                MessageId = $"recovery-{Guid.NewGuid():N}",
                ImapUid = 1,
                UidValidity = 1,
                MailboxName = "INBOX",
                MailboxIdentity = "account@example.com",
                MailReceivedAt = DateTimeOffset.UtcNow,
                MailSubject = "subject",
                MailFrom = "sino-esign@sinotrans.com",
                NormalizedBody = "body",
                ProcessingStatus = ProcessingStatus.Parsed
            };
    }
}
