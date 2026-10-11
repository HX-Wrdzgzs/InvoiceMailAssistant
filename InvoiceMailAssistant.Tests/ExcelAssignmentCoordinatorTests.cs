using ClosedXML.Excel;
using InvoiceMailAssistant.App;
using Xunit;

namespace InvoiceMailAssistant.Tests;

public sealed class ExcelAssignmentCoordinatorTests
{
    [Fact]
    public void TargetIdentityNormalizesAbsolutePathAndWorksheetName()
    {
        var relative = Path.Combine(".", "identity-test.xlsx");
        var absolute = Path.GetFullPath(relative);
        Assert.Equal(
            ExcelTargetIdentity.Create(absolute.ToLowerInvariant(), "中外运"),
            ExcelTargetIdentity.Create(relative, " 中外运 "));
    }

    [Fact]
    public async Task ReservedRowIsReusedAfterProcessCrashBeforeExcelWrite()
    {
        var files = CreateFiles();
        try
        {
            CreateWorkbook(files.ExcelPath);
            var repository = new SqliteInvoiceRepository(files.DatabasePath);
            await repository.InitializeAsync();
            var application = CreateApplication(0);
            application.Id = (await repository.TryInsertAsync(application, Deduplication.CreateFallbackHash(application)))!.Value;
            var key = ExcelTargetIdentity.Create(files.ExcelPath, "中外运");
            Assert.True(await repository.TryReserveExcelRowAsync(application.Id, key, 2));

            var recovered = (await repository.GetPendingExcelAsync()).Single();
            var row = await new ExcelAssignmentCoordinator(repository, new ExcelWriter(), files.DataDirectory)
                .WriteApplicationAsync(recovered, files.ExcelPath, "中外运");

            Assert.Equal(2, row);
            using var workbook = new XLWorkbook(files.ExcelPath);
            Assert.Equal("测试天津企业", workbook.Worksheet("中外运").Cell(2, 2).GetString());
            Assert.Equal(ProcessingStatus.Completed, (await repository.GetRecentAsync(1)).Single().ProcessingStatus);
        }
        finally
        {
            DeleteFiles(files);
        }
    }

    [Fact]
    public async Task ReconcileMarksExcelWrittenBeforeDatabaseCompletionAsCompleted()
    {
        var files = CreateFiles();
        try
        {
            CreateWorkbook(files.ExcelPath);
            var repository = new SqliteInvoiceRepository(files.DatabasePath);
            await repository.InitializeAsync();
            var application = CreateApplication(0);
            application.Id = (await repository.TryInsertAsync(application, Deduplication.CreateFallbackHash(application)))!.Value;
            var key = ExcelTargetIdentity.Create(files.ExcelPath, "中外运");
            Assert.True(await repository.TryReserveExcelRowAsync(application.Id, key, 2));
            application.ExcelTargetKey = key;
            application.ExcelRow = 2;
            await new ExcelWriter().WriteAsync(application, files.ExcelPath, "中外运");

            var result = await new ExcelAssignmentCoordinator(repository, new ExcelWriter(), files.DataDirectory)
                .ReconcileAsync(files.ExcelPath, "中外运");

            Assert.Equal(1, result.MappedExistingRows);
            Assert.Equal(0, result.WrittenRows);
            Assert.Equal(ProcessingStatus.Completed, (await repository.GetRecentAsync(1)).Single().ProcessingStatus);
            using var workbook = new XLWorkbook(files.ExcelPath);
            Assert.Equal(2, workbook.Worksheet("中外运").LastRowUsed()!.RowNumber());
        }
        finally
        {
            DeleteFiles(files);
        }
    }

    [Fact]
    public async Task DistinctMailIdentitiesWithIdenticalBusinessFieldsUseDistinctRowsAndRetryIsIdempotent()
    {
        var files = CreateFiles();
        try
        {
            CreateWorkbook(files.ExcelPath);
            var repository = new SqliteInvoiceRepository(files.DatabasePath);
            await repository.InitializeAsync();
            var coordinator = new ExcelAssignmentCoordinator(repository, new ExcelWriter(), files.DataDirectory);
            var applications = Enumerable.Range(0, 3).Select(CreateApplication).ToArray();

            foreach (var application in applications)
            {
                application.Id = (await repository.TryInsertAsync(application, Deduplication.CreateFallbackHash(application)))!.Value;
                await repository.UpdateStatusAsync(application.Id, ProcessingStatus.PendingExcel);
                await coordinator.WriteApplicationAsync(application, files.ExcelPath, "中外运");
            }

            var rows = applications.Select(application => application.ExcelRow).ToArray();
            Assert.Equal(new int?[] { 2, 3, 4 }, rows);

            var retryRow = await coordinator.WriteApplicationAsync(applications[0], files.ExcelPath, "中外运");
            Assert.Equal(2, retryRow);

            using var workbook = new XLWorkbook(files.ExcelPath);
            var sheet = workbook.Worksheet("中外运");
            Assert.Equal("测试天津企业", sheet.Cell(2, 2).GetString());
            Assert.Equal("测试天津企业", sheet.Cell(3, 2).GetString());
            Assert.Equal("测试天津企业", sheet.Cell(4, 2).GetString());
            Assert.Equal(4, sheet.LastRowUsed()!.RowNumber());

            var stored = await repository.GetRecentAsync(10);
            Assert.Equal(new int?[] { 4, 3, 2 }, stored.Select(item => item.ExcelRow).ToArray());
            Assert.Single(stored.Select(item => item.ExcelTargetKey).Distinct(StringComparer.Ordinal));
            Assert.All(stored, item => Assert.Equal(ExcelTargetIdentity.Create(files.ExcelPath, "中外运"), item.ExcelTargetKey));
            Assert.All(stored, item => Assert.Equal(ProcessingStatus.Completed, item.ProcessingStatus));
        }
        finally
        {
            DeleteFiles(files);
        }
    }

    [Fact]
    public async Task ReconcileClaimsExistingDuplicateRowsOneToOneWithoutAppending()
    {
        var files = CreateFiles();
        try
        {
            CreateWorkbook(files.ExcelPath, rowCount: 3);
            var repository = new SqliteInvoiceRepository(files.DatabasePath);
            await repository.InitializeAsync();
            var applications = Enumerable.Range(0, 3).Select(CreateApplication).ToArray();
            foreach (var application in applications)
            {
                application.ProcessingStatus = ProcessingStatus.Completed;
                application.ExcelRow = 2;
                application.Id = (await repository.TryInsertAsync(application, Deduplication.CreateFallbackHash(application)))!.Value;
            }

            var coordinator = new ExcelAssignmentCoordinator(repository, new ExcelWriter(), files.DataDirectory);
            var result = await coordinator.ReconcileAsync(files.ExcelPath, "中外运");

            Assert.Equal(3, result.MappedExistingRows);
            Assert.Equal(0, result.WrittenRows);
            Assert.Equal(0, result.PendingRows);
            Assert.NotNull(result.BackupDirectory);
            Assert.True(File.Exists(Path.Combine(result.BackupDirectory!, "invoice-mail.db")));
            Assert.True(File.Exists(Path.Combine(result.BackupDirectory!, Path.GetFileName(files.ExcelPath))));

            var stored = (await repository.GetRecentAsync(10)).OrderBy(item => item.Id).ToArray();
            Assert.Equal(new int?[] { 2, 3, 4 }, stored.Select(item => item.ExcelRow).ToArray());
            Assert.All(stored, item => Assert.Equal(ExcelTargetIdentity.Create(files.ExcelPath, "中外运"), item.ExcelTargetKey));
            using var workbook = new XLWorkbook(files.ExcelPath);
            Assert.Equal(4, workbook.Worksheet("中外运").LastRowUsed()!.RowNumber());
        }
        finally
        {
            DeleteFiles(files);
        }
    }

    [Fact]
    public async Task ReconcileAppendsOnlyMissingRowsWhenWorkbookHasOneLegacyRow()
    {
        var files = CreateFiles();
        try
        {
            CreateWorkbook(files.ExcelPath, rowCount: 1);
            var repository = new SqliteInvoiceRepository(files.DatabasePath);
            await repository.InitializeAsync();
            foreach (var application in Enumerable.Range(0, 3).Select(CreateApplication))
            {
                application.ProcessingStatus = ProcessingStatus.Completed;
                application.ExcelRow = 2;
                application.Id = (await repository.TryInsertAsync(application, Deduplication.CreateFallbackHash(application)))!.Value;
            }

            var result = await new ExcelAssignmentCoordinator(repository, new ExcelWriter(), files.DataDirectory)
                .ReconcileAsync(files.ExcelPath, "中外运");

            Assert.Equal(1, result.MappedExistingRows);
            Assert.Equal(2, result.WrittenRows);
            Assert.Equal(0, result.PendingRows);
            var stored = (await repository.GetRecentAsync(10)).OrderBy(item => item.Id).ToArray();
            Assert.Equal(new int?[] { 2, 3, 4 }, stored.Select(item => item.ExcelRow).ToArray());
            using var workbook = new XLWorkbook(files.ExcelPath);
            Assert.Equal(4, workbook.Worksheet("中外运").LastRowUsed()!.RowNumber());
        }
        finally
        {
            DeleteFiles(files);
        }
    }

    [Fact]
    public async Task ExcelRowReservationIsScopedByTargetKeyAndUniqueWithinOneWorkbook()
    {
        var files = CreateFiles();
        var secondWorkbook = Path.Combine(files.DataDirectory, "other.xlsx");
        try
        {
            CreateWorkbook(files.ExcelPath);
            CreateWorkbook(secondWorkbook);
            var repository = new SqliteInvoiceRepository(files.DatabasePath);
            await repository.InitializeAsync();
            var first = CreateApplication(0);
            var second = CreateApplication(1);
            var third = CreateApplication(2);
            first.Id = (await repository.TryInsertAsync(first, "reservation-1"))!.Value;
            second.Id = (await repository.TryInsertAsync(second, "reservation-2"))!.Value;
            third.Id = (await repository.TryInsertAsync(third, "reservation-3"))!.Value;

            var firstKey = ExcelTargetIdentity.Create(files.ExcelPath, "中外运");
            var secondKey = ExcelTargetIdentity.Create(secondWorkbook, "中外运");
            Assert.True(await repository.TryReserveExcelRowAsync(first.Id, firstKey, 2));
            Assert.False(await repository.TryReserveExcelRowAsync(second.Id, firstKey, 2));
            Assert.True(await repository.TryReserveExcelRowAsync(second.Id, secondKey, 2));
            Assert.True(await repository.TryReserveExcelRowAsync(third.Id, firstKey, 3));
        }
        finally
        {
            DeleteFiles(files);
            if (File.Exists(secondWorkbook)) File.Delete(secondWorkbook);
        }
    }

    [Fact]
    public async Task ReadOnlyWorkbookLeavesReconciliationRecordPendingAndDoesNotOverwrite()
    {
        var files = CreateFiles();
        try
        {
            CreateWorkbook(files.ExcelPath);
            var before = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(await File.ReadAllBytesAsync(files.ExcelPath)));
            var repository = new SqliteInvoiceRepository(files.DatabasePath);
            await repository.InitializeAsync();
            var application = CreateApplication(0);
            application.ProcessingStatus = ProcessingStatus.Completed;
            application.ExcelRow = 2;
            application.Id = (await repository.TryInsertAsync(application, Deduplication.CreateFallbackHash(application)))!.Value;
            File.SetAttributes(files.ExcelPath, FileAttributes.ReadOnly);

            var result = await new ExcelAssignmentCoordinator(repository, new ExcelWriter(), files.DataDirectory)
                .ReconcileAsync(files.ExcelPath, "中外运");

            Assert.Equal(1, result.PendingRows);
            var stored = (await repository.GetRecentAsync(1)).Single();
            Assert.Equal(ProcessingStatus.PendingExcel, stored.ProcessingStatus);
            Assert.NotNull(stored.ExcelRow);
            Assert.Equal(ExcelTargetIdentity.Create(files.ExcelPath, "中外运"), stored.ExcelTargetKey);
            var after = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(await File.ReadAllBytesAsync(files.ExcelPath)));
            Assert.Equal(before, after);
        }
        finally
        {
            if (File.Exists(files.ExcelPath)) File.SetAttributes(files.ExcelPath, FileAttributes.Normal);
            DeleteFiles(files);
        }
    }

    [Fact]
    public async Task BackupFailureStopsBeforeChangingDatabaseOrWorkbook()
    {
        var files = CreateFiles();
        var invalidDataDirectory = Path.Combine(files.DataDirectory, "not-a-directory");
        try
        {
            CreateWorkbook(files.ExcelPath);
            var before = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(await File.ReadAllBytesAsync(files.ExcelPath)));
            var repository = new SqliteInvoiceRepository(files.DatabasePath);
            await repository.InitializeAsync();
            var application = CreateApplication(0);
            application.ProcessingStatus = ProcessingStatus.Completed;
            application.ExcelRow = 2;
            application.Id = (await repository.TryInsertAsync(application, Deduplication.CreateFallbackHash(application)))!.Value;
            await File.WriteAllTextAsync(invalidDataDirectory, "this is a file, not the backup directory");

            await Assert.ThrowsAnyAsync<IOException>(() => new ExcelAssignmentCoordinator(repository, new ExcelWriter(), invalidDataDirectory)
                .ReconcileAsync(files.ExcelPath, "中外运"));

            var stored = (await repository.GetRecentAsync(1)).Single();
            Assert.Equal(ProcessingStatus.Completed, stored.ProcessingStatus);
            Assert.Equal(2, stored.ExcelRow);
            Assert.Equal(string.Empty, stored.ExcelTargetKey);
            var after = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(await File.ReadAllBytesAsync(files.ExcelPath)));
            Assert.Equal(before, after);
        }
        finally
        {
            DeleteFiles(files);
        }
    }

    private static InvoiceApplication CreateApplication(int index)
        => new()
        {
            CompanyName = "测试天津企业",
            CreditCode = "91120111MA7LKBW877",
            Amount = 300m,
            ApplyTime = new DateTime(2026, 8, 27, 10, 15, 0),
            Email = "xiriteng@163.com",
            MessageId = $"<tianjin-{index}@example.test>",
            ImapUid = (uint)(100 + index),
            UidValidity = 9,
            MailboxName = "INBOX",
            MailboxIdentity = "account@example.test",
            MailReceivedAt = new DateTimeOffset(2026, 8, 27, 10, 15 + index, 0, TimeSpan.FromHours(8)),
            MailSubject = $"开票申请 {index}",
            MailFrom = "sender@example.test",
            NormalizedBody = $"公司名称：测试天津企业\n信用代码：91120111MA7LKBW877\n申请金额：300\n申请时间：2026-08-27 10:15\n邮箱：xiriteng@163.com\n正文序号：{index}",
            ProcessingStatus = ProcessingStatus.Parsed
        };

    private static void CreateWorkbook(string path, int rowCount = 0)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("中外运");
        sheet.Cell(1, 1).Value = "日期";
        sheet.Cell(1, 2).Value = "企业名称";
        sheet.Cell(1, 3).Value = "信用代码";
        sheet.Cell(1, 4).Value = "金额";
        sheet.Cell(1, 5).Value = "人工列";
        sheet.Cell(1, 6).Value = "邮箱";
        for (var row = 2; row <= rowCount + 1; row++)
        {
            sheet.Cell(row, 1).Value = row == 2 ? "8.27" : string.Empty;
            sheet.Cell(row, 2).Value = "测试天津企业";
            sheet.Cell(row, 3).Value = "91120111MA7LKBW877";
            sheet.Cell(row, 4).Value = 300m;
            sheet.Cell(row, 5).Value = $"人工-{row}";
            sheet.Cell(row, 6).Value = "xiriteng@163.com";
        }
        workbook.SaveAs(path);
    }

    private static (string DataDirectory, string DatabasePath, string ExcelPath) CreateFiles()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"invoice-assignment-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        return (directory, Path.Combine(directory, "invoice-mail.db"), Path.Combine(directory, "登记表.xlsx"));
    }

    private static void DeleteFiles((string DataDirectory, string DatabasePath, string ExcelPath) files)
    {
        if (!Directory.Exists(files.DataDirectory)) return;
        foreach (var file in Directory.EnumerateFiles(files.DataDirectory, "*", SearchOption.AllDirectories))
        {
            try { File.SetAttributes(file, FileAttributes.Normal); }
            catch (IOException) { }
        }
        Directory.Delete(files.DataDirectory, recursive: true);
    }
}
