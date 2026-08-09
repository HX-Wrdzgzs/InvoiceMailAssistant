using ClosedXML.Excel;
using InvoiceMailAssistant.App;
using Xunit;

namespace InvoiceMailAssistant.Tests;

public sealed class ExcelWriterTests
{
    [Fact]
    public async Task WritesOnlyBusinessOutputColumnsAndReusesPlannedRow()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var createdSheet = workbook.AddWorksheet("中外运");
                createdSheet.Cell(1, 1).Value = "日期";
                createdSheet.Cell(1, 2).Value = "企业名称";
                createdSheet.Cell(1, 3).Value = "信用代码";
                createdSheet.Cell(1, 4).Value = "开票金额";
                createdSheet.Cell(1, 5).Value = "到帐时间";
                createdSheet.Cell(1, 6).Value = "邮箱地址";
                createdSheet.Cell(1, 7).Value = "是否开票";
                createdSheet.Cell(1, 8).Value = "发票是否已发";
                workbook.SaveAs(path);
            }

            var app = new InvoiceApplication
            {
                CompanyName = "测试物流有限公司",
                CreditCode = "TEST-CREDIT-20260808",
                Amount = 300m,
                ApplyTime = new DateTime(2026, 8, 7, 10, 28, 0),
                Email = "invoice@example.com"
            };
            var writer = new ExcelWriter();
            app.ExcelRow = writer.ResolveTargetRow(app, path, "中外运");

            var firstRow = await writer.WriteAsync(app, path, "中外运");
            var recoveredRow = writer.ResolveTargetRow(app, path, "中外运");
            var secondRow = await writer.WriteAsync(app, path, "中外运");

            Assert.Equal(2, firstRow);
            Assert.Equal(firstRow, recoveredRow);
            Assert.Equal(firstRow, secondRow);

            using var verify = new XLWorkbook(path);
            var sheet = verify.Worksheet("中外运");
            Assert.Equal("8.7", sheet.Cell(2, 1).GetString());
            Assert.Equal("测试物流有限公司", sheet.Cell(2, 2).GetString());
            Assert.Equal("TEST-CREDIT-20260808", sheet.Cell(2, 3).GetString());
            Assert.Equal(300m, sheet.Cell(2, 4).GetValue<decimal>());
            Assert.True(sheet.Cell(2, 5).IsEmpty());
            Assert.Equal("invoice@example.com", sheet.Cell(2, 6).GetString());
            Assert.True(sheet.Cell(2, 7).IsEmpty());
            Assert.True(sheet.Cell(2, 8).IsEmpty());
            Assert.True(sheet.Cell(3, 2).IsEmpty());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task LeavesDateBlankForSameApplicationDay()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("中外运");
                sheet.Cell(1, 1).Value = "日期";
                sheet.Cell(2, 1).Value = "8.7";
                sheet.Cell(2, 2).Value = "前一企业";
                workbook.SaveAs(path);
            }

            var app = new InvoiceApplication
            {
                CompanyName = "第二企业",
                CreditCode = "CODE-2",
                Amount = 10m,
                ApplyTime = new DateTime(2026, 8, 7, 11, 0, 0),
                Email = "second@example.com"
            };
            var writer = new ExcelWriter();
            app.ExcelRow = writer.ResolveTargetRow(app, path, "中外运");
            var row = await writer.WriteAsync(app, path, "中外运");

            using var verify = new XLWorkbook(path);
            Assert.Equal(3, row);
            Assert.True(verify.Worksheet("中外运").Cell(3, 1).IsEmpty());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task ResumesAnEmptyPersistedRowEvenWhenSheetHasLaterRows()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("中外运");
                sheet.Cell(1, 1).Value = "日期";
                sheet.Cell(10, 2).Value = "后续业务数据";
                workbook.SaveAs(path);
            }

            var app = CreateApplication(2);
            var writer = new ExcelWriter();
            Assert.Equal(2, writer.ResolveTargetRow(app, path, "中外运"));
            Assert.Equal(2, await writer.WriteAsync(app, path, "中外运"));

            using var verify = new XLWorkbook(path);
            Assert.Equal("测试企业", verify.Worksheet("中外运").Cell(2, 2).GetString());
            Assert.Equal("后续业务数据", verify.Worksheet("中外运").Cell(10, 2).GetString());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task ReplansWhenAUserOccupiesThePersistedRow()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("中外运");
                sheet.Cell(1, 1).Value = "日期";
                sheet.Cell(2, 2).Value = "用户手工数据";
                workbook.SaveAs(path);
            }

            var app = CreateApplication(2);
            var writer = new ExcelWriter();
            app.ExcelRow = writer.ResolveTargetRow(app, path, "中外运");
            Assert.Equal(3, app.ExcelRow);
            Assert.Equal(3, await writer.WriteAsync(app, path, "中外运"));

            using var verify = new XLWorkbook(path);
            var verifySheet = verify.Worksheet("中外运");
            Assert.Equal("用户手工数据", verifySheet.Cell(2, 2).GetString());
            Assert.Equal("测试企业", verifySheet.Cell(3, 2).GetString());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task FindsExistingApplicationAfterManualRowInsertion()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("中外运");
                sheet.Cell(1, 1).Value = "日期";
                sheet.Cell(2, 2).Value = "前置人工行";
                sheet.Cell(3, 2).Value = "测试企业";
                sheet.Cell(3, 3).Value = "CODE-1";
                sheet.Cell(3, 4).Value = 100m;
                sheet.Cell(3, 6).Value = "finance@example.com";
                workbook.SaveAs(path);
            }

            var app = CreateApplication(2);
            var writer = new ExcelWriter();
            Assert.Equal(3, writer.ResolveTargetRow(app, path, "中外运"));
            Assert.Equal(3, await writer.WriteAsync(app, path, "中外运"));

            using var verify = new XLWorkbook(path);
            Assert.Equal("前置人工行", verify.Worksheet("中外运").Cell(2, 2).GetString());
            Assert.Equal(3, writer.ResolveTargetRow(app, path, "中外运"));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task KeepsManualColumnsUntouched()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("中外运");
                sheet.Cell(1, 1).Value = "日期";
                sheet.Cell(2, 2).Value = "测试企业";
                sheet.Cell(2, 3).Value = "CODE-1";
                sheet.Cell(2, 4).Value = 100m;
                sheet.Cell(2, 5).Value = "人工到账时间";
                sheet.Cell(2, 6).Value = "finance@example.com";
                sheet.Cell(2, 7).Value = "是";
                sheet.Cell(2, 8).Value = "否";
                workbook.SaveAs(path);
            }

            var app = CreateApplication(2);
            var writer = new ExcelWriter();
            Assert.Equal(2, await writer.WriteAsync(app, path, "中外运"));

            using var verify = new XLWorkbook(path);
            var verifySheet = verify.Worksheet("中外运");
            Assert.Equal("人工到账时间", verifySheet.Cell(2, 5).GetString());
            Assert.Equal("是", verifySheet.Cell(2, 7).GetString());
            Assert.Equal("否", verifySheet.Cell(2, 8).GetString());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task ConcurrentDifferentApplicationsCannotOverwriteTheSameRow()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("中外运");
                sheet.Cell(1, 1).Value = "日期";
                sheet.Cell(1, 2).Value = "企业名称";
                workbook.SaveAs(path);
            }

            var first = CreateApplication(2);
            first.CompanyName = "并发企业一";
            var second = CreateApplication(2);
            second.CompanyName = "并发企业二";
            var writer = new ExcelWriter();

            static async Task<Exception?> Capture(Func<Task<int>> write)
            {
                try { await write(); return null; }
                catch (Exception ex) { return ex; }
            }

            var outcomes = await Task.WhenAll(
                Capture(() => writer.WriteAsync(first, path, "中外运")),
                Capture(() => writer.WriteAsync(second, path, "中外运")));

            Assert.Single(outcomes, x => x is null);
            Assert.Single(outcomes, x => x is ExcelRowOccupiedException);
            using var verify = new XLWorkbook(path);
            Assert.Contains(verify.Worksheet("中外运").Cell(2, 2).GetString(), new[] { "并发企业一", "并发企业二" });
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task StartsNewYearWithNewDisplayedDate()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("中外运");
                sheet.Cell(1, 1).Value = "日期";
                sheet.Cell(2, 1).Value = "12.31";
                sheet.Cell(2, 2).Value = "上一年度企业";
                workbook.SaveAs(path);
            }

            var app = CreateApplication(3);
            app.ApplyTime = new DateTime(2027, 1, 1, 9, 0, 0);
            var writer = new ExcelWriter();
            app.ExcelRow = writer.ResolveTargetRow(app, path, "中外运");
            await writer.WriteAsync(app, path, "中外运");

            using var verify = new XLWorkbook(path);
            Assert.Equal("1.1", verify.Worksheet("中外运").Cell(3, 1).GetString());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static InvoiceApplication CreateApplication(int row)
        => new()
        {
            CompanyName = "测试企业",
            CreditCode = "CODE-1",
            Amount = 100m,
            ApplyTime = new DateTime(2026, 8, 8, 9, 0, 0),
            Email = "finance@example.com",
            ExcelRow = row
        };
}
