using ClosedXML.Excel;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
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
    public async Task RepairsExistingParsedRowWithoutOverwritingManualColumns()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var sourceSheet = workbook.AddWorksheet("中外运");
                sourceSheet.Cell(1, 1).Value = "日期";
                sourceSheet.Cell(2, 1).Value = "1.1";
                sourceSheet.Cell(2, 2).Value = "江苏慧世联网络科技有限公司";
                sourceSheet.Cell(2, 3).Value = "XXXXXXXXXX";
                sourceSheet.Cell(2, 4).Value = 300m;
                sourceSheet.Cell(2, 5).Value = "人工到账时间";
                sourceSheet.Cell(2, 6).Value = "old@example.com";
                sourceSheet.Cell(2, 7).Value = "是";
                sourceSheet.Cell(2, 8).Value = "否";
                workbook.SaveAs(path);
            }

            var original = new InvoiceApplication
            {
                CompanyName = "江苏慧世联网络科技有限公司",
                CreditCode = "XXXXXXXXXX",
                Amount = 300m,
                ApplyTime = new DateTime(2023, 1, 1, 15, 0, 0),
                Email = "old@example.com",
                ExcelRow = 2
            };
            var corrected = new InvoiceApplication
            {
                CompanyName = "艾洛（天津）国际物流有限公司",
                CreditCode = "91120102MA7GXAJX2F",
                Amount = 300m,
                ApplyTime = new DateTime(2026, 8, 10, 15, 11, 0),
                Email = "yoyo.guo@auroragroup-cn.com",
                ExcelRow = 2
            };

            var writer = new ExcelWriter();
            await writer.RepairExistingRowAsync(original, corrected, path, "中外运");

            using var verify = new XLWorkbook(path);
            var sheet = verify.Worksheet("中外运");
            Assert.Equal("8.10", sheet.Cell(2, 1).GetString());
            Assert.Equal(corrected.CompanyName, sheet.Cell(2, 2).GetString());
            Assert.Equal(corrected.CreditCode, sheet.Cell(2, 3).GetString());
            Assert.Equal(corrected.Amount, sheet.Cell(2, 4).GetValue<decimal>());
            Assert.Equal("人工到账时间", sheet.Cell(2, 5).GetString());
            Assert.Equal(corrected.Email, sheet.Cell(2, 6).GetString());
            Assert.Equal("是", sheet.Cell(2, 7).GetString());
            Assert.Equal("否", sheet.Cell(2, 8).GetString());

            // A database update can fail after the Excel replacement. The
            // repair operation must therefore be idempotent on retry.
            await writer.RepairExistingRowAsync(original, corrected, path, "中外运");
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

    [Fact]
    public void ReportsReadOnlyWorkbookClearly()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("中外运");
                workbook.SaveAs(path);
            }

            File.SetAttributes(path, FileAttributes.ReadOnly);
            var error = Assert.Throws<IOException>(() => new ExcelWriter().ResolveTargetRow(CreateApplication(2), path, "中外运"));

            Assert.Contains("只读", error.Message);
        }
        finally
        {
            if (File.Exists(path))
            {
                File.SetAttributes(path, FileAttributes.Normal);
                File.Delete(path);
            }
        }
    }

    [Fact]
    public async Task WritesBoldCompanyCreditAndMailtoHyperlink()
    {
        var path = Path.Combine(Path.GetTempPath(), $"invoice-mail-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("中外运");
                sheet.Cell(1, 1).Value = "日期";
                sheet.Cell(1, 2).Value = "企业名称";
                sheet.Cell(2, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
                sheet.Cell(2, 2).Style.Font.Bold = true;
                workbook.SaveAs(path);
            }

            var stylesBefore = ReadZipEntry(path, "xl/styles.xml");
            var application = CreateApplication(2);
            await new ExcelWriter().WriteAsync(application, path, "中外运");
            var stylesAfter = ReadZipEntry(path, "xl/styles.xml");

            using var verify = new XLWorkbook(path);
            var verifySheet = verify.Worksheet("中外运");
            Assert.Equal(application.CompanyName, verifySheet.Cell(2, 2).GetString());
            Assert.True(verifySheet.Cell(2, 2).Style.Font.Bold);
            Assert.True(verifySheet.Cell(2, 3).Style.Font.Bold);

            var worksheet = XDocument.Parse(Encoding.UTF8.GetString(ReadZipEntry(path, "xl/worksheets/sheet1.xml")).TrimStart('\uFEFF'));
            XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace officeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var hyperlink = worksheet.Descendants(spreadsheet + "hyperlink")
                .Single(x => x.Attribute("ref")?.Value == "F2");
            var relationshipId = hyperlink.Attribute(officeRelationships + "id")?.Value;
            Assert.False(string.IsNullOrWhiteSpace(relationshipId));

            var relationships = XDocument.Parse(Encoding.UTF8.GetString(ReadZipEntry(path, "xl/worksheets/_rels/sheet1.xml.rels")).TrimStart('\uFEFF'));
            XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
            Assert.Equal(
                "mailto:finance@example.com",
                relationships.Descendants(packageRelationships + "Relationship")
                    .Single(x => x.Attribute("Id")?.Value == relationshipId)
                    .Attribute("Target")?.Value);

            Assert.NotEqual(stylesBefore, stylesAfter);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static byte[] ReadZipEntry(string path, string entryName)
    {
        using var archive = ZipFile.OpenRead(path);
        using var stream = archive.GetEntry(entryName)!.Open();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
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
