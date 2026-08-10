using System.IO.Compression;
using System.Globalization;
using System.Xml.Linq;
using ClosedXML.Excel;
using InvoiceMailAssistant.App;
using Xunit;

namespace InvoiceMailAssistant.Tests;

public sealed class RealWorkbookCompatibilityTests
{
    [RealWorkbookFact]
    public async Task RealWorkbookCopyPreservesStructureAndManualColumns()
    {
        var path = Environment.GetEnvironmentVariable("INVOICE_REAL_XLSX_COPY");
        Assert.False(string.IsNullOrWhiteSpace(path));
        Assert.True(File.Exists(path), $"Real workbook copy not found: {path}");

        var before = WorkbookSnapshot.Read(path!);
        var application = new InvoiceApplication
        {
            CompanyName = "合成工作簿回归测试",
            CreditCode = "TEST-CODE-20260808",
            Amount = 123.45m,
            ApplyTime = new DateTime(2026, 8, 8, 10, 30, 0),
            Email = "codex-regression@example.com"
        };
        var writer = new ExcelWriter();
        application.ExcelRow = writer.ResolveTargetRow(application, path!, "中外运");
        var writtenRow = await writer.WriteAsync(application, path!, "中外运");
        var after = WorkbookSnapshot.Read(path!);

        Assert.Contains("中外运", before.SheetNames);
        Assert.True(before.EntryNames.All(after.EntryNames.Contains), "The saved workbook lost ZIP parts.");
        Assert.Equal(before.SheetNames, after.SheetNames);
        Assert.Equal(before.FormulaTexts, after.FormulaTexts);
        Assert.True(before.StylesXml.SequenceEqual(after.StylesXml), "The saved workbook rewrote xl/styles.xml.");
        Assert.Equal(before.TargetSheet.StructureXml, after.TargetSheet.StructureXml);
        Assert.True(before.TargetSheet.ManualCells.All(pair =>
            after.TargetSheet.ManualCells.TryGetValue(pair.Key, out var value) && value == pair.Value),
            "Existing E/G/H cells changed during the write.");

        using var verify = new XLWorkbook(path!);
        var sheet = verify.Worksheet("中外运");
        Assert.Equal(application.CompanyName, sheet.Cell(writtenRow, 2).GetString());
        Assert.Equal(application.CreditCode, sheet.Cell(writtenRow, 3).GetString());
        Assert.Equal(application.Amount, sheet.Cell(writtenRow, 4).GetValue<decimal>());
        Assert.Equal(application.Email, sheet.Cell(writtenRow, 6).GetString());
    }

    private sealed record WorkbookSnapshot(
        IReadOnlyList<string> EntryNames,
        IReadOnlyList<string> SheetNames,
        IReadOnlyList<string> FormulaTexts,
        byte[] StylesXml,
        SheetSnapshot TargetSheet)
    {
        public static WorkbookSnapshot Read(string path)
        {
            using var archive = ZipFile.OpenRead(path);
            var workbook = LoadXml(archive, "xl/workbook.xml");
            var relationships = LoadXml(archive, "xl/_rels/workbook.xml.rels");
            XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace officeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

            var relationshipTargets = relationships.Root!.Elements(packageRelationships + "Relationship")
                .ToDictionary(x => x.Attribute("Id")!.Value, x => ResolveWorksheetPath(x.Attribute("Target")!.Value));
            var sheets = workbook.Root!.Element(spreadsheet + "sheets")!.Elements(spreadsheet + "sheet").ToArray();
            var sheetNames = sheets.Select(x => (string)x.Attribute("name")!).ToArray();
            var targetSheet = sheets.Single(x => x.Attribute("name")?.Value == "中外运");
            var targetPath = relationshipTargets[targetSheet.Attribute(officeRelationships + "id")!.Value];
            var targetXml = LoadXml(archive, targetPath);
            var stylesXml = LoadBytes(archive, "xl/styles.xml");
            var formulaTexts = archive.Entries
                .Where(x => x.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) && x.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .OrderBy(x => x.FullName, StringComparer.Ordinal)
                .SelectMany(x => LoadXml(archive, x.FullName).Descendants(spreadsheet + "f").Select(f => $"{x.FullName}:{f.Value}"))
                .ToArray();

            return new WorkbookSnapshot(
                archive.Entries.Select(x => x.FullName).OrderBy(x => x, StringComparer.Ordinal).ToArray(),
                sheetNames,
                formulaTexts,
                stylesXml,
                SheetSnapshot.Read(targetXml, spreadsheet));
        }

        private static XDocument LoadXml(ZipArchive archive, string path)
        {
            using var stream = archive.GetEntry(path)!.Open();
            return XDocument.Load(stream, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        }

        private static byte[] LoadBytes(ZipArchive archive, string path)
        {
            using var stream = archive.GetEntry(path)!.Open();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }

        private static string ResolveWorksheetPath(string target)
        {
            target = target.Replace('\\', '/');
            if (target.StartsWith("/", StringComparison.Ordinal)) return target.TrimStart('/');
            if (target.StartsWith("../", StringComparison.Ordinal)) return "xl/" + target[3..];
            return "xl/" + target.TrimStart('/');
        }
    }

    private sealed record SheetSnapshot(string StructureXml, IReadOnlyDictionary<string, string> ManualCells)
    {
        public static SheetSnapshot Read(XDocument document, XNamespace spreadsheet)
        {
            var root = document.Root!;
            var structure = string.Join("\n", new[]
            {
                ElementXml(root.Element(spreadsheet + "sheetPr")),
                ElementXml(root.Element(spreadsheet + "sheetFormatPr")),
                ElementXml(root.Element(spreadsheet + "cols")),
                ElementXml(root.Element(spreadsheet + "mergeCells")),
                ElementXml(root.Element(spreadsheet + "sheetViews")?.Element(spreadsheet + "sheetView")?.Element(spreadsheet + "pane")),
                string.Join("", root.Descendants(spreadsheet + "row").Where(x => string.Equals((string?)x.Attribute("hidden"), "1", StringComparison.OrdinalIgnoreCase) || string.Equals((string?)x.Attribute("hidden"), "true", StringComparison.OrdinalIgnoreCase)).Select(ElementXml)),
                string.Join("", root.Elements(spreadsheet + "conditionalFormatting").Select(ElementXml)),
                ElementXml(root.Element(spreadsheet + "dataValidations")),
                ElementXml(root.Element(spreadsheet + "pageMargins")),
                ElementXml(root.Element(spreadsheet + "pageSetup")),
                ElementXml(root.Element(spreadsheet + "printOptions")),
                ElementXml(root.Element(spreadsheet + "headerFooter")),
                ElementXml(root.Element(spreadsheet + "rowBreaks")),
                ElementXml(root.Element(spreadsheet + "colBreaks"))
            });

            var cells = root.Descendants(spreadsheet + "c")
                .Where(x => IsManualColumn((string?)x.Attribute("r")))
                .ToDictionary(x => (string)x.Attribute("r")!, ElementXml, StringComparer.Ordinal);
            return new SheetSnapshot(structure, cells);
        }

        private static bool IsManualColumn(string? reference)
            => !string.IsNullOrWhiteSpace(reference) && reference!.Length > 1 && reference[0] is 'E' or 'G' or 'H';

        private static string ElementXml(XElement? element)
            => element is null ? string.Empty : CanonicalXml(element);

        private static string CanonicalXml(XElement element)
        {
            var attributes = string.Join(";", element.Attributes()
                .Where(x => !x.IsNamespaceDeclaration
                    && !string.Equals(x.Name.LocalName, "customWidth", StringComparison.Ordinal)
                    && !(string.Equals(element.Name.LocalName, "col", StringComparison.Ordinal)
                        && string.Equals(x.Name.LocalName, "style", StringComparison.Ordinal))
                    && !(string.Equals(element.Name.LocalName, "c", StringComparison.Ordinal)
                        && string.Equals(x.Name.LocalName, "s", StringComparison.Ordinal)))
                .OrderBy(x => x.Name.NamespaceName, StringComparer.Ordinal)
                .ThenBy(x => x.Name.LocalName, StringComparer.Ordinal)
                .Select(x => $"{x.Name.NamespaceName}:{x.Name.LocalName}={CanonicalAttributeValue(element, x)}"));
            var children = string.Concat(element.Nodes().Select(node => node switch
            {
                XElement child => CanonicalXml(child),
                XText text => $"T:{text.Value}",
                _ => string.Empty
            }));
            return $"E:{element.Name.NamespaceName}:{element.Name.LocalName}[{attributes}]{{{children}}}";
        }

        private static string CanonicalAttributeValue(XElement element, XAttribute attribute)
        {
            if (string.Equals(element.Name.LocalName, "col", StringComparison.Ordinal)
                && string.Equals(attribute.Name.LocalName, "width", StringComparison.Ordinal)
                && double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var width))
                return width.ToString("0.######", CultureInfo.InvariantCulture);
            return attribute.Value;
        }
    }
}

public sealed class RealWorkbookFactAttribute : FactAttribute
{
    public RealWorkbookFactAttribute()
    {
        if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("INVOICE_REAL_XLSX_COPY")))
            Skip = "Set INVOICE_REAL_XLSX_COPY to run the isolated real-workbook regression.";
    }
}
