using System.IO.Compression;
using System.Text;
using ClosedXML.Excel;
using ExcelTools.Core.Workspace;

namespace ExcelTools.Core.Tests;

public sealed class WorkbookPackageTests : IDisposable
{
    private readonly TempDirectory _tmp = new();

    public void Dispose() => _tmp.Dispose();

    private byte[] ValidBook()
    {
        var path = _tmp.Combine("book.xlsx");
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Data").Cell("A1").Value = 1;
            wb.AddWorksheet("Other").Cell("A1").Value = 2;
            wb.SaveAs(path);
        }

        return File.ReadAllBytes(path);
    }

    private static byte[] ReplaceInZip(byte[] zipBytes, string entryName, Func<string, string> edit)
    {
        using var ms = new MemoryStream();
        ms.Write(zipBytes);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry(entryName)!;
            string xml;
            using (var reader = new StreamReader(entry.Open()))
            {
                xml = reader.ReadToEnd();
            }

            entry.Delete();
            using var writer = new StreamWriter(zip.CreateEntry(entryName).Open());
            writer.Write(edit(xml));
        }

        return ms.ToArray();
    }

    [Fact]
    public void A_valid_workbook_has_no_problems() =>
        Assert.Empty(WorkbookPackage.ValidateSheet(ValidBook(), "Data"));

    [Fact]
    public void Sheet_name_lookup_is_case_insensitive_and_unknown_names_still_validate_the_workbook_part()
    {
        Assert.Empty(WorkbookPackage.ValidateSheet(ValidBook(), "DATA"));
        Assert.Empty(WorkbookPackage.ValidateSheet(ValidBook(), "NoSuchSheet"));
    }

    [Fact]
    public void A_corrupt_target_sheet_is_reported()
    {
        var broken = ReplaceInZip(ValidBook(), "xl/worksheets/sheet1.xml", xml => xml.Replace("<x:sheetData>", "<x:bogus /><x:sheetData>", StringComparison.Ordinal));

        Assert.Contains("<x:sheetData>", Encoding.UTF8.GetString(ReadEntry(ValidBook(), "xl/worksheets/sheet1.xml")), StringComparison.Ordinal); // 對照：替換的目標字串確實存在
        Assert.NotEmpty(WorkbookPackage.ValidateSheet(broken, "Data"));
    }

    [Fact]
    public void Only_the_target_sheet_is_validated_not_every_sheet()
    {
        var broken = ReplaceInZip(ValidBook(), "xl/worksheets/sheet2.xml", xml => xml.Replace("<x:sheetData>", "<x:bogus /><x:sheetData>", StringComparison.Ordinal));

        Assert.Empty(WorkbookPackage.ValidateSheet(broken, "Data"));
        Assert.NotEmpty(WorkbookPackage.ValidateSheet(broken, "Other"));
    }

    [Fact]
    public void A_corrupt_workbook_part_is_reported_whatever_the_sheet()
    {
        var broken = ReplaceInZip(ValidBook(), "xl/workbook.xml", xml => xml.Replace("<x:sheets>", "<x:bogus /><x:sheets>", StringComparison.Ordinal));
        Assert.NotEmpty(WorkbookPackage.ValidateSheet(broken, "Data"));
    }

    [Fact]
    public void Serialize_produces_a_loadable_file_with_the_same_content()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("A1").Value = "hello";
        ws.Cell("A2").FormulaA1 = "1+1";
        ws.Cell("A3").Value = new DateTime(2026, 10, 3);

        var bytes = WorkbookPackage.Serialize(wb);

        using var reloaded = WorkbookPackage.Load(bytes);
        var loaded = reloaded.Worksheet("S");
        Assert.Equal("hello", loaded.Cell("A1").Value.GetText());
        Assert.Equal("1+1", loaded.Cell("A2").FormulaA1);
        Assert.Equal(2.0, loaded.Cell("A2").Value.GetNumber());
        Assert.Equal(new DateTime(2026, 10, 3), loaded.Cell("A3").Value.GetDateTime());
        Assert.Empty(WorkbookPackage.ValidateSheet(bytes, "S"));
    }

    [Fact]
    public void Serialize_runs_in_parallel_without_interference()
    {
        var results = Enumerable.Range(0, 16).AsParallel().Select(i =>
        {
            using var wb = new XLWorkbook();
            wb.AddWorksheet("S").Cell("A1").Value = i;
            using var reloaded = WorkbookPackage.Load(WorkbookPackage.Serialize(wb));
            return reloaded.Worksheet("S").Cell("A1").Value.GetNumber();
        }).OrderBy(x => x).ToArray();

        Assert.Equal(Enumerable.Range(0, 16).Select(i => (double)i), results);
    }

    private static byte[] ReadEntry(byte[] zipBytes, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(zipBytes), ZipArchiveMode.Read);
        using var s = zip.GetEntry(name)!.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }
}
