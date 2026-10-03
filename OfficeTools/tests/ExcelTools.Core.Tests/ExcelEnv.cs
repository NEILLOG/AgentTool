using System.IO.Compression;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using ExcelTools.Core.Operations;
using ExcelTools.Core.Workspace;
using OfficeTools.Common;
using OfficeTools.Common.Security;

namespace ExcelTools.Core.Tests;

internal sealed class FakeTime(DateTimeOffset start) : TimeProvider
{
    private DateTimeOffset _now = start;

    public override DateTimeOffset GetUtcNow() => _now;

    public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;

    public void Advance(TimeSpan by) => _now += by;
}

/// <summary>每個測試一組獨立的暫存資料夾、PathGuard、session 管理器與檔案操作。</summary>
internal sealed class ExcelEnv : IDisposable
{
    private readonly TempDirectory _tmp = new();

    public ExcelEnv(bool allowOverwrite = false, int maxUncompressedMb = 500, TimeSpan? idleTimeout = null)
    {
        Root = _tmp.Combine("root");
        Directory.CreateDirectory(Root);
        Time = new FakeTime(new DateTimeOffset(2026, 10, 3, 9, 0, 0, TimeSpan.Zero));
        Options = new OfficeToolsOptions
        {
            AllowedRoots = [Root],
            AllowOverwrite = allowOverwrite,
            MaxUncompressedMb = maxUncompressedMb,
        };
        Guard = new PathGuard(Options, Time);
        Sessions = new WorkbookSessionManager(Guard, new ExcelToolsOptions { IdleTimeout = idleTimeout ?? TimeSpan.FromMinutes(30) }, Time);
        Files = new FileOperations(Guard, Options, Sessions);
        Sheets = new SheetOperations(Sessions);
    }

    public string Root { get; }

    public FakeTime Time { get; }

    public OfficeToolsOptions Options { get; }

    public PathGuard Guard { get; }

    public WorkbookSessionManager Sessions { get; }

    public FileOperations Files { get; }

    public SheetOperations Sheets { get; }

    public string Path(string relative) => System.IO.Path.Combine(Root, relative);

    /// <summary>API 回傳的是解析 symlink 後的實際路徑（macOS 的 /var 實際是 /private/var）。</summary>
    public string Resolved(string relative) => Guard.ResolveAllowed(Path(relative));

    public string Outside(string relative)
    {
        var full = _tmp.Combine("outside", relative);
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(full)!);
        return full;
    }

    /// <summary>在 root 內建立一個 xlsx，內容由 <paramref name="fill"/> 填入。</summary>
    public string MakeWorkbook(string relative, Action<IXLWorksheet>? fill = null)
    {
        var path = Path(relative);
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path)!);
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        if (fill is null)
        {
            ws.Cell("A1").Value = "name";
            ws.Cell("B1").Value = "qty";
            ws.Cell("A2").Value = "apple";
            ws.Cell("B2").Value = 3;
        }
        else
        {
            fill(ws);
        }

        wb.SaveAs(path);
        return path;
    }

    /// <summary>在 root 內建立多工作表的 xlsx，內容完全由 <paramref name="build"/> 決定。</summary>
    public string MakeBook(string relative, Action<XLWorkbook> build)
    {
        var path = Path(relative);
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path)!);
        using var wb = new XLWorkbook();
        build(wb);
        wb.SaveAs(path);
        return path;
    }

    /// <summary>在既有的 xlsx（zip）裡加入額外部件，用來模擬含圖表、樞紐、巨集等內容的檔案。</summary>
    public static void AddEntries(string path, params (string Name, string Content)[] entries)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        foreach (var (name, content) in entries)
        {
            using var writer = new StreamWriter(zip.CreateEntry(name).Open());
            writer.Write(content);
        }
    }

    public object? Cell(string workbookId, string address, int sheet = 1) =>
        Sessions.Use(workbookId, false, s => Internal.CellValueConverter.ToSerializable(s.Workbook.Worksheet(sheet).Cell(address).Value));

    public void SetCell(string workbookId, string address, object? value, int sheet = 1) =>
        Sessions.Use(workbookId, true, s => Internal.CellValueConverter.Apply(
            s.Workbook.Worksheet(sheet).Cell(address),
            Internal.CellValueConverter.ToCellInput(value, new Internal.CellWriteOptions())));

    /// <summary>用 Open XML SDK 驗證檔案符合 Office 2019 的 schema；有錯誤時回傳描述（空 = 合法）。</summary>
    public static IReadOnlyList<string> ValidateXlsx(string path)
    {
        using var doc = SpreadsheetDocument.Open(path, isEditable: false);
        return new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019)
            .Validate(doc)
            .Select(e => $"{e.Path?.XPath}: {e.Description}")
            .ToList();
    }

    public static object? ReadFromDisk(string path, string address, int sheet = 1)
    {
        using var wb = new XLWorkbook(path);
        return Internal.CellValueConverter.ToSerializable(wb.Worksheet(sheet).Cell(address).Value);
    }

    public void Dispose()
    {
        Sessions.Dispose();
        _tmp.Dispose();
    }
}
