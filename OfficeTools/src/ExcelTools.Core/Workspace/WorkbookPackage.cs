using System.IO.Compression;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeTools.Common;
using OfficeTools.Common.Errors;
using OfficeTools.Common.Security;

namespace ExcelTools.Core.Workspace;

/// <summary>xlsx 檔案層級的工作：共享讀取、結構檢查（含 ClosedXML 無法保留的內容）、載入、原子寫入。</summary>
internal static partial class WorkbookPackage
{
    // 檔名前綴 → 存檔後會遺失的內容
    private static readonly (string Prefix, string Label)[] LossyParts =
    [
        ("xl/charts/", "圖表"),
        ("xl/chartsheets/", "圖表工作表"),
        ("xl/pivotTables/", "樞紐分析表"),
        ("xl/pivotCache/", "樞紐分析表"),
        ("xl/vbaProject.bin", "VBA 巨集"),
        ("xl/slicers/", "交叉分析篩選器"),
        ("xl/slicerCaches/", "交叉分析篩選器"),
        ("xl/timelines/", "時間表篩選器"),
        ("xl/ctrlProps/", "表單控制項"),
        ("xl/activeX/", "ActiveX 控制項"),
        ("xl/embeddings/", "內嵌物件"),
        ("xl/threadedComments/", "討論串註解"),
    ];

    public static byte[] ReadShared(string fullPath) => PackageChecks.ReadShared(fullPath);

    /// <summary>驗證結構與大小，回傳存檔時會遺失的內容清單。</summary>
    public static IReadOnlyList<string> Inspect(byte[] bytes, OfficeToolsOptions options)
    {
        var names = PackageChecks.ValidateZip(bytes, options, "[Content_Types].xml", "xl/workbook.xml");

        var warnings = new List<string>();
        foreach (var (prefix, label) in LossyParts)
        {
            if (names.Any(n => n.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) && !warnings.Contains(label))
            {
                warnings.Add(label);
            }
        }

        try
        {
            using var zip = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
            if (HasDrawingShapes(zip))
            {
                warnings.Add("圖形 / 文字方塊");
            }
        }
        catch (InvalidDataException ex)
        {
            throw PackageChecks.Corrupt("檔案結構損壞，無法讀取", ex);
        }

        return warnings;
    }

    public static XLWorkbook Load(byte[] bytes)
    {
        try
        {
            return new XLWorkbook(new MemoryStream(bytes, writable: false));
        }
        catch (Exception ex) when (ex is not OfficeToolException and not OutOfMemoryException)
        {
            throw Corrupt("活頁簿內容無法解析", ex);
        }
    }

    /// <summary>
    /// 先寫到同資料夾的暫存檔，再換掉目標檔，避免寫到一半造成壞檔。
    /// 目標檔被鎖住時回 FILE_LOCKED，原檔不受影響。
    /// </summary>
    /// <remarks>
    /// ClosedXML 0.105 的活頁簿實際上只能存一次：它會記住上一次存檔的串流或檔案，
    /// 第二次存檔時重新開啟，目標已釋放或被換名就會失敗。所以呼叫端必須用回傳的內容
    /// 重新載入一個乾淨的活頁簿（<see cref="Load"/>）取代原本的，之後的存檔才會正常。
    /// </remarks>
    /// <returns>實際寫出的檔案內容。</returns>
    public static byte[] WriteAtomically(IXLWorkbook workbook, string fullPath)
    {
        var dir = Path.GetDirectoryName(fullPath)!;
        var temp = Path.Combine(dir, $".{Path.GetFileNameWithoutExtension(fullPath)}.{Guid.NewGuid():N}.tmp.xlsx"); // ClosedXML 要求副檔名為 .xlsx
        try
        {
            // 用路徑版本；SaveAs(Stream) 會保留傳入的串流，釋放後再存就會丟 ObjectDisposedException
            workbook.SaveAs(temp);
            var written = File.ReadAllBytes(temp);

            if (File.Exists(fullPath))
            {
                File.Replace(temp, fullPath, destinationBackupFileName: null);
            }
            else
            {
                File.Move(temp, fullPath);
            }

            return written;
        }
        catch (DirectoryNotFoundException ex)
        {
            throw new OfficeToolException(ErrorCodes.FileNotFound, $"資料夾不存在：{dir}", "請先建立資料夾，或改存到既有的資料夾", ex);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw Locked(fullPath, ex);
        }
        finally
        {
            if (File.Exists(temp))
            {
                File.Delete(temp);
            }
        }
    }

    /// <summary>
    /// 把活頁簿序列化成 xlsx 內容，不經過磁碟。呼叫後這個 ClosedXML 實例就「用掉了」（SaveAs(Stream) 會保留該串流，
    /// 再存一次會失敗），呼叫端必須用 <see cref="WorkbookSession.ReloadFrom"/> 換新（原因見 <see cref="WriteAtomically"/>）。
    /// </summary>
    public static byte[] Serialize(IXLWorkbook workbook)
    {
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>用 Open XML SDK 驗證指定工作表（含它的表格部件）與活頁簿部件，回傳錯誤描述；空 = 合法。</summary>
    public static IReadOnlyList<string> ValidateSheet(byte[] xlsx, string sheetName)
    {
        using var doc = SpreadsheetDocument.Open(new MemoryStream(xlsx, writable: false), isEditable: false);
        var workbookPart = doc.WorkbookPart!;
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(workbookPart).Select(e => $"{e.Path?.XPath}: {e.Description}").ToList();

        var sheet = workbookPart.Workbook?.Sheets?.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id?.Value is { } relationshipId && workbookPart.GetPartById(relationshipId) is WorksheetPart part)
        {
            errors.AddRange(validator.Validate(part).Select(e => $"{e.Path?.XPath}: {e.Description}"));
            foreach (var table in part.TableDefinitionParts)
            {
                errors.AddRange(validator.Validate(table).Select(e => $"{e.Path?.XPath}: {e.Description}"));
            }
        }

        return errors;
    }

    private static bool HasDrawingShapes(ZipArchive zip)
    {
        foreach (var entry in zip.Entries.Where(e =>
                     e.FullName.StartsWith("xl/drawings/", StringComparison.OrdinalIgnoreCase)
                     && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            using var reader = new StreamReader(entry.Open());
            if (ShapeElement().IsMatch(reader.ReadToEnd()))
            {
                return true;
            }
        }

        return false;
    }

    // <xdr:sp ...>、<xdr:cxnSp>、<xdr:grpSp>：圖形、連接線、群組；圖片（pic）ClosedXML 可保留，不算
    [GeneratedRegex(@"<(\w+:)?(sp|cxnSp|grpSp)[\s>/]")]
    private static partial Regex ShapeElement();

    private static OfficeToolException Locked(string path, Exception inner) =>
        new(
            ErrorCodes.FileLocked,
            $"檔案無法存取（被其他程式鎖定或沒有權限）：{path}",
            "檔案在 Office 中開啟，請關閉或改用 save_as",
            inner);

    private static OfficeToolException Corrupt(string message, Exception? inner) =>
        new(ErrorCodes.CorruptFile, message, "請用 Excel 開啟並修復後另存新檔", inner);
}
