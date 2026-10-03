using System.IO.Compression;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using OfficeTools.Common;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Workspace;

/// <summary>xlsx 檔案層級的工作：共享讀取、結構檢查（含 ClosedXML 無法保留的內容）、載入、原子寫入。</summary>
internal static partial class WorkbookPackage
{
    private static readonly byte[] OleSignature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    private static readonly byte[] ZipSignature = [0x50, 0x4B, 0x03, 0x04];

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

    /// <summary>讀進記憶體後立即放開檔案，使用者可繼續在 Excel 中開啟與編輯。</summary>
    public static byte[] ReadShared(string fullPath)
    {
        try
        {
            using var fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var ms = new MemoryStream();
            fs.CopyTo(ms);
            return ms.ToArray();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw Locked(fullPath, ex);
        }
    }

    /// <summary>驗證結構與大小，回傳存檔時會遺失的內容清單。</summary>
    public static IReadOnlyList<string> Inspect(byte[] bytes, OfficeToolsOptions options)
    {
        if (bytes.AsSpan().StartsWith(OleSignature))
        {
            throw new OfficeToolException(
                ErrorCodes.PasswordProtected,
                "檔案有密碼保護，或是舊版 .xls 格式",
                "請使用者先在 Excel 移除密碼，或另存為 .xlsx 後再處理");
        }

        if (!bytes.AsSpan().StartsWith(ZipSignature))
        {
            throw Corrupt("檔案不是有效的 xlsx（不是 zip 封裝）", null);
        }

        try
        {
            using var zip = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);

            var total = zip.Entries.Sum(e => e.Length);
            if (total > options.MaxUncompressedMb * 1024L * 1024)
            {
                throw new OfficeToolException(
                    ErrorCodes.FileTooLarge,
                    $"檔案解壓縮後約 {total / (1024 * 1024)} MB，超過上限 {options.MaxUncompressedMb} MB",
                    "請先將檔案分割或縮小後再處理");
            }

            var names = zip.Entries.Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (!names.Contains("[Content_Types].xml") || !names.Contains("xl/workbook.xml"))
            {
                throw Corrupt("檔案缺少活頁簿的必要部件，不是有效的 xlsx", null);
            }

            var warnings = new List<string>();
            foreach (var (prefix, label) in LossyParts)
            {
                if (names.Any(n => n.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) && !warnings.Contains(label))
                {
                    warnings.Add(label);
                }
            }

            if (HasDrawingShapes(zip))
            {
                warnings.Add("圖形 / 文字方塊");
            }

            return warnings;
        }
        catch (InvalidDataException ex)
        {
            throw Corrupt("檔案結構損壞，無法讀取", ex);
        }
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
