using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Internal;

/// <summary>
/// 寫入前檢查公式語法。ClosedXML 設定公式時完全不驗證，壞公式（例如 SUM(A1:）會被存進檔案，
/// Excel 開啟時要求修復。這裡在獨立的暫存活頁簿裡試著解析，只攔語法錯誤。
/// </summary>
/// <remarks>
/// 不攔的情況：未知函式（得到 #NAME?，Excel 本身也是如此）、參照不存在的工作表、循環參照、
/// ClosedXML 尚未支援的運算（例如範圍交集）——這些在 Excel 中語法都合法，寫入後由警告回報。
/// </remarks>
internal sealed class FormulaValidator : IDisposable
{
    private readonly XLWorkbook _scratch = new();
    private readonly IXLCell _cell;

    public FormulaValidator()
    {
        _cell = _scratch.AddWorksheet("Scratch").Cell("ZZ1000"); // 遠離 A1，避免公式剛好自我參照
    }

    /// <summary>語法錯誤時丟 INVALID_VALUE。</summary>
    public void Validate(string address, string formula)
    {
        try
        {
            _cell.FormulaA1 = formula;
            _ = _cell.Value; // 讀取時才會真正解析
        }
        catch (ExpressionParseException ex)
        {
            throw new OfficeToolException(
                ErrorCodes.InvalidValue,
                $"{address} 的公式語法錯誤：={formula}（{ex.Message}）",
                "請檢查括號、引號與範圍是否完整；公式要用英文函式名稱與逗號分隔參數，例如 =SUM(A1:A10)。這次沒有寫入任何資料");
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            // 語法合法但無法計算（循環參照、尚未支援的運算等）：寫入後由警告回報
        }
    }

    public void Dispose() => _scratch.Dispose();
}
