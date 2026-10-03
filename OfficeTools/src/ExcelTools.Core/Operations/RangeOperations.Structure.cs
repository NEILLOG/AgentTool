using ClosedXML.Excel;
using ExcelTools.Core.Internal;
using ExcelTools.Core.Models;
using ExcelTools.Core.Workspace;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Operations;

// 插入 / 刪除列欄。ClosedXML 在這類操作上有不少地雷（見各處註解），所以每次操作都有檢查點，
// 操作後序列化並驗證，有問題就復原，不會把壞檔留在 session 裡。
public sealed partial class RangeOperations
{
    /// <summary>在第 <paramref name="row"/> 列的上方插入 <paramref name="count"/> 列（與 Excel 的「插入」一致）。</summary>
    public StructureChangeResult InsertRows(string workbookId, string sheet, int row, int count = 1) =>
        ChangeStructure(workbookId, sheet, rows: true, insert: true, row, count);

    /// <summary>刪除從第 <paramref name="row"/> 列開始的 <paramref name="count"/> 列，下方的列上移。</summary>
    public StructureChangeResult DeleteRows(string workbookId, string sheet, int row, int count = 1) =>
        ChangeStructure(workbookId, sheet, rows: true, insert: false, row, count);

    /// <summary>在 <paramref name="column"/>（欄名稱，例如 "C"）的左邊插入 <paramref name="count"/> 欄。</summary>
    public StructureChangeResult InsertColumns(string workbookId, string sheet, string column, int count = 1) =>
        ChangeStructure(workbookId, sheet, rows: false, insert: true, ParseColumn(column), count);

    /// <summary>刪除從 <paramref name="column"/> 開始的 <paramref name="count"/> 欄，右邊的欄左移。</summary>
    public StructureChangeResult DeleteColumns(string workbookId, string sheet, string column, int count = 1) =>
        ChangeStructure(workbookId, sheet, rows: false, insert: false, ParseColumn(column), count);

    private StructureChangeResult ChangeStructure(string workbookId, string sheet, bool rows, bool insert, int index, int count)
    {
        // 不用 mutates:true：操作被復原時不該把活頁簿標成有變更，所以成功後才自己標記
        return sessions.Use(workbookId, mutates: false, s =>
        {
            if (s.ReadOnly)
            {
                throw WorkbookSessionManager.ReadOnlyError();
            }

            var ws = SheetLookup.Find(s.Workbook, sheet);
            var max = rows ? A1Address.MaxRow : A1Address.MaxColumn;
            var unit = rows ? "列" : "欄";
            var start = index;
            if (start < 1 || start > max)
            {
                throw new OfficeToolException(ErrorCodes.InvalidRange, $"{unit}位置 {index} 超出範圍（1 到 {max}）", rows ? "列號是 1 起算的整數" : "欄請用欄名稱，例如 C");
            }

            if (count < 1 || start + count - 1 > max)
            {
                throw new OfficeToolException(ErrorCodes.InvalidRange, $"{unit}數 {count} 無效：從 {start} 開始最多可處理 {max - start + 1} {unit}", $"{unit}數至少是 1");
            }

            var last = start + count - 1;
            var band = rows ? $"{start}:{last}" : $"{A1Address.ColumnName(start)}:{A1Address.ColumnName(last)}";
            var warnings = new List<string>();

            if (insert)
            {
                RejectIfContentWouldFallOffTheSheet(ws, rows, count, max, unit);
            }
            else
            {
                var (locations, total) = FormulaReferences.FindDangling(s.Workbook, ws, rows, start, last);
                if (total > 0)
                {
                    var more = total > locations.Count ? $" 等共 {total} 處" : string.Empty;
                    warnings.Add(
                        $"有 {total} 處公式或名稱的參照落在被刪除的{unit}內：{string.Join("、", locations)}{more}。" +
                        "Excel 會把它們變成 #REF!，但這裡不會——它們現在指向別的儲存格（或維持原樣），請檢查並修正");
                }
            }

            var before = WorkbookPackage.Serialize(s.Workbook);
            s.ReloadFrom(before); // 序列化會「用掉」目前的實例，換一個乾淨的來操作
            try
            {
                var target = SheetLookup.Find(s.Workbook, sheet);
                Apply(target, rows, insert, start, count);

                var after = WorkbookPackage.Serialize(s.Workbook);
                var problems = WorkbookPackage.ValidateSheet(after, target.Name);
                if (problems.Count > 0)
                {
                    throw new OfficeToolException(
                        ErrorCodes.UnsafeOperation,
                        $"這個操作會讓檔案損壞，已復原（{problems[0]}）",
                        "ClosedXML 對這張工作表上的表格、資料驗證或條件式格式處理有限制；請改成較小的操作，或先在 Excel 中調整");
                }

                s.ReloadFrom(after);
                s.IsDirty = true;
                var used = s.Workbook.Worksheet(target.Name).RangeUsed(XLCellsUsedOptions.Contents)?.RangeAddress.ToStringRelative(false);
                return new StructureChangeResult(target.Name, band, used, warnings);
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                s.ReloadFrom(before); // 復原
                if (ex is OfficeToolException)
                {
                    throw;
                }

                throw new OfficeToolException(
                    ErrorCodes.UnsafeOperation,
                    $"無法完成這個操作，已復原（{ex.GetType().Name}：{ex.Message}）",
                    "ClosedXML 對這張工作表上的表格、資料驗證或條件式格式處理有限制；請改成較小的操作，或先在 Excel 中調整",
                    ex);
            }
        });
    }

    private static void Apply(IXLWorksheet ws, bool rows, bool insert, int start, int count)
    {
        if (rows)
        {
            if (insert)
            {
                ws.Row(start).InsertRowsAbove(count);
            }
            else
            {
                ws.Rows(start, start + count - 1).Delete();
            }

            return;
        }

        if (insert)
        {
            ws.Column(start).InsertColumnsBefore(count);
        }
        else
        {
            ws.Columns(start, start + count - 1).Delete();
        }
    }

    /// <summary>ClosedXML 插入列欄時，擠出工作表邊界的內容會悄悄消失；Excel 則會拒絕這個操作。</summary>
    private static void RejectIfContentWouldFallOffTheSheet(IXLWorksheet ws, bool rows, int count, int max, string unit)
    {
        var used = ws.RangeUsed(XLCellsUsedOptions.All)?.RangeAddress;
        if (used is null)
        {
            return;
        }

        var lastUsed = rows ? used.LastAddress.RowNumber : used.LastAddress.ColumnNumber;
        if (lastUsed + count > max)
        {
            throw new OfficeToolException(
                ErrorCodes.InvalidRange,
                $"插入 {count} {unit}會把最後{count}{unit}內的儲存格擠出工作表邊界（內容會遺失）",
                $"請先清除工作表最後的{unit}，或減少插入的{unit}數");
        }
    }

    private static int ParseColumn(string column)
    {
        var letters = column?.Trim().TrimStart('$') ?? string.Empty;
        if (letters.Length == 0 || !letters.All(char.IsAsciiLetter))
        {
            throw new OfficeToolException(ErrorCodes.InvalidRange, $"無效的欄「{column}」", "欄請用欄名稱，例如 A、C、AB");
        }

        return A1Address.ParseRange($"{letters}:{letters}").FirstColumn;
    }
}
