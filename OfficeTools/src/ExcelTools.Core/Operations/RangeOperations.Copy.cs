using ClosedXML.Excel;
using ExcelTools.Core.Internal;
using ExcelTools.Core.Models;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Operations;

public sealed partial class RangeOperations
{
    private const string TempSheetName = "__officetools_copy__";

    /// <summary>
    /// 複製範圍到 <paramref name="target"/>（目標範圍的左上角，單一儲存格），可複製到同一活頁簿的另一張工作表。
    /// 來源會縮到實際使用的區域；會覆蓋目標範圍內的既有內容，覆蓋了有內容的儲存格時會在警告中說明。
    /// 來源與目標重疊時結果與 Excel 相同（先取得來源的快照再貼上）。
    /// </summary>
    public RangeChangeResult CopyRange(string workbookId, string sheet, string source, string target, string? targetSheet = null, CopyMode mode = CopyMode.All) =>
        sessions.Use(workbookId, mutates: true, s =>
        {
            var wb = s.Workbook;
            var from = SheetLookup.Find(wb, sheet);
            var to = targetSheet is null ? from : SheetLookup.Find(wb, targetSheet);
            var requested = A1Address.ParseRange(source);
            var anchor = A1Address.ParseCell(target);

            var bounds = from.RangeUsed(XLCellsUsedOptions.All)?.RangeAddress;
            var lastRow = Math.Min(requested.LastRow, bounds?.LastAddress.RowNumber ?? 0);
            var lastColumn = Math.Min(requested.LastColumn, bounds?.LastAddress.ColumnNumber ?? 0);
            if (lastRow < requested.FirstRow || lastColumn < requested.FirstColumn)
            {
                throw new OfficeToolException(ErrorCodes.InvalidValue, $"來源範圍 {AsCells(requested)} 沒有任何內容或格式", "請確認來源範圍；可用 read_range 查看有內容的區域");
            }

            var src = new RangeAddress(requested.FirstRow, requested.FirstColumn, lastRow, lastColumn);
            var endRow = anchor.Row + src.RowCount - 1;
            var endColumn = anchor.Column + src.ColumnCount - 1;
            if (endRow > A1Address.MaxRow || endColumn > A1Address.MaxColumn)
            {
                throw new OfficeToolException(
                    ErrorCodes.InvalidRange,
                    $"從 {anchor} 開始貼上 {src.RowCount} 列 × {src.ColumnCount} 欄會超出工作表邊界（最大 XFD1048576）",
                    "請改貼到較前面的位置");
            }

            var dest = new RangeAddress(anchor.Row, anchor.Column, endRow, endColumn);
            var overwritten = to.Range(dest.FirstRow, dest.FirstColumn, dest.LastRow, dest.LastColumn).CellsUsed(XLCellsUsedOptions.Contents).Count();

            var sourceRange = from.Range(src.FirstRow, src.FirstColumn, src.LastRow, src.LastColumn);
            var destCell = to.Cell(dest.FirstRow, dest.FirstColumn);
            var overlaps = ReferenceEquals(from, to)
                && src.FirstRow <= dest.LastRow && dest.FirstRow <= src.LastRow
                && src.FirstColumn <= dest.LastColumn && dest.FirstColumn <= src.LastColumn;

            if (mode == CopyMode.ValuesOnly)
            {
                CopyValues(sourceRange, to, dest);
            }
            else if (overlaps)
            {
                CopyViaTempSheet(wb, sourceRange, destCell);
            }
            else
            {
                sourceRange.CopyTo(destCell);
            }

            var warnings = new List<string>();
            if (overwritten > 0)
            {
                warnings.Add($"目標範圍 {dest} 內原有 {overwritten} 個有內容的儲存格已被覆蓋");
            }

            return new RangeChangeResult(to.Name, dest.ToString(), checked((int)dest.CellCount), warnings);
        });

    /// <summary>只貼值：先把來源整個讀成快照，再寫出，所以來源與目標重疊也正確。</summary>
    private static void CopyValues(IXLRange source, IXLWorksheet targetSheet, RangeAddress dest)
    {
        var rows = source.RowCount();
        var columns = source.ColumnCount();
        var snapshot = new XLCellValue[rows, columns];
        foreach (var cell in source.CellsUsed(XLCellsUsedOptions.Contents))
        {
            var r = cell.Address.RowNumber - source.RangeAddress.FirstAddress.RowNumber;
            var c = cell.Address.ColumnNumber - source.RangeAddress.FirstAddress.ColumnNumber;
            try
            {
                snapshot[r, c] = cell.Value;
            }
            catch (Exception ex) when (ex is not (OutOfMemoryException or OfficeToolException))
            {
                snapshot[r, c] = cell.CachedValue; // 公式算不出來：改用檔案裡的快取值
            }
        }

        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
            {
                var cell = targetSheet.Cell(dest.FirstRow + r, dest.FirstColumn + c);
                if (snapshot[r, c].IsBlank)
                {
                    cell.Clear(XLClearOptions.Contents);
                }
                else
                {
                    cell.Value = snapshot[r, c];
                }
            }
        }
    }

    /// <summary>
    /// 來源與目標重疊時，ClosedXML 直接複製會讀到已被覆蓋的格子。先複製到暫存工作表再貼回去。
    /// 暫存區刻意使用與來源「相同的座標」：第一步位移為零（相對參照原封不動），第二步才套用完整的
    /// 「來源 → 目標」位移，與直接複製的結果相同。若放在 A1，第一步的位移可能把相對參照推出工作表邊界而變成 #REF!。
    /// </summary>
    private static void CopyViaTempSheet(XLWorkbook wb, IXLRange source, IXLCell destination)
    {
        var temp = wb.AddWorksheet(TempSheetName);
        try
        {
            var first = source.RangeAddress.FirstAddress;
            var last = source.RangeAddress.LastAddress;
            source.CopyTo(temp.Cell(first.RowNumber, first.ColumnNumber));
            temp.Range(first.RowNumber, first.ColumnNumber, last.RowNumber, last.ColumnNumber).CopyTo(destination);
        }
        finally
        {
            temp.Delete();
        }
    }
}
