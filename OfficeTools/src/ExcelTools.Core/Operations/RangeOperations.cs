using System.Globalization;
using ClosedXML.Excel;
using ExcelTools.Core.Internal;
using ExcelTools.Core.Models;
using ExcelTools.Core.Workspace;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Operations;

/// <summary>儲存格範圍的讀取、寫入、追加與清除。範圍一律是 A1 格式（A1、A1:C10、A:A、1:1）。</summary>
public sealed partial class RangeOperations(WorkbookSessionManager sessions, ExcelToolsOptions options)
{
    internal const int MaxWarnings = 20;

    /// <summary>
    /// 讀取範圍。結果限制在工作表有內容的區域內（起點維持你要求的位置，終點縮到最後有內容的列 / 欄）；
    /// 超過 <see cref="ExcelToolsOptions.MaxCellsPerRead"/> 時以整列為單位截斷（至少回傳一整列），並回傳 NextRange。
    /// </summary>
    public RangeData ReadRange(string workbookId, string sheet, string range, ReadOptions? readOptions = null) =>
        sessions.Use(workbookId, mutates: false, s =>
        {
            var opts = readOptions ?? new ReadOptions();
            var ws = SheetLookup.Find(s.Workbook, sheet);
            var requested = A1Address.ParseRange(range);

            var usedBounds = ws.RangeUsed(XLCellsUsedOptions.Contents)?.RangeAddress;
            var used = usedBounds?.ToStringRelative(false);
            var lastRow = usedBounds?.LastAddress.RowNumber ?? 0;
            var lastColumn = usedBounds?.LastAddress.ColumnNumber ?? 0;

            var lastRowToRead = Math.Min(requested.LastRow, lastRow);
            var lastColumnToRead = Math.Min(requested.LastColumn, lastColumn);
            if (lastRowToRead < requested.FirstRow || lastColumnToRead < requested.FirstColumn)
            {
                return new RangeData(ws.Name, AsCells(requested).ToString(), used, [], null, false, null, []);
            }

            var columns = lastColumnToRead - requested.FirstColumn + 1;
            var totalRows = lastRowToRead - requested.FirstRow + 1;
            var rows = Math.Min(totalRows, Math.Max(1, options.MaxCellsPerRead / columns));
            var truncated = rows < totalRows;

            var firstRow = requested.FirstRow;
            var firstColumn = requested.FirstColumn;
            var returned = new RangeAddress(firstRow, firstColumn, firstRow + rows - 1, lastColumnToRead);

            var values = new object?[rows][];
            var formulas = opts.IncludeFormulas ? new string?[rows][] : null;
            for (var r = 0; r < rows; r++)
            {
                values[r] = new object?[columns];
                if (formulas is not null)
                {
                    formulas[r] = new string?[columns];
                }
            }

            var warnings = new List<string>();
            var warningCount = 0;
            var block = ws.Range(returned.FirstRow, returned.FirstColumn, returned.LastRow, returned.LastColumn);
            foreach (var cell in block.CellsUsed(XLCellsUsedOptions.Contents)) // 只走有內容的格子，不會把空格子建出來
            {
                var r = cell.Address.RowNumber - firstRow;
                var c = cell.Address.ColumnNumber - firstColumn;
                if (r < 0 || r >= rows || c < 0 || c >= columns)
                {
                    continue; // ClosedXML 會把範圍擴張到涵蓋跨出範圍的合併儲存格，超出要求的部分不回傳
                }

                values[r][c] = ReadCell(cell, opts.UseFormattedText, warnings, ref warningCount);
                if (formulas is not null && cell.HasFormula)
                {
                    formulas[r][c] = "=" + cell.FormulaA1;
                }
            }

            AddOverflowNote(warnings, warningCount);
            var next = truncated
                ? new RangeAddress(firstRow + rows, firstColumn, lastRowToRead, lastColumnToRead).ToString()
                : null;
            return new RangeData(ws.Name, returned.ToString(), used, values, formulas, truncated, next, warnings);
        });

    /// <summary>
    /// 從 <paramref name="range"/> 的左上角開始寫入 <paramref name="values"/>。<paramref name="range"/> 可以是單一儲存格，
    /// 或與資料大小完全相同的範圍。資料列可以長短不一；null 代表清空該格。全有或全無：任何一格有問題就什麼都不寫。
    /// </summary>
    public RangeChangeResult WriteRange(string workbookId, string sheet, string range, object?[][] values, WriteOptions? writeOptions = null) =>
        sessions.Use(workbookId, mutates: true, s =>
        {
            var ws = SheetLookup.Find(s.Workbook, sheet);
            var target = A1Address.ParseRange(range);
            if (target.Kind != RangeKind.Cells)
            {
                throw new OfficeToolException(ErrorCodes.InvalidRange, $"寫入的位置不可以是整欄或整列：{range}", "請指定起始儲存格（例如 A1）或與資料大小相同的範圍（例如 A1:C3）");
            }

            return WriteBlock(ws, target, values, writeOptions ?? new WriteOptions(), mustMatchRange: target.CellCount != 1);
        });

    /// <summary>把資料加在工作表現有內容的下方，欄位預設對齊有內容區域的第一欄。</summary>
    /// <param name="startColumn">起始欄（例如 "C"）；省略則用有內容區域的第一欄，空工作表從 A 開始。</param>
    public RangeChangeResult AppendRows(string workbookId, string sheet, object?[][] rows, string? startColumn = null, WriteOptions? writeOptions = null) =>
        sessions.Use(workbookId, mutates: true, s =>
        {
            var ws = SheetLookup.Find(s.Workbook, sheet);
            var used = ws.RangeUsed(XLCellsUsedOptions.Contents)?.RangeAddress;

            var column = used?.FirstAddress.ColumnNumber ?? 1;
            if (startColumn is not null)
            {
                var letters = startColumn.Trim().TrimStart('$');
                if (letters.Length == 0 || !letters.All(char.IsAsciiLetter))
                {
                    throw new OfficeToolException(ErrorCodes.InvalidRange, $"無效的起始欄「{startColumn}」", "起始欄只能是欄名稱，例如 A、C、AB");
                }

                column = A1Address.ParseRange($"{letters}:{letters}").FirstColumn; // 超過 XFD 會在這裡回 INVALID_RANGE
            }

            var firstRow = (used?.LastAddress.RowNumber ?? 0) + 1;
            return WriteBlock(ws, new RangeAddress(firstRow, column, firstRow, column), rows, writeOptions ?? new WriteOptions(), mustMatchRange: false);
        });

    /// <summary>清除範圍（預設只清內容，保留格式）。範圍會縮到工作表實際使用的區域，所以清整欄整列也很快。</summary>
    public RangeChangeResult ClearRange(string workbookId, string sheet, string range, ClearMode mode = ClearMode.Contents) =>
        sessions.Use(workbookId, mutates: true, s =>
        {
            var ws = SheetLookup.Find(s.Workbook, sheet);
            var requested = A1Address.ParseRange(range);

            // 含格式的邊界：只有格式的格子也算在內
            var bounds = ws.RangeUsed(XLCellsUsedOptions.All)?.RangeAddress;
            var lastRow = Math.Min(requested.LastRow, bounds?.LastAddress.RowNumber ?? 0);
            var lastColumn = Math.Min(requested.LastColumn, bounds?.LastAddress.ColumnNumber ?? 0);
            if (lastRow < requested.FirstRow || lastColumn < requested.FirstColumn)
            {
                return new RangeChangeResult(ws.Name, AsCells(requested).ToString(), 0, []);
            }

            var cleared = new RangeAddress(requested.FirstRow, requested.FirstColumn, lastRow, lastColumn);
            ws.Range(cleared.FirstRow, cleared.FirstColumn, cleared.LastRow, cleared.LastColumn).Clear(mode switch
            {
                ClearMode.Formats => XLClearOptions.AllFormats,
                ClearMode.All => XLClearOptions.All,
                _ => XLClearOptions.Contents,
            });
            return new RangeChangeResult(ws.Name, cleared.ToString(), checked((int)cleared.CellCount), []);
        });

    private RangeChangeResult WriteBlock(IXLWorksheet ws, RangeAddress target, object?[][] values, WriteOptions writeOptions, bool mustMatchRange)
    {
        ArgumentNullException.ThrowIfNull(values);

        var height = values.Length;
        var width = values.Length == 0 ? 0 : values.Max(row => row?.Length ?? 0);
        if (height == 0 || width == 0)
        {
            throw new OfficeToolException(ErrorCodes.InvalidValue, "沒有要寫入的資料", "values 要是二維陣列，例如 [[\"a\", 1], [\"b\", 2]]");
        }

        if ((long)height * width > options.MaxCellsPerWrite)
        {
            throw new OfficeToolException(
                ErrorCodes.InvalidValue,
                $"一次寫入 {(long)height * width} 格，超過上限 {options.MaxCellsPerWrite}",
                "請分成多次寫入，或用 append_rows 分批加入");
        }

        var firstRow = target.FirstRow;
        var firstColumn = target.FirstColumn;
        var lastRow = firstRow + height - 1;
        var lastColumn = firstColumn + width - 1;
        if (lastRow > A1Address.MaxRow || lastColumn > A1Address.MaxColumn)
        {
            throw new OfficeToolException(
                ErrorCodes.InvalidRange,
                $"從 {new CellAddress(firstRow, firstColumn)} 開始寫入 {height} 列 × {width} 欄會超出工作表邊界（最大 XFD1048576）",
                "請改從較前面的位置開始，或減少資料量");
        }

        if (mustMatchRange && (target.RowCount != height || target.ColumnCount != width))
        {
            throw new OfficeToolException(
                ErrorCodes.InvalidRange,
                $"範圍 {target} 是 {target.RowCount} 列 × {target.ColumnCount} 欄，但資料是 {height} 列 × {width} 欄",
                "範圍大小要和資料相同；只想指定起點的話，直接給單一儲存格（例如 A1）");
        }

        // 先全部轉換與驗證，確認沒問題才動工作表
        var cellOptions = new CellWriteOptions(writeOptions.ParseFormulas, writeOptions.ParseIsoDates);
        var inputs = new CellInput?[height][];
        var formulas = new List<(string Address, string Formula)>();
        for (var r = 0; r < height; r++)
        {
            var row = values[r];
            inputs[r] = new CellInput?[width];
            for (var c = 0; row is not null && c < row.Length; c++)
            {
                var address = new CellAddress(firstRow + r, firstColumn + c).ToString();
                CellInput input;
                try
                {
                    input = CellValueConverter.ToCellInput(row[c], cellOptions);
                }
                catch (OfficeToolException ex)
                {
                    throw new OfficeToolException(ex.Code, $"{address}：{ex.Message}", ex.Hint, ex);
                }

                inputs[r][c] = input;
                if (input.Kind == CellInputKind.Formula)
                {
                    formulas.Add((address, input.Formula!));
                }
            }
        }

        using (var validator = new FormulaValidator())
        {
            foreach (var (address, formula) in formulas)
            {
                validator.Validate(address, formula);
            }
        }

        for (var r = 0; r < height; r++)
        {
            for (var c = 0; c < width; c++)
            {
                if (inputs[r][c] is { } input)
                {
                    CellValueConverter.Apply(ws.Cell(firstRow + r, firstColumn + c), input);
                }
            }
        }

        var written = new RangeAddress(firstRow, firstColumn, lastRow, lastColumn);
        return new RangeChangeResult(ws.Name, written.ToString(), height * width, CheckFormulaResults(ws, firstRow, firstColumn, inputs));
    }

    /// <summary>寫入後試算剛寫的公式：算出錯誤值或無法計算時提醒，不影響寫入結果。</summary>
    private static List<string> CheckFormulaResults(IXLWorksheet ws, int firstRow, int firstColumn, CellInput?[][] inputs)
    {
        var warnings = new List<string>();
        var count = 0;
        for (var r = 0; r < inputs.Length; r++)
        {
            for (var c = 0; c < inputs[r].Length; c++)
            {
                if (inputs[r][c] is not { Kind: CellInputKind.Formula } input)
                {
                    continue;
                }

                var address = new CellAddress(firstRow + r, firstColumn + c);
                string? problem = null;
                try
                {
                    var value = ws.Cell(address.Row, address.Column).Value;
                    if (value.IsError)
                    {
                        problem = $"結果是 {CellValueConverter.ToSerializable(value)}";
                    }
                }
                catch (Exception ex) when (ex is not (OutOfMemoryException or OfficeToolException))
                {
                    problem = $"無法計算（{Describe(ex)}）";
                }

                if (problem is not null)
                {
                    count++;
                    if (warnings.Count < MaxWarnings)
                    {
                        warnings.Add($"{address}：公式 ={input.Formula} {problem}");
                    }
                }
            }
        }

        AddOverflowNote(warnings, count);
        return warnings;
    }

    internal static object? ReadCell(IXLCell cell, bool formatted, List<string> warnings, ref int warningCount)
    {
        try
        {
            var value = cell.Value;
            if (formatted)
            {
                return value.IsBlank ? null : cell.GetFormattedString(CultureInfo.InvariantCulture); // 不隨使用者的地區設定改變
            }

            return CellValueConverter.ToSerializable(value);
        }
        catch (Exception ex) when (ex is not (OutOfMemoryException or OfficeToolException))
        {
            // 公式語法錯誤、循環參照、不支援的運算：改回傳檔案裡存的快取值，不讓整個讀取失敗
            warningCount++;
            if (warnings.Count < MaxWarnings)
            {
                warnings.Add($"{cell.Address.ToStringRelative(false)}：公式 ={cell.FormulaA1} 無法計算（{Describe(ex)}），回傳檔案中的快取值");
            }

            var cached = cell.CachedValue;
            return formatted ? (cached.IsBlank ? null : cached.ToString(CultureInfo.InvariantCulture)) : CellValueConverter.ToSerializable(cached);
        }
    }

    private static string Describe(Exception ex) => ex switch
    {
        InvalidOperationException e when e.Message.Contains("cycle", StringComparison.OrdinalIgnoreCase) => "循環參照",
        NotImplementedException => "ClosedXML 尚未支援這個運算",
        _ when ex.GetType().Name == "ExpressionParseException" => "公式語法錯誤",
        _ => ex.GetType().Name,
    };

    internal static void AddOverflowNote(List<string> warnings, int total)
    {
        if (total > warnings.Count)
        {
            warnings.Add($"…另外還有 {total - warnings.Count} 個類似的問題未列出");
        }
    }

    internal static RangeAddress AsCells(RangeAddress a) => a with { Kind = RangeKind.Cells };
}
