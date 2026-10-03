using System.Globalization;
using ClosedXML.Excel;
using ExcelTools.Core.Internal;
using ExcelTools.Core.Models;
using ExcelTools.Core.Workspace;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Operations;

/// <summary>格式、欄寬、合併儲存格與凍結窗格。</summary>
public sealed class FormatOperations(WorkbookSessionManager sessions, ExcelToolsOptions options)
{
    private const int MaxListedColumns = 200;
    private const int MaxFormattedRows = 10_000;
    private const int MaxFrozenRows = 1_000;
    private const int MaxFrozenColumns = 100;
    private const double MaxColumnWidth = 255;

    private const string WidthNote = "自動調整的寬度是估算值（全形字算 2 個單位，不依賴作業系統字型），與 Excel 實際的自動調整可能略有差異";

    /// <summary>
    /// 套用格式，只改 <paramref name="spec"/> 裡指定的屬性。整欄、整列、整張表（例如 <c>A:A</c>、<c>1:1</c>）
    /// 以欄 / 列 / 工作表樣式處理（已有的與之後新增的儲存格都會套用）；其他矩形範圍受 <see cref="ExcelToolsOptions.MaxCellsPerFormat"/> 限制。
    /// 全有或全無：格式有任何問題就什麼都不套用。
    /// </summary>
    public FormatResult FormatRange(string workbookId, string sheet, string range, FormatSpec spec)
    {
        var format = FormatRules.Parse(spec);
        var target = A1Address.ParseRange(range);

        return sessions.Use(workbookId, mutates: true, s =>
        {
            var ws = SheetLookup.Find(s.Workbook, sheet);
            var allRows = target.FirstRow == 1 && target.LastRow == A1Address.MaxRow;
            var allColumns = target.FirstColumn == 1 && target.LastColumn == A1Address.MaxColumn;

            if (format.OutlineOnly && (allRows || allColumns))
            {
                throw Invalid("BorderSides = outline 不適用於整欄、整列或整張工作表", "請改用 all，或指定有限的矩形範圍");
            }

            if (allRows && allColumns)
            {
                Apply(ws.Style, format);
            }
            else if (allRows)
            {
                Apply(ws.Columns(target.FirstColumn, target.LastColumn).Style, format);
            }
            else if (allColumns)
            {
                if (target.RowCount > MaxFormattedRows)
                {
                    throw new OfficeToolException(ErrorCodes.InvalidRange, $"一次最多格式化 {MaxFormattedRows} 列整列（要求 {target.RowCount} 列）", "請縮小範圍，或改用整欄範圍（例如 A:Z）");
                }

                Apply(ws.Rows(target.FirstRow, target.LastRow).Style, format);
            }
            else
            {
                if (target.CellCount > options.MaxCellsPerFormat)
                {
                    throw new OfficeToolException(
                        ErrorCodes.InvalidRange,
                        $"範圍 {target} 有 {target.CellCount} 格，超過單次格式化上限 {options.MaxCellsPerFormat}",
                        "請縮小範圍，或改用整欄 / 整列範圍（例如 A:A、1:1）");
                }

                Apply(ws.Range(target.FirstRow, target.FirstColumn, target.LastRow, target.LastColumn).Style, format);
            }

            return new FormatResult(ws.Name, target.ToString(), target.CellCount, []);
        });
    }

    /// <summary>設定欄寬（Excel 的字元寬度單位，0 到 255；0 會隱藏該欄）。<paramref name="columns"/> 例如 "A" 或 "A:C"。</summary>
    public ColumnWidthResult SetColumnWidth(string workbookId, string sheet, string columns, double width)
    {
        if (double.IsNaN(width) || width < 0 || width > MaxColumnWidth)
        {
            throw Invalid($"無效的欄寬 {width}", $"欄寬要介於 0 到 {MaxColumnWidth}（單位是字元寬度，預設約 8.43）");
        }

        var span = ParseColumns(columns);
        return sessions.Use(workbookId, mutates: true, s =>
        {
            var ws = SheetLookup.Find(s.Workbook, sheet);
            ws.Columns(span.First, span.Last).Width = width;

            var notes = width == 0 ? (IReadOnlyList<string>)["寬度 0 會隱藏這些欄"] : [];
            return new ColumnWidthResult(ws.Name, Widths(span.First, span.Last, _ => width), [], notes);
        });
    }

    /// <summary>
    /// 依內容自動調整欄寬。省略 <paramref name="columns"/> 時處理所有有內容的欄。
    /// 以顯示文字估算（含數字格式），略過自動換行與合併的儲存格；沒有內容的欄維持原寬度。
    /// </summary>
    public ColumnWidthResult AutoFitColumns(string workbookId, string sheet, string? columns = null, double minWidth = 3, double maxWidth = 100)
    {
        if (double.IsNaN(minWidth) || double.IsNaN(maxWidth) || minWidth < 0 || maxWidth > MaxColumnWidth || minWidth > maxWidth)
        {
            throw Invalid($"無效的寬度範圍 {minWidth} 到 {maxWidth}", $"要滿足 0 ≤ minWidth ≤ maxWidth ≤ {MaxColumnWidth}");
        }

        var requested = columns is null ? (ColumnSpan?)null : ParseColumns(columns);
        return sessions.Use(workbookId, mutates: true, s =>
        {
            var ws = SheetLookup.Find(s.Workbook, sheet);
            var used = ws.RangeUsed(XLCellsUsedOptions.Contents)?.RangeAddress;
            if (used is null)
            {
                return new ColumnWidthResult(ws.Name, [], [], [WidthNote]);
            }

            var span = requested ?? new ColumnSpan(used.FirstAddress.ColumnNumber, used.LastAddress.ColumnNumber);
            var widest = new Dictionary<int, double>();
            foreach (var cell in ws.Range(1, span.First, A1Address.MaxRow, span.Last).CellsUsed(XLCellsUsedOptions.Contents))
            {
                if (cell.IsMerged() || cell.Style.Alignment.WrapText)
                {
                    continue;
                }

                var text = DisplayText(cell);
                if (text.Length == 0)
                {
                    continue;
                }

                var (size, bold) = ColumnWidthEstimator.FontOf(cell);
                var estimate = ColumnWidthEstimator.Estimate(text, size, bold);
                var column = cell.Address.ColumnNumber;
                widest[column] = Math.Max(widest.GetValueOrDefault(column), estimate);
            }

            var set = new List<ColumnWidth>();
            var unchanged = new List<string>();
            for (var c = span.First; c <= span.Last; c++)
            {
                if (widest.TryGetValue(c, out var width) && width > 0)
                {
                    var clamped = Math.Round(Math.Clamp(width, minWidth, maxWidth), 2);
                    ws.Column(c).Width = clamped;
                    if (set.Count < MaxListedColumns)
                    {
                        set.Add(new ColumnWidth(A1Address.ColumnName(c), clamped));
                    }
                }
                else if (unchanged.Count < MaxListedColumns)
                {
                    unchanged.Add(A1Address.ColumnName(c));
                }
            }

            return new ColumnWidthResult(ws.Name, set, unchanged, [WidthNote]);
        });
    }

    /// <summary>
    /// 合併儲存格，內容保留在左上角。合併會清掉其他儲存格的內容，所以那些格子有內容時預設拒絕，
    /// 要帶 <paramref name="discardOtherValues"/> 才會執行。與既有合併範圍重疊時，舊的合併會被取消（以警告回報）。
    /// </summary>
    public MergeResult Merge(string workbookId, string sheet, string range, bool discardOtherValues = false)
    {
        var target = A1Address.ParseRange(range);
        if (target.Kind != RangeKind.Cells)
        {
            throw new OfficeToolException(ErrorCodes.InvalidRange, $"不可以合併整欄或整列：{range}", "請指定有限的矩形範圍，例如 A1:C1");
        }

        if (target.CellCount < 2)
        {
            throw Invalid($"合併至少需要兩個儲存格：{target}", "請指定多格的範圍，例如 A1:C1");
        }

        if (target.CellCount > options.MaxCellsPerFormat)
        {
            throw new OfficeToolException(ErrorCodes.InvalidRange, $"範圍 {target} 有 {target.CellCount} 格，超過上限 {options.MaxCellsPerFormat}", "請縮小範圍");
        }

        return sessions.Use(workbookId, mutates: true, s =>
        {
            var ws = SheetLookup.Find(s.Workbook, sheet);

            if (ws.Tables.Any(t => Intersects(ToAddress(t.RangeAddress), target)))
            {
                throw Invalid($"範圍 {target} 與表格重疊", "Excel 不允許合併表格內的儲存格；請先把範圍移到表格之外");
            }

            var existing = ws.MergedRanges.Where(m => Intersects(ToAddress(m.RangeAddress), target)).ToList();
            if (existing.Count == 1 && ToAddress(existing[0].RangeAddress) == AsCells(target))
            {
                return new MergeResult(ws.Name, [target.ToString()], []); // 已經是這個合併範圍
            }

            var block = ws.Range(target.FirstRow, target.FirstColumn, target.LastRow, target.LastColumn);
            var others = block.CellsUsed(XLCellsUsedOptions.Contents)
                .Where(c => c.Address.RowNumber != target.FirstRow || c.Address.ColumnNumber != target.FirstColumn)
                .Select(c => c.Address.ToStringRelative(false))
                .ToList();
            if (others.Count > 0 && !discardOtherValues)
            {
                var shown = string.Join("、", others.Take(5));
                var more = others.Count > 5 ? $" 等共 {others.Count} 格" : string.Empty;
                throw Invalid(
                    $"合併 {target} 會清掉左上角以外的 {others.Count} 個儲存格的內容：{shown}{more}",
                    "內容只會保留在左上角。確定要放棄其他內容的話，請帶 discardOtherValues = true");
            }

            var warnings = new List<string>();
            if (others.Count > 0)
            {
                warnings.Add($"已清除 {others.Count} 個儲存格的內容（只保留左上角 {new CellAddress(target.FirstRow, target.FirstColumn)}）");
            }

            if (existing.Count > 0)
            {
                warnings.Add($"原有的合併範圍 {string.Join("、", existing.Select(m => m.RangeAddress.ToStringRelative(false)))} 已被取消並併入新的合併範圍");
            }

            block.Merge();
            return new MergeResult(ws.Name, [target.ToString()], warnings);
        });
    }

    /// <summary>取消與 <paramref name="range"/> 有交集的所有合併（只要範圍碰到合併區的任何一格就會取消整個合併）。沒有合併時回傳空清單。</summary>
    public MergeResult Unmerge(string workbookId, string sheet, string range)
    {
        var target = A1Address.ParseRange(range);
        return sessions.Use(workbookId, mutates: true, s =>
        {
            var ws = SheetLookup.Find(s.Workbook, sheet);
            var hits = ws.MergedRanges.Where(m => Intersects(ToAddress(m.RangeAddress), target)).ToList();
            var names = hits.Select(m => m.RangeAddress.ToStringRelative(false)).ToList();
            foreach (var merged in hits)
            {
                merged.Unmerge(); // 直接對合併範圍本身操作；對較大或只涵蓋一部分的範圍呼叫 Unmerge 在 ClosedXML 不會生效
            }

            return new MergeResult(ws.Name, names, []);
        });
    }

    /// <summary>凍結窗格：上方 <paramref name="rows"/> 列與左邊 <paramref name="columns"/> 欄；兩者都是 0 代表取消凍結。</summary>
    public FreezeResult FreezePanes(string workbookId, string sheet, int rows, int columns)
    {
        if (rows < 0 || rows > MaxFrozenRows || columns < 0 || columns > MaxFrozenColumns)
        {
            throw Invalid($"無效的凍結範圍：{rows} 列、{columns} 欄", $"列數要介於 0 到 {MaxFrozenRows}、欄數 0 到 {MaxFrozenColumns}；兩者都是 0 代表取消凍結");
        }

        return sessions.Use(workbookId, mutates: true, s =>
        {
            var ws = SheetLookup.Find(s.Workbook, sheet);
            ws.SheetView.Freeze(rows, columns); // 要用 Freeze(r, c)：FreezeRows / FreezeColumns 只改一個方向，另一個方向維持舊值
            return new FreezeResult(ws.Name, ws.SheetView.SplitRow, ws.SheetView.SplitColumn);
        });
    }

    private static void Apply(IXLStyle style, FormatRules.Parsed f)
    {
        if (f.Bold is { } bold)
        {
            style.Font.Bold = bold;
        }

        if (f.Italic is { } italic)
        {
            style.Font.Italic = italic;
        }

        if (f.Underline is { } underline)
        {
            style.Font.Underline = underline ? XLFontUnderlineValues.Single : XLFontUnderlineValues.None;
        }

        if (f.Strikethrough is { } strike)
        {
            style.Font.Strikethrough = strike;
        }

        if (f.FontName is not null)
        {
            style.Font.FontName = f.FontName;
        }

        if (f.FontSize is { } size)
        {
            style.Font.FontSize = size;
        }

        if (f.FontColor is not null)
        {
            style.Font.FontColor = f.FontColor;
        }

        if (f.ClearFill)
        {
            style.Fill.PatternType = XLFillPatternValues.None;
        }
        else if (f.FillColor is not null)
        {
            style.Fill.BackgroundColor = f.FillColor;
        }

        if (f.Horizontal is { } horizontal)
        {
            style.Alignment.Horizontal = horizontal;
        }

        if (f.Vertical is { } vertical)
        {
            style.Alignment.Vertical = vertical;
        }

        if (f.WrapText is { } wrap)
        {
            style.Alignment.WrapText = wrap;
        }

        if (f.NumberFormat is not null)
        {
            style.NumberFormat.Format = f.NumberFormat;
        }

        if (f.BorderStyle is { } border)
        {
            if (f.OutlineOnly)
            {
                style.Border.OutsideBorder = border;
                if (f.BorderColor is not null)
                {
                    style.Border.OutsideBorderColor = f.BorderColor;
                }
            }
            else
            {
                style.Border.TopBorder = border;
                style.Border.BottomBorder = border;
                style.Border.LeftBorder = border;
                style.Border.RightBorder = border;
                if (f.BorderColor is not null)
                {
                    style.Border.TopBorderColor = f.BorderColor;
                    style.Border.BottomBorderColor = f.BorderColor;
                    style.Border.LeftBorderColor = f.BorderColor;
                    style.Border.RightBorderColor = f.BorderColor;
                }
            }
        }
    }

    private readonly record struct ColumnSpan(int First, int Last);

    private static ColumnSpan ParseColumns(string columns)
    {
        var text = columns?.Trim() ?? string.Empty;
        if (!text.Contains(':', StringComparison.Ordinal))
        {
            text = $"{text}:{text}";
        }

        if (!A1Address.TryParseRange(text, out var range) || range.Kind != RangeKind.WholeColumns)
        {
            throw new OfficeToolException(ErrorCodes.InvalidRange, $"無效的欄「{columns}」", "請用欄名稱，例如 A，或欄範圍，例如 A:C");
        }

        return new ColumnSpan(range.FirstColumn, range.LastColumn);
    }

    private static List<ColumnWidth> Widths(int first, int last, Func<int, double> width) =>
        Enumerable.Range(first, Math.Min(last - first + 1, MaxListedColumns))
            .Select(c => new ColumnWidth(A1Address.ColumnName(c), width(c)))
            .ToList();

    private static string DisplayText(IXLCell cell)
    {
        try
        {
            _ = cell.Value; // 公式要先計算，GetFormattedString 才會用到計算結果
            return cell.GetFormattedString(CultureInfo.InvariantCulture);
        }
        catch (Exception ex) when (ex is not (OutOfMemoryException or OfficeToolException))
        {
            return string.Empty; // 公式算不出來：不納入估算
        }
    }

    private static RangeAddress ToAddress(IXLRangeAddress a) =>
        new(a.FirstAddress.RowNumber, a.FirstAddress.ColumnNumber, a.LastAddress.RowNumber, a.LastAddress.ColumnNumber);

    private static RangeAddress AsCells(RangeAddress a) => a with { Kind = RangeKind.Cells };

    private static bool Intersects(RangeAddress a, RangeAddress b) =>
        a.FirstRow <= b.LastRow && b.FirstRow <= a.LastRow && a.FirstColumn <= b.LastColumn && b.FirstColumn <= a.LastColumn;

    private static OfficeToolException Invalid(string message, string hint) => new(ErrorCodes.InvalidValue, message, hint);
}
