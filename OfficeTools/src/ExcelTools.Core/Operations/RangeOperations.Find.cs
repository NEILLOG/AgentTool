using System.Globalization;
using ClosedXML.Excel;
using ExcelTools.Core.Internal;
using ExcelTools.Core.Models;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Operations;

public sealed partial class RangeOperations
{
    /// <summary>
    /// 搜尋儲存格。比對三種文字：顯示文字（含格式，例如 "1,234.50"）、原始值（例如 "1234.5"、ISO 日期）、
    /// 以及選用的公式文字。結果依工作表分頁順序、列優先排列。
    /// </summary>
    /// <param name="sheet">省略則搜尋所有工作表。</param>
    /// <param name="range">限制範圍；指定時必須同時指定工作表。</param>
    public FindResult Find(string workbookId, string text, string? sheet = null, string? range = null, FindOptions? findOptions = null) =>
        sessions.Use(workbookId, mutates: false, s =>
        {
            var opts = findOptions ?? new FindOptions();
            if (string.IsNullOrEmpty(text))
            {
                throw new OfficeToolException(ErrorCodes.InvalidValue, "要搜尋的文字不可為空", "請提供要找的文字");
            }

            if (opts.MaxResults < 1)
            {
                throw new OfficeToolException(ErrorCodes.InvalidValue, "MaxResults 至少要是 1", "請提供正整數");
            }

            if (range is not null && sheet is null)
            {
                throw new OfficeToolException(ErrorCodes.InvalidValue, "指定範圍時必須同時指定工作表", "請加上 sheet 參數");
            }

            var sheets = sheet is null
                ? s.Workbook.Worksheets.OrderBy(w => w.Position).ToList()
                : [SheetLookup.Find(s.Workbook, sheet)];
            var scope = range is null ? (RangeAddress?)null : A1Address.ParseRange(range);

            var comparison = opts.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var matches = new List<FindMatch>();
            var warnings = new List<string>();
            var warningCount = 0;
            var truncated = false;

            foreach (var ws in sheets)
            {
                var bounds = ws.RangeUsed(XLCellsUsedOptions.Contents)?.RangeAddress;
                if (bounds is null)
                {
                    continue;
                }

                var region = ws.Range(
                    scope?.FirstRow ?? 1,
                    scope?.FirstColumn ?? 1,
                    Math.Min(scope?.LastRow ?? A1Address.MaxRow, bounds.LastAddress.RowNumber),
                    Math.Min(scope?.LastColumn ?? A1Address.MaxColumn, bounds.LastAddress.ColumnNumber));
                if (scope is { } sc && (sc.FirstRow > bounds.LastAddress.RowNumber || sc.FirstColumn > bounds.LastAddress.ColumnNumber))
                {
                    continue;
                }

                foreach (var cell in region.CellsUsed(XLCellsUsedOptions.Contents).OrderBy(c => c.Address.RowNumber).ThenBy(c => c.Address.ColumnNumber))
                {
                    var match = Match(cell, text, opts, comparison, ws.Name, warnings, ref warningCount);
                    if (match is null)
                    {
                        continue;
                    }

                    if (matches.Count >= opts.MaxResults)
                    {
                        truncated = true;
                        break;
                    }

                    matches.Add(match);
                }

                if (truncated)
                {
                    break;
                }
            }

            AddOverflowNote(warnings, warningCount);
            return new FindResult(matches, truncated, warnings);
        });

    private static FindMatch? Match(IXLCell cell, string text, FindOptions options, StringComparison comparison, string sheetName, List<string> warnings, ref int warningCount)
    {
        var value = ReadCell(cell, formatted: false, warnings, ref warningCount);
        string? display = null;
        try
        {
            display = cell.Value.IsBlank ? null : cell.GetFormattedString(CultureInfo.InvariantCulture);
        }
        catch (Exception ex) when (ex is not (OutOfMemoryException or OfficeToolException))
        {
            // 公式算不出來：ReadCell 已經回報過警告，這裡只比對原始值
        }

        var raw = value switch
        {
            null => null,
            double d => d.ToString("R", CultureInfo.InvariantCulture),
            bool b => b ? "TRUE" : "FALSE",
            _ => value.ToString(),
        };
        var formula = cell.HasFormula ? "=" + cell.FormulaA1 : null;

        bool Hit(string? candidate) =>
            candidate is not null && (options.WholeCell ? string.Equals(candidate, text, comparison) : candidate.Contains(text, comparison));

        var matchedIn =
            Hit(display) ? "Text"
            : Hit(raw) ? "Value"
            : options.SearchFormulas && Hit(formula) ? "Formula"
            : null;

        return matchedIn is null ? null : new FindMatch(sheetName, cell.Address.ToStringRelative(false), value, formula, matchedIn);
    }
}
