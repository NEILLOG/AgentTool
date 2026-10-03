using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace ExcelTools.Core.Internal;

/// <summary>
/// 從公式文字找出對某張工作表的 A1 參照（儲存格、範圍、整欄、整列）。
/// 用於刪除列欄時警告失效的參照；是啟發式的（不是完整的公式解析器），寧可多報不可漏報。
/// </summary>
internal static partial class FormulaReferences
{
    private const int MaxLocations = 10;

    // [工作表!]（A1 | A1:B2 | A:B | 1:3）；前面不能緊接字母數字或點（避免取到函式名或名稱的一部分），後面不能接字母數字、左括號或中括號（函式、表格名稱）
    [GeneratedRegex(
        @"(?<![\w.$])(?<sheet>(?:'(?:[^']|'')+'|[\p{L}_][\w.]*)!)?(?<ref>\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})(?![\w(\[])",
        RegexOptions.CultureInvariant,
        matchTimeoutMilliseconds: 1000)]
    private static partial Regex ReferencePattern();

    [GeneratedRegex("\"(?:[^\"]|\"\")*\"", RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 1000)]
    private static partial Regex StringLiteral();

    /// <param name="formulaSheet">公式所在的工作表；沒有前綴的參照視為指向它。已定義名稱傳 null（名稱裡的參照一定有前綴）。</param>
    public static IEnumerable<RangeAddress> Find(string formula, string? formulaSheet, string targetSheet)
    {
        var text = StringLiteral().Replace(formula, "\"\"");
        foreach (Match match in ReferencePattern().Matches(text))
        {
            var sheet = match.Groups["sheet"].Success ? Unquote(match.Groups["sheet"].Value) : formulaSheet;
            if (sheet is null || !string.Equals(sheet, targetSheet, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (A1Address.TryParseRange(match.Groups["ref"].Value, out var range))
            {
                yield return range;
            }
        }
    }

    /// <summary>
    /// 找出「整個參照都落在被刪除的列（或欄）內」的公式與已定義名稱——這種參照 Excel 會變成 #REF!，
    /// 但 ClosedXML 不會（刪列時會悄悄改指別的儲存格，刪欄時完全不更新）。位於被刪區域內的公式本身會被刪掉，不計。
    /// </summary>
    public static (IReadOnlyList<string> Locations, int Total) FindDangling(IXLWorkbook workbook, IXLWorksheet target, bool rows, int first, int last)
    {
        var locations = new List<string>();
        var total = 0;

        void Record(string location)
        {
            total++;
            if (locations.Count < MaxLocations)
            {
                locations.Add(location);
            }
        }

        bool Dangling(RangeAddress r) => rows
            ? r.FirstRow >= first && r.LastRow <= last
            : r.FirstColumn >= first && r.LastColumn <= last;

        foreach (var ws in workbook.Worksheets.OrderBy(w => w.Position))
        {
            foreach (var cell in ws.CellsUsed(XLCellsUsedOptions.Contents).Where(c => c.HasFormula))
            {
                var insideDeleted = ReferenceEquals(ws, target)
                    && (rows ? cell.Address.RowNumber >= first && cell.Address.RowNumber <= last
                             : cell.Address.ColumnNumber >= first && cell.Address.ColumnNumber <= last);
                if (!insideDeleted && Find(cell.FormulaA1, ws.Name, target.Name).Any(Dangling))
                {
                    Record($"{ws.Name}!{cell.Address.ToStringRelative(false)}");
                }
            }
        }

        foreach (var name in workbook.DefinedNames)
        {
            if (Find(name.RefersTo ?? string.Empty, null, target.Name).Any(Dangling))
            {
                Record($"已定義名稱 {name.Name}");
            }
        }

        return (locations, total);
    }

    private static string Unquote(string sheetWithBang)
    {
        var name = sheetWithBang[..^1];
        return name.StartsWith('\'') ? name[1..^1].Replace("''", "'", StringComparison.Ordinal) : name;
    }
}
