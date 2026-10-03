using System.Text.RegularExpressions;
using ClosedXML.Excel;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Internal;

internal static class SheetLookup
{
    private const int MaxReferenceLocations = 10;

    /// <summary>以名稱找工作表（不分大小寫，與 Excel 一致）。找不到時列出現有的工作表。</summary>
    public static IXLWorksheet Find(IXLWorkbook workbook, string? name)
    {
        var trimmed = name?.Trim();
        if (!string.IsNullOrEmpty(trimmed) && workbook.TryGetWorksheet(trimmed, out var sheet))
        {
            return sheet;
        }

        throw new OfficeToolException(
            ErrorCodes.SheetNotFound,
            $"找不到工作表「{name}」",
            $"現有的工作表：{string.Join("、", Names(workbook))}");
    }

    public static IEnumerable<string> Names(IXLWorkbook workbook) =>
        workbook.Worksheets.OrderBy(w => w.Position).Select(w => w.Name);

    /// <summary>名稱是否已被「其他」工作表使用（不分大小寫）。</summary>
    public static bool NameTaken(IXLWorkbook workbook, string name, IXLWorksheet? except = null) =>
        workbook.Worksheets.Any(w => !ReferenceEquals(w, except) && string.Equals(w.Name, name, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// 找出引用指定工作表的公式與已定義名稱（不含該工作表自己）。ClosedXML 刪除工作表時不會處理這些參照，
    /// Excel 開啟後會變成 #REF!。回傳最多 <see cref="MaxReferenceLocations"/> 個位置與總數。
    /// </summary>
    public static (IReadOnlyList<string> Locations, int Total) FindReferences(IXLWorkbook workbook, IXLWorksheet target)
    {
        var escaped = Regex.Escape(target.Name);
        var quoted = Regex.Escape(target.Name.Replace("'", "''", StringComparison.Ordinal));
        var pattern = new Regex(
            $@"(?<![\w.']){escaped}!|'{quoted}'!",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
            TimeSpan.FromSeconds(1));

        var locations = new List<string>();
        var total = 0;

        void Record(string location)
        {
            total++;
            if (locations.Count < MaxReferenceLocations)
            {
                locations.Add(location);
            }
        }

        foreach (var ws in workbook.Worksheets.Where(w => !ReferenceEquals(w, target)).OrderBy(w => w.Position))
        {
            foreach (var cell in ws.CellsUsed(XLCellsUsedOptions.Contents).Where(c => c.HasFormula))
            {
                if (pattern.IsMatch(cell.FormulaA1))
                {
                    Record($"{ws.Name}!{cell.Address.ToStringRelative(false)}");
                }
            }
        }

        foreach (var name in workbook.DefinedNames)
        {
            if (pattern.IsMatch(name.RefersTo ?? string.Empty))
            {
                Record($"已定義名稱 {name.Name}");
            }
        }

        return (locations, total);
    }
}
