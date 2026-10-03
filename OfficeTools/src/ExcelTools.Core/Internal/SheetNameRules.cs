using ClosedXML.Excel;
using ExcelTools.Core.Models;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Internal;

internal static class SheetNameRules
{
    private const int MaxLength = 31;
    private static readonly char[] Forbidden = ['\\', '/', '?', '*', '[', ']', ':'];

    public static string Validate(string? name)
    {
        var trimmed = name?.Trim();
        if (string.IsNullOrEmpty(trimmed)
            || trimmed.Length > MaxLength
            || trimmed.IndexOfAny(Forbidden) >= 0
            || trimmed.StartsWith('\'')
            || trimmed.EndsWith('\''))
        {
            throw new OfficeToolException(
                ErrorCodes.InvalidValue,
                $"工作表名稱無效：「{name}」",
                $"名稱不可為空、最長 {MaxLength} 字元、不可含 \\ / ? * [ ] :，也不可以單引號開頭或結尾");
        }

        return trimmed;
    }

    public static IReadOnlyList<SheetSummary> Summarize(IXLWorkbook workbook) =>
        workbook.Worksheets
            .Select(ws => new SheetSummary(
                ws.Position,
                ws.Name,
                ws.RangeUsed(XLCellsUsedOptions.Contents)?.RangeAddress.ToStringRelative(false),
                ws.Visibility != XLWorksheetVisibility.Visible))
            .ToList();
}
