using ClosedXML.Excel;
using ExcelTools.Core.Internal;
using ExcelTools.Core.Models;
using ExcelTools.Core.Workspace;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Operations;

/// <summary>工作表層級的操作。位置一律是 1 起算的分頁順序。</summary>
public sealed class SheetOperations(WorkbookSessionManager sessions)
{
    public const int MaxMergedRangesListed = 50;

    private const int MaxSheetNameLength = 31;

    public IReadOnlyList<SheetSummary> ListSheets(string workbookId) =>
        sessions.Use(workbookId, mutates: false, s => SheetNameRules.Summarize(s.Workbook));

    public SheetInfo GetSheetInfo(string workbookId, string sheet) =>
        sessions.Use(workbookId, mutates: false, s =>
        {
            var ws = SheetLookup.Find(s.Workbook, sheet);
            var merged = ws.MergedRanges.Select(r => r.RangeAddress.ToStringRelative(false)).ToList();
            return new SheetInfo(
                ws.Name,
                ws.Position,
                ws.Visibility.ToString(),
                ws.RangeUsed(XLCellsUsedOptions.Contents)?.RangeAddress.ToStringRelative(false),
                merged.Take(MaxMergedRangesListed).ToList(),
                merged.Count,
                ws.SheetView.SplitRow,
                ws.SheetView.SplitColumn,
                ws.Tables.Select(t => t.Name).ToList(),
                ws.AutoFilter.IsEnabled);
        });

    /// <param name="position">插入位置（1 起算）；省略則加在最後。</param>
    public SheetChangeResult Add(string workbookId, string name, int? position = null) =>
        sessions.Use(workbookId, mutates: true, s =>
        {
            var wb = s.Workbook;
            var clean = RequireFreeName(wb, name);
            if (position is not null)
            {
                RequirePosition(position.Value, wb.Worksheets.Count + 1);
            }

            var ws = wb.AddWorksheet(clean);
            if (position is not null)
            {
                ws.Position = position.Value;
            }

            return Result(wb);
        });

    public SheetChangeResult Rename(string workbookId, string sheet, string newName) =>
        sessions.Use(workbookId, mutates: true, s =>
        {
            var wb = s.Workbook;
            var ws = SheetLookup.Find(wb, sheet);
            var clean = RequireFreeName(wb, newName, except: ws);

            ws.Name = clean; // ClosedXML 會同步更新其他工作表公式裡的參照
            return Result(wb);
        });

    /// <param name="newName">省略則自動命名為「原名 (2)」。</param>
    /// <param name="position">新工作表的位置；省略則緊接在來源之後（與 Excel 一致）。</param>
    public SheetChangeResult Copy(string workbookId, string sheet, string? newName = null, int? position = null) =>
        sessions.Use(workbookId, mutates: true, s =>
        {
            var wb = s.Workbook;
            var source = SheetLookup.Find(wb, sheet);
            var name = newName is null ? NextCopyName(wb, source.Name) : RequireFreeName(wb, newName);
            if (position is not null)
            {
                RequirePosition(position.Value, wb.Worksheets.Count + 1);
            }

            var copy = source.CopyTo(name);
            copy.Position = position ?? source.Position + 1;
            return Result(wb);
        });

    public SheetChangeResult Delete(string workbookId, string sheet) =>
        sessions.Use(workbookId, mutates: true, s =>
        {
            var wb = s.Workbook;
            var ws = SheetLookup.Find(wb, sheet);

            if (wb.Worksheets.Count == 1)
            {
                throw Invalid("活頁簿至少要保留一張工作表", "請先新增另一張工作表再刪除");
            }

            if (ws.Visibility == XLWorksheetVisibility.Visible && wb.Worksheets.Count(w => w.Visibility == XLWorksheetVisibility.Visible) == 1)
            {
                throw Invalid("活頁簿至少要保留一張可見的工作表（其餘是隱藏的）", "請先新增一張工作表再刪除");
            }

            var (locations, total) = SheetLookup.FindReferences(wb, ws);
            var deletedName = ws.Name;
            ws.Delete();

            var warnings = new List<string>();
            if (total > 0)
            {
                var shown = string.Join("、", locations);
                var more = total > locations.Count ? $" 等共 {total} 處" : string.Empty;
                warnings.Add($"有 {total} 處公式或名稱引用了已刪除的工作表「{deletedName}」，Excel 開啟後會顯示 #REF!：{shown}{more}");
            }

            return Result(wb, warnings);
        });

    public SheetChangeResult Move(string workbookId, string sheet, int position) =>
        sessions.Use(workbookId, mutates: true, s =>
        {
            var wb = s.Workbook;
            var ws = SheetLookup.Find(wb, sheet);
            RequirePosition(position, wb.Worksheets.Count);

            ws.Position = position;
            return Result(wb);
        });

    private static string RequireFreeName(IXLWorkbook wb, string name, IXLWorksheet? except = null)
    {
        var clean = SheetNameRules.Validate(name);
        if (SheetLookup.NameTaken(wb, clean, except))
        {
            throw Invalid($"已經有名為「{clean}」的工作表（工作表名稱不分大小寫）", $"現有的工作表：{string.Join("、", SheetLookup.Names(wb))}");
        }

        return clean;
    }

    private static void RequirePosition(int position, int max)
    {
        if (position < 1 || position > max)
        {
            throw Invalid($"位置 {position} 超出範圍", $"位置是 1 起算的分頁順序，這次可用的範圍是 1 到 {max}");
        }
    }

    private static string NextCopyName(IXLWorkbook wb, string baseName)
    {
        for (var n = 2; ; n++)
        {
            var suffix = $" ({n})";
            var stem = baseName.Length + suffix.Length > MaxSheetNameLength ? baseName[..(MaxSheetNameLength - suffix.Length)] : baseName;
            var candidate = stem + suffix;
            if (!SheetLookup.NameTaken(wb, candidate))
            {
                return candidate;
            }
        }
    }

    private static SheetChangeResult Result(IXLWorkbook wb, IReadOnlyList<string>? warnings = null) =>
        new(SheetNameRules.Summarize(wb), warnings ?? []);

    private static OfficeToolException Invalid(string message, string hint) => new(ErrorCodes.InvalidValue, message, hint);
}
