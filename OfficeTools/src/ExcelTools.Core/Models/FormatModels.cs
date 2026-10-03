namespace ExcelTools.Core.Models;

/// <summary>
/// 要套用的格式。只會改你指定的屬性（null = 不動）；至少要指定一個。
/// 字串屬性不分大小寫，允許的值見各欄說明，無效的值會回 INVALID_VALUE 並列出可用的值。
/// </summary>
public sealed record FormatSpec
{
    public bool? Bold { get; init; }

    public bool? Italic { get; init; }

    public bool? Underline { get; init; }

    public bool? Strikethrough { get; init; }

    /// <summary>字型名稱，例如 "Arial"、"微軟正黑體"。</summary>
    public string? FontName { get; init; }

    /// <summary>字型大小，1 到 409。</summary>
    public double? FontSize { get; init; }

    /// <summary>"#RRGGBB"、"RRGGBB"、"#RGB" 或常用顏色名稱（red、blue…）。</summary>
    public string? FontColor { get; init; }

    /// <summary>背景色，格式同 FontColor；"none" 清除背景色。</summary>
    public string? FillColor { get; init; }

    /// <summary>general、left、center、right、justify、fill、centerContinuous、distributed。</summary>
    public string? HorizontalAlignment { get; init; }

    /// <summary>top、center、bottom、justify、distributed。</summary>
    public string? VerticalAlignment { get; init; }

    public bool? WrapText { get; init; }

    /// <summary>Excel 的數字格式碼，例如 "0.00"、"#,##0"、"0.0%"、"yyyy-mm-dd"、"@"（文字）、"General"。</summary>
    public string? NumberFormat { get; init; }

    /// <summary>none、thin、medium、thick、dashed、dotted、double、hair 等；none 清除邊框。</summary>
    public string? BorderStyle { get; init; }

    /// <summary>邊框顏色，格式同 FontColor；必須搭配 BorderStyle。</summary>
    public string? BorderColor { get; init; }

    /// <summary>all（每個儲存格四邊，預設）或 outline（只有範圍外框，不適用整欄整列）。</summary>
    public string? BorderSides { get; init; }
}

/// <param name="Range">受影響的範圍。</param>
/// <param name="CellsAffected">範圍的儲存格數；整欄整列是名義上的數量（不是實際建立的儲存格）。</param>
public sealed record FormatResult(string Sheet, string Range, long CellsAffected, IReadOnlyList<string> Warnings);

public sealed record ColumnWidth(string Column, double Width);

/// <param name="Columns">這次設定的欄與寬度（最多列出前 200 欄）。</param>
/// <param name="Unchanged">自動調整時沒有內容而維持原寬度的欄。</param>
/// <param name="Notes">例如寬度是估算值的說明。</param>
public sealed record ColumnWidthResult(string Sheet, IReadOnlyList<ColumnWidth> Columns, IReadOnlyList<string> Unchanged, IReadOnlyList<string> Notes);

/// <param name="Ranges">這次合併或取消合併的範圍。</param>
public sealed record MergeResult(string Sheet, IReadOnlyList<string> Ranges, IReadOnlyList<string> Warnings);

public sealed record FreezeResult(string Sheet, int FrozenRows, int FrozenColumns);
