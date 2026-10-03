namespace ExcelTools.Core.Models;

/// <param name="Range">實際回傳的範圍（已限制在工作表有內容的區域內）。</param>
/// <param name="UsedRange">工作表目前有內容的範圍，方便判斷資料在哪；空工作表為 null。</param>
/// <param name="Values">數字 / 字串 / 布林 / null；日期與時間是 ISO 字串；公式錯誤是 #DIV/0! 等文字。</param>
/// <param name="Formulas">只有 IncludeFormulas 時才有；公式格是 "=..."，其他是 null。</param>
/// <param name="Truncated">超過單次讀取上限而被截斷。</param>
/// <param name="NextRange">被截斷時，下一段要讀的範圍。</param>
/// <param name="CalculationWarnings">公式無法計算的格子（已改回傳快取值）。</param>
public sealed record RangeData(
    string Sheet,
    string Range,
    string? UsedRange,
    object?[][] Values,
    string?[][]? Formulas,
    bool Truncated,
    string? NextRange,
    IReadOnlyList<string> CalculationWarnings);

public sealed record ReadOptions(bool IncludeFormulas = false, bool UseFormattedText = false);

/// <param name="ParseFormulas">以 = 開頭的字串當公式。</param>
/// <param name="ParseIsoDates">ISO 日期 / 日期時間字串轉成日期。</param>
public sealed record WriteOptions(bool ParseFormulas = true, bool ParseIsoDates = true);

public enum ClearMode
{
    /// <summary>只清內容與公式，保留格式。</summary>
    Contents,

    /// <summary>只清格式，保留內容。</summary>
    Formats,

    /// <summary>內容與格式都清。</summary>
    All,
}

/// <param name="Range">實際寫入或清除的範圍。</param>
/// <param name="CellsAffected">範圍內的儲存格數。</param>
/// <param name="Warnings">寫入的公式算出錯誤值（#NAME?、#REF!、#DIV/0!）或無法計算時的提醒。</param>
public sealed record RangeChangeResult(string Sheet, string Range, int CellsAffected, IReadOnlyList<string> Warnings);

/// <param name="Range">受影響的帶狀範圍：插入 / 刪除的列（例如 5:6）或欄（例如 B:C）。</param>
/// <param name="UsedRange">操作後工作表有內容的範圍。</param>
/// <param name="Warnings">刪除時，被刪儲存格仍被公式或名稱參照的位置。</param>
public sealed record StructureChangeResult(string Sheet, string Range, string? UsedRange, IReadOnlyList<string> Warnings);

/// <param name="MatchCase">區分大小寫。</param>
/// <param name="WholeCell">整格內容必須完全相同，否則只要包含即可。</param>
/// <param name="SearchFormulas">也搜尋公式文字（例如 "=SUM("）。</param>
/// <param name="MaxResults">最多回傳幾筆。</param>
public sealed record FindOptions(bool MatchCase = false, bool WholeCell = false, bool SearchFormulas = false, int MaxResults = 100);

/// <param name="Value">儲存格的值（與 ReadRange 相同的表示法）。</param>
/// <param name="Formula">有公式時的公式文字（"=..."）。</param>
/// <param name="MatchedIn">命中的是 Text（顯示文字）、Value（原始值）還是 Formula。</param>
public sealed record FindMatch(string Sheet, string Address, object? Value, string? Formula, string MatchedIn);

public sealed record FindResult(IReadOnlyList<FindMatch> Matches, bool Truncated, IReadOnlyList<string> CalculationWarnings);

public enum CopyMode
{
    /// <summary>值、公式、格式與合併儲存格都複製；公式裡的相對參照會像 Excel 一樣跟著位移。</summary>
    All,

    /// <summary>只複製值（公式以計算結果貼上），不複製格式。</summary>
    ValuesOnly,
}
