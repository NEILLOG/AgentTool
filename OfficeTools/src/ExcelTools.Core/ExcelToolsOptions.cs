namespace ExcelTools.Core;

public sealed class ExcelToolsOptions
{
    public int MaxCellsPerRead { get; init; } = 2000;

    /// <summary>單次寫入的儲存格上限，防止 agent 一次丟入過大的資料。</summary>
    public int MaxCellsPerWrite { get; init; } = 20_000;

    /// <summary>格式化 / 合併矩形範圍時的儲存格上限；整欄、整列以欄 / 列樣式處理，不受此限。</summary>
    public int MaxCellsPerFormat { get; init; } = 50_000;

    public TimeSpan IdleTimeout { get; init; } = TimeSpan.FromMinutes(30);
}
