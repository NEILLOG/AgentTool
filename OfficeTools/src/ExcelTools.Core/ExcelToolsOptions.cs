namespace ExcelTools.Core;

public sealed class ExcelToolsOptions
{
    public int MaxCellsPerRead { get; init; } = 2000;

    /// <summary>單次寫入的儲存格上限，防止 agent 一次丟入過大的資料。</summary>
    public int MaxCellsPerWrite { get; init; } = 20_000;

    public TimeSpan IdleTimeout { get; init; } = TimeSpan.FromMinutes(30);
}
