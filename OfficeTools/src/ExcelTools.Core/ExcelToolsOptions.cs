namespace ExcelTools.Core;

public sealed class ExcelToolsOptions
{
    public int MaxCellsPerRead { get; init; } = 2000;

    public TimeSpan IdleTimeout { get; init; } = TimeSpan.FromMinutes(30);
}
