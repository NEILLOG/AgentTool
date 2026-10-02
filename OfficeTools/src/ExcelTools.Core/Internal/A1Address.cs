using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Internal;

internal enum RangeKind
{
    Cells,
    WholeColumns,
    WholeRows,
}

internal readonly record struct CellAddress(int Row, int Column)
{
    public override string ToString() => $"{A1Address.ColumnName(Column)}{Row}";
}

/// <summary>正規化後的範圍（First ≤ Last）。整欄 / 整列會展開成完整的列 / 欄範圍，並以 <see cref="Kind"/> 保留原本的形式。</summary>
internal readonly record struct RangeAddress(int FirstRow, int FirstColumn, int LastRow, int LastColumn, RangeKind Kind = RangeKind.Cells)
{
    public int RowCount => LastRow - FirstRow + 1;

    public int ColumnCount => LastColumn - FirstColumn + 1;

    public long CellCount => (long)RowCount * ColumnCount;

    public override string ToString() => Kind switch
    {
        RangeKind.WholeColumns => $"{A1Address.ColumnName(FirstColumn)}:{A1Address.ColumnName(LastColumn)}",
        RangeKind.WholeRows => $"{FirstRow}:{LastRow}",
        _ when FirstRow == LastRow && FirstColumn == LastColumn => new CellAddress(FirstRow, FirstColumn).ToString(),
        _ => $"{new CellAddress(FirstRow, FirstColumn)}:{new CellAddress(LastRow, LastColumn)}",
    };
}

/// <summary>A1 位址解析與格式化。支援 A1、A1:C10、整欄 A:A、整列 1:1，容許 $ 絕對符號、小寫與前後空白。</summary>
internal static class A1Address
{
    public const int MaxRow = 1_048_576;
    public const int MaxColumn = 16_384; // XFD

    private const int MaxColumnLetters = 3;
    private const int MaxRowDigits = 7;

    private const string FormatHint =
        "正確格式例如 A1、A1:C10、A:A（整欄）、1:1（整列）；工作表名稱請放在 sheet 參數，不要寫在範圍裡";

    public static string ColumnName(int column)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(column, 1);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(column, MaxColumn);

        Span<char> buffer = stackalloc char[MaxColumnLetters];
        var pos = buffer.Length;
        while (column > 0)
        {
            column--;
            buffer[--pos] = (char)('A' + (column % 26));
            column /= 26;
        }

        return new string(buffer[pos..]);
    }

    public static CellAddress ParseCell(string text)
    {
        var range = ParseRange(text);
        if (range.Kind != RangeKind.Cells || range.CellCount != 1)
        {
            throw Invalid(text, "這裡需要單一儲存格");
        }

        return new CellAddress(range.FirstRow, range.FirstColumn);
    }

    public static RangeAddress ParseRange(string text)
    {
        if (TryParseRange(text, out var range, out var error))
        {
            return range;
        }

        throw Invalid(text, error!);
    }

    public static bool TryParseRange(string? text, out RangeAddress range) => TryParseRange(text, out range, out _);

    private static bool TryParseRange(string? text, out RangeAddress range, out string? error)
    {
        range = default;
        error = null;

        var s = text?.Trim();
        if (string.IsNullOrEmpty(s))
        {
            error = "範圍是空的";
            return false;
        }

        if (s.Contains('!'))
        {
            error = "範圍不可包含工作表名稱";
            return false;
        }

        var colon = s.IndexOf(':', StringComparison.Ordinal);
        if (colon >= 0 && s.IndexOf(':', colon + 1) >= 0)
        {
            error = "只能有一個冒號";
            return false;
        }

        if (colon < 0)
        {
            if (!TryParsePart(s, out var column, out var row, out error))
            {
                return false;
            }

            if (column == 0 || row == 0)
            {
                error = "單一位址必須同時有欄與列";
                return false;
            }

            range = new RangeAddress(row, column, row, column);
            return true;
        }

        if (!TryParsePart(s.AsSpan(0, colon), out var c1, out var r1, out error)
            || !TryParsePart(s.AsSpan(colon + 1), out var c2, out var r2, out error))
        {
            return false;
        }

        var firstIsCell = c1 != 0 && r1 != 0;
        var secondIsCell = c2 != 0 && r2 != 0;

        if (firstIsCell && secondIsCell)
        {
            range = new RangeAddress(Math.Min(r1, r2), Math.Min(c1, c2), Math.Max(r1, r2), Math.Max(c1, c2));
            return true;
        }

        if (r1 == 0 && r2 == 0 && c1 != 0 && c2 != 0)
        {
            range = new RangeAddress(1, Math.Min(c1, c2), MaxRow, Math.Max(c1, c2), RangeKind.WholeColumns);
            return true;
        }

        if (c1 == 0 && c2 == 0 && r1 != 0 && r2 != 0)
        {
            range = new RangeAddress(Math.Min(r1, r2), 1, Math.Max(r1, r2), MaxColumn, RangeKind.WholeRows);
            return true;
        }

        error = "冒號兩側的格式必須一致（都是儲存格、都是欄，或都是列）";
        return false;
    }

    /// <summary>解析一側：[$]欄字母[$]列數字。欄或列缺少時分別回傳 0。</summary>
    private static bool TryParsePart(ReadOnlySpan<char> s, out int column, out int row, out string? error)
    {
        column = 0;
        row = 0;
        error = null;

        var i = 0;
        if (i < s.Length && s[i] == '$')
        {
            i++;
        }

        var lettersStart = i;
        while (i < s.Length && char.IsAsciiLetter(s[i]))
        {
            i++;
        }

        var letterCount = i - lettersStart;
        if (letterCount > MaxColumnLetters)
        {
            error = $"欄名稱最多 {MaxColumnLetters} 個字母（最大欄為 XFD）";
            return false;
        }

        for (var k = lettersStart; k < i; k++)
        {
            column = (column * 26) + (char.ToUpperInvariant(s[k]) - 'A' + 1);
        }

        if (column > MaxColumn)
        {
            error = $"欄超出範圍（最大欄為 XFD）";
            return false;
        }

        var rowDollar = false;
        if (letterCount > 0 && i < s.Length && s[i] == '$')
        {
            rowDollar = true;
            i++;
        }

        var digitsStart = i;
        while (i < s.Length && char.IsAsciiDigit(s[i]))
        {
            i++;
        }

        var digitCount = i - digitsStart;
        if (i != s.Length)
        {
            error = $"含有無效字元「{s[i]}」";
            return false;
        }

        if (letterCount == 0 && digitCount == 0 || rowDollar && digitCount == 0)
        {
            error = "位址不完整";
            return false;
        }

        if (digitCount > 0)
        {
            if (digitCount > MaxRowDigits || !int.TryParse(s[digitsStart..], out row) || row < 1 || row > MaxRow)
            {
                error = $"列超出範圍（有效範圍 1 到 {MaxRow}）";
                return false;
            }
        }

        return true;
    }

    private static OfficeToolException Invalid(string? text, string reason) =>
        new(ErrorCodes.InvalidRange, $"無效的範圍「{text}」：{reason}", FormatHint);
}
