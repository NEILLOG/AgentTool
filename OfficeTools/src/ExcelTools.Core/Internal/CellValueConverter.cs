using System.Globalization;
using System.Text.Json;
using ClosedXML.Excel;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Internal;

internal enum CellInputKind
{
    Clear,
    Value,
    Formula,
}

/// <summary>要寫進儲存格的內容，已完成型別判斷，與工作表無關，可單獨測試。</summary>
internal readonly record struct CellInput(CellInputKind Kind, XLCellValue Value = default, string? Formula = null);

/// <param name="ParseFormulas">以 = 開頭的字串當公式。</param>
/// <param name="ParseIsoDates">ISO 日期 / 日期時間字串轉成日期。</param>
/// <remarks>刻意用 class：struct 的 <c>new()</c> 不會套用建構式參數的預設值（會得到全部 false）。</remarks>
internal sealed record CellWriteOptions(bool ParseFormulas = true, bool ParseIsoDates = true);

/// <summary>
/// 儲存格值與可序列化值（數字 / 字串 / 布林 / null）之間的轉換。
/// 日期與時間轉成不含時區的 ISO 字串；錯誤值轉成 Excel 的錯誤文字（#DIV/0! 等）。
/// </summary>
internal static class CellValueConverter
{
    public const int MaxTextLength = 32_767; // Excel 單一儲存格字元上限

    private static readonly DateTime MinExcelDate = new(1900, 1, 1);

    private static readonly string[] IsoFormats =
    [
        "yyyy-MM-dd",
        "yyyy-MM-dd'T'HH:mm",
        "yyyy-MM-dd'T'HH:mm:ss",
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF",
    ];

    // ---- 讀取 ----

    public static object? ToSerializable(XLCellValue value) => value.Type switch
    {
        XLDataType.Blank => null,
        XLDataType.Boolean => value.GetBoolean(),
        XLDataType.Number => value.GetNumber(),
        XLDataType.Text => value.GetText(),
        XLDataType.DateTime => FormatDateTime(value.GetDateTime()),
        XLDataType.TimeSpan => value.GetTimeSpan().ToString("c", CultureInfo.InvariantCulture),
        XLDataType.Error => ErrorText(value.GetError()),
        _ => value.ToString(CultureInfo.InvariantCulture),
    };

    public static string FormatDateTime(DateTime value)
    {
        if (value.TimeOfDay == TimeSpan.Zero)
        {
            return value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        return value.ToString("yyyy-MM-dd'T'HH:mm:ss.FFFFFFF", CultureInfo.InvariantCulture).TrimEnd('.');
    }

    private static string ErrorText(XLError error) => error switch
    {
        XLError.DivisionByZero => "#DIV/0!",
        XLError.NameNotRecognized => "#NAME?",
        XLError.NoValueAvailable => "#N/A",
        XLError.NullValue => "#NULL!",
        XLError.NumberInvalid => "#NUM!",
        XLError.CellReference => "#REF!",
        XLError.IncompatibleValue => "#VALUE!",
        _ => error.ToString(),
    };

    // ---- 寫入 ----

    public static CellInput ToCellInput(object? value, CellWriteOptions options) => value switch
    {
        null => new CellInput(CellInputKind.Clear),
        string s => FromString(s, options),
        char c => FromString(c.ToString(), options),
        bool b => Of(b),
        sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal => FromNumber(Convert.ToDouble(value, CultureInfo.InvariantCulture)),
        DateTime dt => FromDate(dt),
        DateTimeOffset dto => FromDate(dto.DateTime),
        DateOnly d => FromDate(d.ToDateTime(TimeOnly.MinValue)),
        TimeOnly t => Of(t.ToTimeSpan()),
        TimeSpan ts => Of(ts),
        JsonElement je => FromJson(je, options),
        _ => throw InvalidValue($"不支援的值型別 {value.GetType().Name}"),
    };

    public static void Apply(IXLCell cell, CellInput input)
    {
        switch (input.Kind)
        {
            case CellInputKind.Clear:
                cell.Clear(XLClearOptions.Contents); // 保留格式
                break;
            case CellInputKind.Formula:
                cell.FormulaA1 = input.Formula!;
                break;
            default:
                cell.Value = input.Value;
                break;
        }
    }

    private static CellInput Of(XLCellValue value) => new(CellInputKind.Value, value);

    private static CellInput FromNumber(double number) =>
        double.IsFinite(number) ? Of(number) : throw InvalidValue("數字不可為 NaN 或無限大");

    private static CellInput FromDate(DateTime date) =>
        date >= MinExcelDate ? Of(date) : throw InvalidValue($"日期不可早於 1900-01-01：{FormatDateTime(date)}");

    private static CellInput FromString(string s, CellWriteOptions options)
    {
        if (s.Length > MaxTextLength)
        {
            throw InvalidValue($"字串長度 {s.Length} 超過 Excel 單一儲存格上限 {MaxTextLength}");
        }

        if (options.ParseFormulas && s.StartsWith('=') && !string.IsNullOrWhiteSpace(s[1..]))
        {
            return new CellInput(CellInputKind.Formula, Formula: s[1..]);
        }

        if (options.ParseIsoDates && TryParseIsoDate(s, out var date))
        {
            return Of(date);
        }

        return Of(s);
    }

    private static CellInput FromJson(JsonElement e, CellWriteOptions options) => e.ValueKind switch
    {
        JsonValueKind.Null or JsonValueKind.Undefined => new CellInput(CellInputKind.Clear),
        JsonValueKind.True => Of(true),
        JsonValueKind.False => Of(false),
        JsonValueKind.Number => FromNumber(e.GetDouble()),
        JsonValueKind.String => FromString(e.GetString()!, options),
        _ => throw InvalidValue("儲存格值不可為物件或陣列"),
    };

    private static bool TryParseIsoDate(string s, out DateTime date) =>
        DateTime.TryParseExact(s, IsoFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out date) && date >= MinExcelDate;

    private static OfficeToolException InvalidValue(string reason) =>
        new(ErrorCodes.InvalidValue, $"無法寫入的儲存格值：{reason}", "儲存格值只能是數字、字串、布林、日期（ISO 字串）或 null");
}
