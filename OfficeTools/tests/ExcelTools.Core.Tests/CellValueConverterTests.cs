using System.Globalization;
using System.Text.Json;
using ClosedXML.Excel;
using ExcelTools.Core.Internal;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Tests;

public class CellValueConverterTests
{
    private static CellInput Input(object? v, CellWriteOptions? o = null) => CellValueConverter.ToCellInput(v, o ?? new CellWriteOptions());

    private static OfficeToolException InvalidValue(object? v) =>
        Assert.Throws<OfficeToolException>(() => Input(v));

    /// <summary>寫入 → 存成 xlsx → 重新載入 → 讀回可序列化值。</summary>
    private static object? RoundTrip(object? value, CellWriteOptions? options = null)
    {
        using var wb = new XLWorkbook();
        var cell = wb.AddWorksheet("S").Cell("A1");
        CellValueConverter.Apply(cell, Input(value, options));

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        return CellValueConverter.ToSerializable(reloaded.Worksheet("S").Cell("A1").Value);
    }

    // ---- 讀取 ----

    [Fact]
    public void Reads_each_cell_type()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("A1").Value = 3.5;
        ws.Cell("A2").Value = "文字";
        ws.Cell("A3").Value = true;
        ws.Cell("A4").Value = new DateTime(2026, 10, 3);
        ws.Cell("A5").Value = new DateTime(2026, 10, 3, 14, 5, 9);
        ws.Cell("A6").Value = TimeSpan.FromMinutes(90);
        ws.Cell("A7").FormulaA1 = "1/0";

        Assert.Equal(3.5, CellValueConverter.ToSerializable(ws.Cell("A1").Value));
        Assert.Equal("文字", CellValueConverter.ToSerializable(ws.Cell("A2").Value));
        Assert.Equal(true, CellValueConverter.ToSerializable(ws.Cell("A3").Value));
        Assert.Equal("2026-10-03", CellValueConverter.ToSerializable(ws.Cell("A4").Value));
        Assert.Equal("2026-10-03T14:05:09", CellValueConverter.ToSerializable(ws.Cell("A5").Value));
        Assert.Equal("01:30:00", CellValueConverter.ToSerializable(ws.Cell("A6").Value));
        Assert.Equal("#DIV/0!", CellValueConverter.ToSerializable(ws.Cell("A7").Value));
        Assert.Null(CellValueConverter.ToSerializable(ws.Cell("A8").Value));
    }

    [Theory]
    [InlineData(XLError.DivisionByZero, "#DIV/0!")]
    [InlineData(XLError.NameNotRecognized, "#NAME?")]
    [InlineData(XLError.NoValueAvailable, "#N/A")]
    [InlineData(XLError.NullValue, "#NULL!")]
    [InlineData(XLError.NumberInvalid, "#NUM!")]
    [InlineData(XLError.CellReference, "#REF!")]
    [InlineData(XLError.IncompatibleValue, "#VALUE!")]
    public void Error_values_use_excel_error_text(XLError error, string expected) =>
        Assert.Equal(expected, CellValueConverter.ToSerializable(error));

    [Fact]
    public void Dates_format_without_timezone_and_trim_fractional_zeros()
    {
        Assert.Equal("2026-10-03", CellValueConverter.FormatDateTime(new DateTime(2026, 10, 3)));
        Assert.Equal("2026-10-03T14:05:09", CellValueConverter.FormatDateTime(new DateTime(2026, 10, 3, 14, 5, 9)));
        Assert.Equal("2026-10-03T14:05:09.25", CellValueConverter.FormatDateTime(new DateTime(2026, 10, 3, 14, 5, 9, 250)));
    }

    [Fact]
    public void Formatting_ignores_the_current_culture()
    {
        var old = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("ar-SA"); // 非西曆、非拉丁數字
            Assert.Equal("2026-10-03T14:05:09", CellValueConverter.FormatDateTime(new DateTime(2026, 10, 3, 14, 5, 9)));
            Assert.Equal(3, Input(3.0).Value.GetNumber());
        }
        finally
        {
            CultureInfo.CurrentCulture = old;
        }
    }

    // ---- 寫入的型別判斷 ----

    [Fact]
    public void Default_options_parse_both_formulas_and_dates()
    {
        var options = new CellWriteOptions();
        Assert.True(options.ParseFormulas);
        Assert.True(options.ParseIsoDates);
    }

    [Fact]
    public void Null_clears() => Assert.Equal(CellInputKind.Clear, Input(null).Kind);

    [Theory]
    [InlineData((sbyte)-1)]
    [InlineData((byte)1)]
    [InlineData((short)2)]
    [InlineData((ushort)3)]
    [InlineData(4)]
    [InlineData(5u)]
    [InlineData(6L)]
    [InlineData(7UL)]
    [InlineData(8.5f)]
    [InlineData(9.25)]
    public void Numeric_types_become_numbers(object n)
    {
        var input = Input(n);
        Assert.Equal(CellInputKind.Value, input.Kind);
        Assert.Equal(XLDataType.Number, input.Value.Type);
        Assert.Equal(Convert.ToDouble(n, CultureInfo.InvariantCulture), input.Value.GetNumber());
    }

    [Fact]
    public void Decimal_becomes_number() => Assert.Equal(12.34, Input(12.34m).Value.GetNumber());

    [Fact]
    public void Bool_becomes_boolean() => Assert.True(Input(true).Value.GetBoolean());

    [Fact]
    public void Plain_string_stays_text()
    {
        var input = Input("00123");
        Assert.Equal(XLDataType.Text, input.Value.Type);
        Assert.Equal("00123", input.Value.GetText());
    }

    [Theory]
    [InlineData("123")]
    [InlineData("1e5")]
    [InlineData("TRUE")]
    [InlineData("12:30")]
    [InlineData("  ")]
    [InlineData("")]
    public void Strings_that_look_like_other_types_stay_text(string s) =>
        Assert.Equal(XLDataType.Text, Input(s).Value.Type);

    [Fact]
    public void Equals_prefix_is_a_formula_without_the_equals_sign()
    {
        var input = Input("=SUM(A1:A3)");
        Assert.Equal(CellInputKind.Formula, input.Kind);
        Assert.Equal("SUM(A1:A3)", input.Formula);
    }

    [Theory]
    [InlineData("=")]
    [InlineData("=   ")]
    public void Lone_equals_is_text(string s) => Assert.Equal(CellInputKind.Value, Input(s).Kind);

    [Fact]
    public void Formula_parsing_can_be_disabled()
    {
        var input = Input("=SUM(A1:A3)", new CellWriteOptions(ParseFormulas: false));
        Assert.Equal(CellInputKind.Value, input.Kind);
        Assert.Equal("=SUM(A1:A3)", input.Value.GetText());
    }

    [Theory]
    [InlineData("2026-10-03", 2026, 10, 3, 0, 0, 0)]
    [InlineData("2026-10-03T14:05", 2026, 10, 3, 14, 5, 0)]
    [InlineData("2026-10-03T14:05:09", 2026, 10, 3, 14, 5, 9)]
    public void Iso_date_strings_become_dates(string s, int y, int mo, int d, int h, int mi, int sec)
    {
        var input = Input(s);
        Assert.Equal(XLDataType.DateTime, input.Value.Type);
        Assert.Equal(new DateTime(y, mo, d, h, mi, sec), input.Value.GetDateTime());
    }

    [Fact]
    public void Iso_date_with_fraction_is_a_date() =>
        Assert.Equal(XLDataType.DateTime, Input("2026-10-03T14:05:09.250").Value.Type);

    [Theory]
    [InlineData("2026-13-45")]
    [InlineData("2026-02-30")]
    [InlineData("2026/10/03")]
    [InlineData("10/03/2026")]
    [InlineData("2026-10-03T14:05:09Z")]
    [InlineData("2026-10-03T14:05:09+08:00")]
    [InlineData("2026-10-03 14:05:09")]
    [InlineData("1899-12-31")]
    [InlineData("0001-01-01")]
    public void Non_iso_or_invalid_or_out_of_range_dates_stay_text(string s) =>
        Assert.Equal(XLDataType.Text, Input(s).Value.Type);

    [Fact]
    public void Iso_date_parsing_can_be_disabled()
    {
        var input = Input("2026-10-03", new CellWriteOptions(ParseIsoDates: false));
        Assert.Equal(XLDataType.Text, input.Value.Type);
    }

    [Fact]
    public void Dotnet_date_types_become_dates()
    {
        Assert.Equal(new DateTime(2026, 10, 3, 1, 2, 3), Input(new DateTime(2026, 10, 3, 1, 2, 3)).Value.GetDateTime());
        Assert.Equal(new DateTime(2026, 10, 3), Input(new DateOnly(2026, 10, 3)).Value.GetDateTime());
        Assert.Equal(new DateTime(2026, 10, 3, 1, 2, 3), Input(new DateTimeOffset(2026, 10, 3, 1, 2, 3, TimeSpan.FromHours(8))).Value.GetDateTime());
        Assert.Equal(TimeSpan.FromMinutes(90), Input(new TimeOnly(1, 30)).Value.GetTimeSpan());
        Assert.Equal(TimeSpan.FromMinutes(90), Input(TimeSpan.FromMinutes(90)).Value.GetTimeSpan());
    }

    [Fact]
    public void Json_elements_are_converted_by_kind()
    {
        using var doc = JsonDocument.Parse("""{"n":1.5,"s":"=A1+1","t":true,"f":false,"z":null,"d":"2026-10-03"}""");
        var root = doc.RootElement;
        Assert.Equal(1.5, Input(root.GetProperty("n")).Value.GetNumber());
        Assert.Equal("A1+1", Input(root.GetProperty("s")).Formula);
        Assert.True(Input(root.GetProperty("t")).Value.GetBoolean());
        Assert.False(Input(root.GetProperty("f")).Value.GetBoolean());
        Assert.Equal(CellInputKind.Clear, Input(root.GetProperty("z")).Kind);
        Assert.Equal(XLDataType.DateTime, Input(root.GetProperty("d")).Value.Type);
    }

    // ---- 無效值 ----

    [Fact]
    public void Invalid_values_return_INVALID_VALUE()
    {
        using var doc = JsonDocument.Parse("""{"o":{},"a":[1]}""");
        foreach (var bad in new object[]
                 {
                     double.NaN, double.PositiveInfinity, float.NegativeInfinity,
                     new string('x', CellValueConverter.MaxTextLength + 1),
                     new DateTime(1899, 12, 31), new DateOnly(1800, 1, 1),
                     Guid.NewGuid(), new object(), new[] { 1, 2 },
                     doc.RootElement.GetProperty("o"), doc.RootElement.GetProperty("a"),
                 })
        {
            var ex = InvalidValue(bad);
            Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
            Assert.NotNull(ex.Hint);
        }
    }

    [Fact]
    public void Text_at_the_excel_length_limit_is_accepted() =>
        Assert.Equal(CellValueConverter.MaxTextLength, RoundTrip(new string('x', CellValueConverter.MaxTextLength))!.ToString()!.Length);

    // ---- 來回一致（寫入 → 存檔 → 重新載入 → 讀回） ----

    [Theory]
    [InlineData("00123")]
    [InlineData("  前後有空白  ")]
    [InlineData("第一行\n第二行")]
    [InlineData("中文、emoji 😀、ñ")]
    [InlineData("it's \"quoted\" & <tagged>")]
    [InlineData("TRUE")]
    [InlineData("-")]
    public void Text_round_trips_exactly(string s) => Assert.Equal(s, RoundTrip(s));

    [Theory]
    [InlineData(0.0)]
    [InlineData(-1.5)]
    [InlineData(3.14159265358979)]
    [InlineData(1234567890123.0)]
    [InlineData(1e-9)]
    [InlineData(1.5e300)]
    public void Numbers_round_trip_exactly(double n) => Assert.Equal(n, RoundTrip(n));

    [Fact]
    public void Integers_round_trip_as_numbers() => Assert.Equal(42.0, RoundTrip(42));

    [Fact]
    public void Booleans_round_trip()
    {
        Assert.Equal(true, RoundTrip(true));
        Assert.Equal(false, RoundTrip(false));
    }

    [Fact]
    public void Null_round_trips_as_null() => Assert.Null(RoundTrip(null));

    [Theory]
    [InlineData("2026-10-03")]
    [InlineData("2026-10-03T14:05:09")]
    [InlineData("1900-01-01")]
    [InlineData("2099-12-31T23:59:59")]
    public void Iso_dates_round_trip(string iso) => Assert.Equal(iso, RoundTrip(iso));

    [Fact]
    public void Date_with_minutes_only_round_trips_with_seconds()
    {
        Assert.Equal("2026-10-03T14:05:00", RoundTrip("2026-10-03T14:05"));
    }

    [Fact]
    public void Text_that_looks_like_a_date_round_trips_as_text_when_date_parsing_is_off() =>
        Assert.Equal("2026-10-03", RoundTrip("2026-10-03", new CellWriteOptions(ParseIsoDates: false)));

    [Fact]
    public void Formula_text_round_trips_as_text_when_formula_parsing_is_off() =>
        Assert.Equal("=1+1", RoundTrip("=1+1", new CellWriteOptions(ParseFormulas: false)));

    [Fact]
    public void Formula_is_stored_as_a_formula_and_keeps_its_value_after_reload()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        CellValueConverter.Apply(ws.Cell("A1"), Input(20));
        CellValueConverter.Apply(ws.Cell("A2"), Input("=A1*2"));
        Assert.Equal("A1*2", ws.Cell("A2").FormulaA1);
        Assert.Equal(40.0, CellValueConverter.ToSerializable(ws.Cell("A2").Value));

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        Assert.Equal("A1*2", reloaded.Worksheet("S").Cell("A2").FormulaA1);
        Assert.Equal(40.0, CellValueConverter.ToSerializable(reloaded.Worksheet("S").Cell("A2").Value));
    }

    // ---- Apply 的副作用 ----

    [Fact]
    public void Clear_removes_content_but_keeps_formatting()
    {
        using var wb = new XLWorkbook();
        var cell = wb.AddWorksheet("S").Cell("A1");
        cell.Value = "x";
        cell.Style.Font.Bold = true;

        CellValueConverter.Apply(cell, Input(null));

        Assert.True(cell.IsEmpty(XLCellsUsedOptions.Contents));
        Assert.True(cell.Style.Font.Bold);
    }

    [Fact]
    public void Writing_a_value_over_a_formula_removes_the_formula()
    {
        using var wb = new XLWorkbook();
        var cell = wb.AddWorksheet("S").Cell("A1");
        CellValueConverter.Apply(cell, Input("=1+1"));
        Assert.True(cell.HasFormula);

        CellValueConverter.Apply(cell, Input("plain"));

        Assert.False(cell.HasFormula);
        Assert.Equal("plain", cell.Value.GetText());
    }

    [Fact]
    public void Writing_a_formula_over_a_value_replaces_it()
    {
        using var wb = new XLWorkbook();
        var cell = wb.AddWorksheet("S").Cell("A1");
        cell.Value = "old";
        CellValueConverter.Apply(cell, Input("=2+3"));
        Assert.Equal(5.0, CellValueConverter.ToSerializable(cell.Value));
    }
}
