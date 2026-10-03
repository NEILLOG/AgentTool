using System.Globalization;
using ClosedXML.Excel;
using ExcelTools.Core.Internal;
using ExcelTools.Core.Models;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Tests;

public sealed class FormatOperationsTests : IDisposable
{
    private readonly ExcelEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    private string Open(Action<IXLWorksheet>? fill = null, ExcelEnv? env = null)
    {
        var e = env ?? _env;
        return e.Files.Open(e.MakeWorkbook($"{Guid.NewGuid():N}.xlsx", fill ?? (ws => { ws.Cell("A1").Value = "x"; }))).WorkbookId;
    }

    private T Style<T>(string id, Func<IXLWorksheet, T> read) => _env.Sessions.Use(id, false, s => read(s.Workbook.Worksheet("Data")));

    private bool IsDirty() => _env.Files.ListOpen().Single().IsDirty;

    // ---- FormatRange：套用與只改指定的屬性 ----

    [Fact]
    public void Applies_font_fill_alignment_and_wrap_to_a_rectangle()
    {
        var id = Open();

        var result = _env.Formats.FormatRange(id, "Data", "A1:C2", new FormatSpec
        {
            Bold = true,
            Italic = true,
            Underline = true,
            Strikethrough = true,
            FontName = "Arial",
            FontSize = 14,
            FontColor = "#FF0000",
            FillColor = "ffff00",
            HorizontalAlignment = "Center",
            VerticalAlignment = "top",
            WrapText = true,
        });

        Assert.Equal("Data", result.Sheet);
        Assert.Equal("A1:C2", result.Range);
        Assert.Equal(6, result.CellsAffected);
        Assert.Empty(result.Warnings);
        foreach (var address in new[] { "A1", "C2", "B1" }) // 包含原本是空的格子
        {
            var style = Style(id, ws => ws.Cell(address).Style);
            Assert.True(style.Font.Bold);
            Assert.True(style.Font.Italic);
            Assert.Equal(XLFontUnderlineValues.Single, style.Font.Underline);
            Assert.True(style.Font.Strikethrough);
            Assert.Equal("Arial", style.Font.FontName);
            Assert.Equal(14, style.Font.FontSize);
            Assert.Equal(XLColor.FromHtml("#FF0000"), style.Font.FontColor);
            Assert.Equal(XLColor.FromHtml("#FFFF00"), style.Fill.BackgroundColor);
            Assert.Equal(XLAlignmentHorizontalValues.Center, style.Alignment.Horizontal);
            Assert.Equal(XLAlignmentVerticalValues.Top, style.Alignment.Vertical);
            Assert.True(style.Alignment.WrapText);
        }

        Assert.True(IsDirty());
    }

    [Fact]
    public void Only_the_specified_properties_change()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "x";
            ws.Cell("A1").Style.Font.Italic = true;
            ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.Yellow;
            ws.Cell("A1").Style.NumberFormat.Format = "0.00";
        });

        _env.Formats.FormatRange(id, "Data", "A1", new FormatSpec { Bold = true });

        var style = Style(id, ws => ws.Cell("A1").Style);
        Assert.True(style.Font.Bold);
        Assert.True(style.Font.Italic);
        Assert.Equal(XLColor.Yellow, style.Fill.BackgroundColor);
        Assert.Equal("0.00", style.NumberFormat.Format);
    }

    [Fact]
    public void Properties_can_be_turned_off_again_and_fill_can_be_cleared()
    {
        var id = Open();
        _env.Formats.FormatRange(id, "Data", "A1", new FormatSpec { Bold = true, FillColor = "red", Underline = true });
        _env.Formats.FormatRange(id, "Data", "A1", new FormatSpec { Bold = false, FillColor = "none", Underline = false });

        var style = Style(id, ws => ws.Cell("A1").Style);
        Assert.False(style.Font.Bold);
        Assert.Equal(XLFillPatternValues.None, style.Fill.PatternType);
        Assert.Equal(XLFontUnderlineValues.None, style.Font.Underline);
    }

    [Fact]
    public void Formatting_never_changes_the_cell_contents()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "text";
            ws.Cell("A2").Value = 12;
            ws.Cell("A3").FormulaA1 = "A2*2";
        });

        _env.Formats.FormatRange(id, "Data", "A1:A3", new FormatSpec { Bold = true, NumberFormat = "0.00", HorizontalAlignment = "right" });

        Assert.Equal([["text"], [12.0], [24.0]], _env.Ranges.ReadRange(id, "Data", "A1:A3").Values);
        Assert.Equal("A2*2", Style(id, ws => ws.Cell("A3").FormulaA1));
    }

    [Fact]
    public void Number_formats_change_the_displayed_text()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = 1234.5;
            ws.Cell("A2").Value = 0.256;
            ws.Cell("A3").Value = new DateTime(2026, 10, 3);
            ws.Cell("A4").Value = 7;
        });

        _env.Formats.FormatRange(id, "Data", "A1", new FormatSpec { NumberFormat = "#,##0.00" });
        _env.Formats.FormatRange(id, "Data", "A2", new FormatSpec { NumberFormat = "0.0%" });
        _env.Formats.FormatRange(id, "Data", "A3", new FormatSpec { NumberFormat = "yyyy/mm/dd" });
        _env.Formats.FormatRange(id, "Data", "A4", new FormatSpec { NumberFormat = "000" });

        var shown = _env.Ranges.ReadRange(id, "Data", "A1:A4", new ReadOptions(UseFormattedText: true)).Values.Select(r => r[0]).ToArray();
        Assert.Equal(new object?[] { "1,234.50", "25.6%", "2026/10/03", "007" }, shown);
    }

    // ---- 範圍種類 ----

    [Fact]
    public void A_whole_column_is_formatted_for_existing_and_future_cells()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "x";
            ws.Cell("A5").Value = "y";
            ws.Cell("B1").Value = "other";
        });

        var result = _env.Formats.FormatRange(id, "Data", "A:A", new FormatSpec { Bold = true });

        Assert.Equal("A:A", result.Range);
        Assert.Equal(A1Address.MaxRow, result.CellsAffected);
        Assert.True(Style(id, ws => ws.Cell("A1").Style.Font.Bold));
        Assert.True(Style(id, ws => ws.Cell("A5").Style.Font.Bold));
        Assert.True(Style(id, ws => ws.Cell("A500").Style.Font.Bold)); // 還不存在的格子
        Assert.False(Style(id, ws => ws.Cell("B1").Style.Font.Bold));
        Assert.Equal("A1:B5", _env.Sheets.GetSheetInfo(id, "Data").UsedRange); // 不會把整欄建成儲存格
    }

    [Fact]
    public void Several_whole_columns_and_whole_rows_work()
    {
        var id = Open();

        _env.Formats.FormatRange(id, "Data", "B:C", new FormatSpec { FillColor = "red" });
        _env.Formats.FormatRange(id, "Data", "3:4", new FormatSpec { Italic = true });

        Assert.Equal(XLColor.Red, Style(id, ws => ws.Cell("C9").Style.Fill.BackgroundColor));
        Assert.True(Style(id, ws => ws.Cell("Z3").Style.Font.Italic));
        Assert.True(Style(id, ws => ws.Cell("A4").Style.Font.Italic));
        Assert.False(Style(id, ws => ws.Cell("A5").Style.Font.Italic));
    }

    [Fact]
    public void The_whole_sheet_can_be_formatted()
    {
        var id = Open();
        var result = _env.Formats.FormatRange(id, "Data", "A1:XFD1048576", new FormatSpec { FontName = "Arial" });

        Assert.Equal("Arial", Style(id, ws => ws.Cell("C7").Style.Font.FontName));
        Assert.Equal(17_179_869_184L, result.CellsAffected);
    }

    [Fact]
    public void A_rectangle_over_the_limit_is_rejected_and_nothing_is_applied()
    {
        using var env = new ExcelEnv(maxCellsPerFormat: 100);
        var id = Open(env: env);

        var ex = Throws(() => env.Formats.FormatRange(id, "Data", "A1:K10", new FormatSpec { Bold = true }));

        Assert.Equal(ErrorCodes.InvalidRange, ex.Code);
        Assert.Contains("A:A", ex.Hint, StringComparison.Ordinal);
        Assert.False(env.Sessions.Use(id, false, s => s.Workbook.Worksheet("Data").Cell("A1").Style.Font.Bold));
        env.Formats.FormatRange(id, "Data", "A1:J10", new FormatSpec { Bold = true }); // 剛好 100 格
    }

    [Fact]
    public void Too_many_whole_rows_are_rejected()
    {
        var id = Open();
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Formats.FormatRange(id, "Data", "1:20000", new FormatSpec { Bold = true })).Code);
        _env.Formats.FormatRange(id, "Data", "1:10000", new FormatSpec { Bold = true });
    }

    // ---- 驗證：全有或全無 ----

    [Fact]
    public void An_invalid_property_means_nothing_at_all_is_applied()
    {
        var id = Open();

        var ex = Throws(() => _env.Formats.FormatRange(id, "Data", "A1", new FormatSpec { Bold = true, FillColor = "notacolor" }));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("FillColor", ex.Message, StringComparison.Ordinal);
        Assert.False(Style(id, ws => ws.Cell("A1").Style.Font.Bold));
        Assert.False(IsDirty());
    }

    [Fact]
    public void An_empty_spec_is_rejected() =>
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Formats.FormatRange(Open(), "Data", "A1", new FormatSpec())).Code);

    [Theory]
    [InlineData("#FF0000")]
    [InlineData("FF0000")]
    [InlineData("ff0000")]
    [InlineData("#F00")]
    [InlineData("red")]
    [InlineData("RED")]
    [InlineData("  blue  ")]
    public void Accepted_color_forms(string color)
    {
        var id = Open();
        _env.Formats.FormatRange(id, "Data", "A1", new FormatSpec { FontColor = color });
        Assert.NotEqual(XLColor.NoColor, Style(id, ws => ws.Cell("A1").Style.Font.FontColor));
    }

    [Theory]
    [InlineData("notacolor")]
    [InlineData("#GG0000")]
    [InlineData("#12345678")]
    [InlineData("#12")]
    [InlineData("")]
    [InlineData("rgb(1,2,3)")]
    public void Rejected_color_forms_instead_of_silently_becoming_transparent(string color)
    {
        var ex = Throws(() => _env.Formats.FormatRange(Open(), "Data", "A1", new FormatSpec { FontColor = color }));
        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("#RRGGBB", ex.Hint, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(0.5)]
    [InlineData(410)]
    [InlineData(-3)]
    [InlineData(double.NaN)]
    public void Invalid_font_sizes_are_rejected(double size) =>
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Formats.FormatRange(Open(), "Data", "A1", new FormatSpec { FontSize = size })).Code);

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("this font name is definitely longer than 31 characters")]
    public void Invalid_font_names_are_rejected(string name) =>
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Formats.FormatRange(Open(), "Data", "A1", new FormatSpec { FontName = name })).Code);

    [Fact]
    public void Font_names_with_chinese_are_fine()
    {
        var id = Open();
        _env.Formats.FormatRange(id, "Data", "A1", new FormatSpec { FontName = "微軟正黑體", FontSize = 409 });
        Assert.Equal("微軟正黑體", Style(id, ws => ws.Cell("A1").Style.Font.FontName));
    }

    [Theory]
    [InlineData("HorizontalAlignment", "middle")]
    [InlineData("VerticalAlignment", "centre")]
    [InlineData("BorderStyle", "bold")]
    [InlineData("BorderSides", "top")]
    public void Unknown_enum_values_are_rejected_listing_the_valid_ones(string field, string value)
    {
        var spec = field switch
        {
            "HorizontalAlignment" => new FormatSpec { HorizontalAlignment = value },
            "VerticalAlignment" => new FormatSpec { VerticalAlignment = value },
            "BorderStyle" => new FormatSpec { BorderStyle = value },
            _ => new FormatSpec { BorderStyle = "thin", BorderSides = value },
        };

        var ex = Throws(() => _env.Formats.FormatRange(Open(), "Data", "A1", spec));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains(field, ex.Message, StringComparison.Ordinal);
        Assert.Contains("可用的值", ex.Hint, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("0")]
    [InlineData("0.00")]
    [InlineData("#,##0")]
    [InlineData("#,##0.00_);(#,##0.00)")]
    [InlineData("0.0%")]
    [InlineData("0.00E+00")]
    [InlineData("yyyy-mm-dd")]
    [InlineData("yyyy/mm/dd hh:mm:ss")]
    [InlineData("[$-409]d mmm yyyy")]
    [InlineData("[$¥-411]#,##0")]
    [InlineData("[Red]0.00")]
    [InlineData("[Blue]#,##0;[Red]-#,##0")]
    [InlineData("[h]:mm:ss")]
    [InlineData("[>=100]0;[<100]0.0")]
    [InlineData("@")]
    [InlineData("General")]
    [InlineData("\"NT$\"#,##0")]
    [InlineData("0\" 元\"")]
    [InlineData("[DBNum1]0")]
    [InlineData("#,##0;[Red](#,##0);\"-\"")]
    public void Valid_number_formats_are_accepted(string format)
    {
        var id = Open();
        _env.Formats.FormatRange(id, "Data", "A1", new FormatSpec { NumberFormat = format });
        Assert.Equal(format, Style(id, ws => ws.Cell("A1").Style.NumberFormat.Format));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("0.00\"")]
    [InlineData("[Bogus]0")]
    [InlineData("[Rad]0.00")]
    [InlineData("0.00 [x")]
    public void Invalid_number_formats_are_rejected_with_examples(string format)
    {
        var id = Open();
        var before = Style(id, ws => ws.Cell("A1").Style.NumberFormat.Format);

        var ex = Throws(() => _env.Formats.FormatRange(id, "Data", "A1", new FormatSpec { NumberFormat = format }));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("yyyy-mm-dd", ex.Hint, StringComparison.Ordinal);
        Assert.Equal(before, Style(id, ws => ws.Cell("A1").Style.NumberFormat.Format)); // 沒有被改動
        Assert.False(IsDirty());
    }

    [Fact]
    public void A_bracket_inside_quotes_is_not_treated_as_a_format_code() =>
        _env.Formats.FormatRange(Open(), "Data", "A1", new FormatSpec { NumberFormat = "0\" [備註]\"" });

    // ---- 邊框 ----

    [Fact]
    public void Borders_on_every_side_of_every_cell_with_a_color()
    {
        var id = Open();
        _env.Formats.FormatRange(id, "Data", "A1:B2", new FormatSpec { BorderStyle = "thin", BorderColor = "#808080" });

        foreach (var address in new[] { "A1", "B2" })
        {
            var b = Style(id, ws => ws.Cell(address).Style.Border);
            Assert.Equal(XLBorderStyleValues.Thin, b.TopBorder);
            Assert.Equal(XLBorderStyleValues.Thin, b.BottomBorder);
            Assert.Equal(XLBorderStyleValues.Thin, b.LeftBorder);
            Assert.Equal(XLBorderStyleValues.Thin, b.RightBorder);
            Assert.Equal(XLColor.FromHtml("#808080"), b.TopBorderColor);
        }
    }

    [Fact]
    public void Outline_borders_only_touch_the_outer_edge()
    {
        var id = Open();
        _env.Formats.FormatRange(id, "Data", "A1:C3", new FormatSpec { BorderStyle = "medium", BorderSides = "outline" });

        Assert.Equal(XLBorderStyleValues.Medium, Style(id, ws => ws.Cell("A1").Style.Border.TopBorder));
        Assert.Equal(XLBorderStyleValues.Medium, Style(id, ws => ws.Cell("C3").Style.Border.BottomBorder));
        Assert.Equal(XLBorderStyleValues.None, Style(id, ws => ws.Cell("B2").Style.Border.TopBorder));
        Assert.Equal(XLBorderStyleValues.None, Style(id, ws => ws.Cell("B2").Style.Border.LeftBorder));
    }

    [Fact]
    public void Border_style_none_removes_borders()
    {
        var id = Open();
        _env.Formats.FormatRange(id, "Data", "A1:B2", new FormatSpec { BorderStyle = "thin" });
        _env.Formats.FormatRange(id, "Data", "A1:B2", new FormatSpec { BorderStyle = "none" });

        Assert.Equal(XLBorderStyleValues.None, Style(id, ws => ws.Cell("A1").Style.Border.TopBorder));
        Assert.Equal(XLBorderStyleValues.None, Style(id, ws => ws.Cell("B2").Style.Border.RightBorder));
    }

    [Fact]
    public void Border_color_or_sides_without_a_style_is_rejected()
    {
        var id = Open();
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Formats.FormatRange(id, "Data", "A1", new FormatSpec { BorderColor = "red" })).Code);
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Formats.FormatRange(id, "Data", "A1", new FormatSpec { BorderSides = "outline" })).Code);
    }

    [Theory]
    [InlineData("A:A")]
    [InlineData("1:2")]
    [InlineData("A1:XFD1048576")]
    public void Outline_borders_are_rejected_for_whole_columns_rows_and_sheet(string range) =>
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Formats.FormatRange(Open(), "Data", range, new FormatSpec { BorderStyle = "thin", BorderSides = "outline" })).Code);

    // ---- 檔案與錯誤 ----

    [Fact]
    public void Formatted_workbooks_survive_save_and_reopen_and_are_valid_open_xml()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "h";
            ws.Cell("B1").Value = 1234.5;
        });
        var path = _env.Files.ListOpen().Single().Path;

        _env.Formats.FormatRange(id, "Data", "A1:B1", new FormatSpec { Bold = true, FillColor = "#DDEEFF", BorderStyle = "thin", HorizontalAlignment = "center" });
        _env.Formats.FormatRange(id, "Data", "B:B", new FormatSpec { NumberFormat = "#,##0.00", FontColor = "blue" });
        _env.Formats.FormatRange(id, "Data", "3:3", new FormatSpec { Italic = true });
        _env.Formats.FormatRange(id, "Data", "A1:C5", new FormatSpec { BorderStyle = "dotted", BorderSides = "outline" });
        _env.Files.Save(id);
        _env.Files.Close(id);

        var reopened = _env.Files.Open(path).WorkbookId;

        Assert.Empty(ExcelEnv.ValidateXlsx(path));
        Assert.True(Style(reopened, ws => ws.Cell("A1").Style.Font.Bold));
        Assert.Equal("1,234.50", _env.Ranges.ReadRange(reopened, "Data", "B1", new ReadOptions(UseFormattedText: true)).Values[0][0]);
        Assert.True(Style(reopened, ws => ws.Cell("Z3").Style.Font.Italic));
    }

    [Fact]
    public void Errors_use_the_right_codes_and_read_only_workbooks_refuse_formatting()
    {
        var id = Open();
        Assert.Equal(ErrorCodes.SheetNotFound, Throws(() => _env.Formats.FormatRange(id, "Nope", "A1", new FormatSpec { Bold = true })).Code);
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Formats.FormatRange(id, "Data", "zz", new FormatSpec { Bold = true })).Code);
        Assert.Equal(ErrorCodes.SessionNotFound, Throws(() => _env.Formats.FormatRange("wb_nope", "Data", "A1", new FormatSpec { Bold = true })).Code);

        File.Move(_env.MakeWorkbook("ro.xlsx"), _env.Path("macro.xlsm"));
        var ro = _env.Files.Open(_env.Path("macro.xlsm")).WorkbookId;
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Formats.FormatRange(ro, "Data", "A1", new FormatSpec { Bold = true })).Code);
    }
}

public sealed class ColumnWidthTests : IDisposable
{
    private readonly ExcelEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    private string Open(Action<IXLWorksheet>? fill = null) =>
        _env.Files.Open(_env.MakeWorkbook($"{Guid.NewGuid():N}.xlsx", fill ?? (ws => { ws.Cell("A1").Value = "x"; }))).WorkbookId;

    private double Width(string id, int column) => _env.Sessions.Use(id, false, s => s.Workbook.Worksheet("Data").Column(column).Width);

    // ---- SetColumnWidth ----

    [Theory]
    [InlineData("B", 2, 2)]
    [InlineData("b", 2, 2)]
    [InlineData("$B", 2, 2)]
    [InlineData("B:D", 2, 4)]
    [InlineData("D:B", 2, 4)]
    [InlineData("B:B", 2, 2)]
    public void Sets_the_width_of_one_or_several_columns(string columns, int first, int last)
    {
        var id = Open();

        var result = _env.Formats.SetColumnWidth(id, "Data", columns, 25.5);

        Assert.Equal(last - first + 1, result.Columns.Count);
        Assert.All(result.Columns, c => Assert.Equal(25.5, c.Width));
        for (var c = first; c <= last; c++)
        {
            Assert.Equal(25.5, Width(id, c));
        }

        Assert.NotEqual(25.5, Width(id, 1));
        Assert.NotEqual(25.5, Width(id, last + 1));
    }

    [Fact]
    public void Width_zero_hides_and_the_maximum_is_accepted()
    {
        var id = Open();

        var hidden = _env.Formats.SetColumnWidth(id, "Data", "A", 0);
        _env.Formats.SetColumnWidth(id, "Data", "B", 255);

        Assert.Contains("隱藏", Assert.Single(hidden.Notes), StringComparison.Ordinal);
        Assert.Equal(0, Width(id, 1));
        Assert.Equal(255, Width(id, 2));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(255.1)]
    [InlineData(1000)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void Invalid_widths_are_rejected(double width)
    {
        var id = Open();
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Formats.SetColumnWidth(id, "Data", "A", width)).Code);
    }

    [Theory]
    [InlineData("")]
    [InlineData("A1")]
    [InlineData("1")]
    [InlineData("A:1")]
    [InlineData("1:3")]
    [InlineData("XFE")]
    [InlineData("A:B:C")]
    public void Invalid_column_specs_are_rejected(string columns) =>
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Formats.SetColumnWidth(Open(), "Data", columns, 10)).Code);

    [Fact]
    public void Many_columns_are_applied_but_only_the_first_200_are_listed()
    {
        var id = Open();
        var result = _env.Formats.SetColumnWidth(id, "Data", "A:ZZ", 12);

        Assert.Equal(200, result.Columns.Count);
        Assert.Equal(12, Width(id, 702));
    }

    // ---- AutoFitColumns ----

    [Fact]
    public void Autofit_sizes_each_column_to_its_widest_content()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "id";
            ws.Cell("A2").Value = "a much longer english sentence here";
            ws.Cell("B1").Value = "short";
            ws.Cell("B2").Value = "tiny";
        });

        var result = _env.Formats.AutoFitColumns(id, "Data");

        Assert.Equal(["A", "B"], result.Columns.Select(c => c.Column));
        Assert.True(Width(id, 1) > 28 && Width(id, 1) < 34, $"A = {Width(id, 1)}");
        Assert.True(Width(id, 2) > 5 && Width(id, 2) < 8, $"B = {Width(id, 2)}");
        Assert.Empty(result.Unchanged);
        Assert.Contains("估算", Assert.Single(result.Notes), StringComparison.Ordinal);
        Assert.True(Width(id, 1) > Width(id, 2));
    }

    [Fact]
    public void Chinese_characters_count_double_unlike_closedxml_on_a_mac()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "這是一段比較長的中文內容用來測試欄寬"; // 18 個全形字
            ws.Cell("B1").Value = "abcdefghijklmnopqr";                     // 18 個半形字
        });

        _env.Formats.AutoFitColumns(id, "Data");

        Assert.True(Width(id, 1) > 36, $"中文欄只有 {Width(id, 1)}"); // ClosedXML 在 Mac 上只算出 9.75
        Assert.True(Width(id, 1) > Width(id, 2) * 1.8, $"{Width(id, 1)} vs {Width(id, 2)}");
    }

    [Fact]
    public void Number_formats_count_as_displayed()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = 1.5;
            ws.Cell("B1").Value = 1.5;
            ws.Cell("B1").Style.NumberFormat.Format = "0.000000000000"; // 顯示成 1.500000000000
            ws.Cell("C1").Value = 1.5;
            ws.Cell("C1").Style.NumberFormat.Format = "0\" 公斤\"";    // 顯示成 2 公斤（含全形字）
        });

        _env.Formats.AutoFitColumns(id, "Data");

        Assert.True(Width(id, 2) > Width(id, 1) * 3, $"{Width(id, 2)} vs {Width(id, 1)}");
        Assert.True(Width(id, 3) > Width(id, 1), $"{Width(id, 3)} vs {Width(id, 1)}");
    }

    [Fact]
    public void Bigger_and_bold_fonts_need_wider_columns()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "same text";
            ws.Cell("B1").Value = "same text";
            ws.Cell("B1").Style.Font.FontSize = 22;
            ws.Cell("C1").Value = "same text";
            ws.Cell("C1").Style.Font.Bold = true;
        });

        _env.Formats.AutoFitColumns(id, "Data");

        Assert.True(Width(id, 2) > Width(id, 1) * 1.8);
        Assert.True(Width(id, 3) > Width(id, 1));
    }

    [Fact]
    public void Multi_line_text_uses_the_longest_line()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "abc\nabcdefghijklmnopqrst\nab";
            ws.Cell("A1").Style.Alignment.WrapText = false; // ClosedXML 指派含換行的字串時會自動設為自動換行，這裡關掉才會納入估算
            ws.Cell("B1").Value = "abcdefghijklmnopqrst";
        });

        _env.Formats.AutoFitColumns(id, "Data");

        Assert.Equal(Width(id, 2), Width(id, 1), 2);
    }

    [Fact]
    public void Text_with_line_breaks_is_wrapped_automatically_so_it_is_ignored()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "abc\nabcdefghijklmnopqrstuvwxyz";
            ws.Cell("A2").Value = "ok";
        });

        _env.Formats.AutoFitColumns(id, "Data");

        Assert.True(Width(id, 1) < 5, $"A = {Width(id, 1)}");
    }

    [Fact]
    public void Wrapped_and_merged_cells_are_ignored()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "a very very very very long wrapped text";
            ws.Cell("A1").Style.Alignment.WrapText = true;
            ws.Cell("A2").Value = "ok";
            ws.Cell("B1").Value = "a very very very very long merged text";
            ws.Range("B1:C1").Merge();
            ws.Cell("B2").Value = "ok";
        });

        _env.Formats.AutoFitColumns(id, "Data");

        Assert.True(Width(id, 1) < 5, $"A = {Width(id, 1)}");
        Assert.True(Width(id, 2) < 5, $"B = {Width(id, 2)}");
    }

    [Fact]
    public void Formula_cells_use_their_computed_display_text()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = 123456789;
            ws.Cell("B1").FormulaA1 = "A1*1000"; // 顯示 123456789000，比 A 欄長
        });

        _env.Formats.AutoFitColumns(id, "Data");

        Assert.True(Width(id, 2) > Width(id, 1), $"{Width(id, 2)} vs {Width(id, 1)}");
    }

    [Fact]
    public void A_formula_that_cannot_be_calculated_does_not_break_autofit()
    {
        // 實測：ClosedXML 在計算時，工作表裡只要有一個語法錯誤的公式，其他尚未計算的公式取值時也會一併失敗，
        // 所以這裡的好公式 B1 也會被略過。重點是自動調整照常完成，不會丟例外。
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "some text";
            ws.Cell("B1").FormulaA1 = "SUM(A1:";
        });

        var result = _env.Formats.AutoFitColumns(id, "Data");

        Assert.Equal(["A"], result.Columns.Select(c => c.Column));
        Assert.Equal(["B"], result.Unchanged);
    }

    [Fact]
    public void Empty_columns_are_left_unchanged_and_reported()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "hello";
            ws.Cell("C1").Value = "world";
        });
        _env.Formats.SetColumnWidth(id, "Data", "B", 42);

        var result = _env.Formats.AutoFitColumns(id, "Data");

        Assert.Equal(["B"], result.Unchanged);
        Assert.Equal(42, Width(id, 2));
    }

    [Fact]
    public void Specific_columns_are_fitted_and_others_are_left_alone()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "a long text in column a";
            ws.Cell("B1").Value = "a long text in column b";
        });
        var before = Width(id, 1);

        var result = _env.Formats.AutoFitColumns(id, "Data", "B");

        Assert.Equal(["B"], result.Columns.Select(c => c.Column));
        Assert.Equal(before, Width(id, 1));
        Assert.NotEqual(before, Width(id, 2));
    }

    [Fact]
    public void Min_and_max_widths_clamp_the_result()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "x";
            ws.Cell("B1").Value = new string('w', 200);
        });

        _env.Formats.AutoFitColumns(id, "Data", minWidth: 10, maxWidth: 40);

        Assert.Equal(10, Width(id, 1));
        Assert.Equal(40, Width(id, 2));
    }

    [Theory]
    [InlineData(-1, 10)]
    [InlineData(20, 10)]
    [InlineData(0, 256)]
    [InlineData(double.NaN, 10)]
    public void Invalid_min_max_are_rejected(double min, double max) =>
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Formats.AutoFitColumns(Open(), "Data", null, min, max)).Code);

    [Fact]
    public void An_empty_sheet_has_nothing_to_fit()
    {
        var id = _env.Files.Create(_env.Path("empty.xlsx")).WorkbookId;
        var result = _env.Formats.AutoFitColumns(id, "Sheet1");

        Assert.Empty(result.Columns);
        Assert.Empty(result.Unchanged);
    }

    [Fact]
    public void The_result_does_not_depend_on_the_users_culture()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = 1234567.891;
            ws.Cell("A1").Style.NumberFormat.Format = "#,##0.00";
            ws.Cell("B1").Value = new DateTime(2026, 10, 3);
            ws.Cell("B1").Style.NumberFormat.Format = "mmm d, yyyy";
        });
        _env.Formats.AutoFitColumns(id, "Data");
        var expected = new[] { Width(id, 1), Width(id, 2) };

        var old = CultureInfo.CurrentCulture;
        try
        {
            foreach (var culture in new[] { "fr-FR", "de-DE", "ar-SA", "zh-TW" })
            {
                CultureInfo.CurrentCulture = new CultureInfo(culture);
                _env.Formats.SetColumnWidth(id, "Data", "A:B", 1);
                _env.Formats.AutoFitColumns(id, "Data");
                Assert.Equal(expected, new[] { Width(id, 1), Width(id, 2) });
            }
        }
        finally
        {
            CultureInfo.CurrentCulture = old;
        }
    }

    [Fact]
    public void Widths_survive_save_and_the_file_is_valid_open_xml()
    {
        var id = Open(ws => ws.Cell("A1").Value = "some content to fit");
        var path = _env.Files.ListOpen().Single().Path;
        _env.Formats.AutoFitColumns(id, "Data");
        _env.Formats.SetColumnWidth(id, "Data", "C:D", 30);
        var a = Width(id, 1);
        _env.Files.Save(id);

        Assert.Empty(ExcelEnv.ValidateXlsx(path));
        var reopened = _env.Files.Open(path).WorkbookId;
        Assert.Equal(a, Width(reopened, 1), 2);
        Assert.Equal(30, Width(reopened, 3));
    }

    [Fact]
    public void Read_only_workbooks_refuse_width_changes()
    {
        File.Move(_env.MakeWorkbook("ro.xlsx"), _env.Path("macro.xlsm"));
        var ro = _env.Files.Open(_env.Path("macro.xlsm")).WorkbookId;
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Formats.SetColumnWidth(ro, "Data", "A", 10)).Code);
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Formats.AutoFitColumns(ro, "Data")).Code);
    }
}

public class ColumnWidthEstimatorTests
{
    [Theory]
    [InlineData("", 0)]
    [InlineData("0", 1)]
    [InlineData("0123456789", 10)]
    [InlineData("中", 2)]
    [InlineData("中文", 4)]
    [InlineData("，", 2)]       // 全形標點
    [InlineData("ＡＢ", 4)]     // 全形英數
    [InlineData("한글", 4)]     // 韓文
    [InlineData("ひらがな", 8)]
    [InlineData("😀", 2)]       // 表情符號（代理對）
    [InlineData(" ", 0.5)]
    public void Units_per_text(string text, double expected) =>
        Assert.Equal(expected, ColumnWidthEstimator.Units(text), 6);

    [Fact]
    public void Mixed_text_adds_up() =>
        Assert.Equal(4 + 3 * 0.95, ColumnWidthEstimator.Units("中文abc"), 6);

    [Fact]
    public void Font_size_bold_and_padding_are_applied()
    {
        var plain = ColumnWidthEstimator.Estimate("0123456789", 11, bold: false);
        Assert.Equal(11, plain, 6); // 10 單位 + 1 邊距
        Assert.Equal((10 * 2) + 1, ColumnWidthEstimator.Estimate("0123456789", 22, bold: false), 6);
        Assert.Equal((10 * 1.1) + 1, ColumnWidthEstimator.Estimate("0123456789", 11, bold: true), 6);
    }

    [Fact]
    public void Newlines_use_the_longest_line_and_empty_text_is_zero()
    {
        Assert.Equal(ColumnWidthEstimator.Estimate("0123456789", 11, false), ColumnWidthEstimator.Estimate("01\r\n0123456789\n012", 11, false), 6);
        Assert.Equal(0, ColumnWidthEstimator.Estimate(string.Empty, 11, false));
        Assert.Equal(0, ColumnWidthEstimator.Estimate("\n\n", 11, false));
    }
}
