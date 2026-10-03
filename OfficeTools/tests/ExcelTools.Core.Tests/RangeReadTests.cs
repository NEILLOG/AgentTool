using System.Globalization;
using ClosedXML.Excel;
using ExcelTools.Core.Models;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Tests;

public sealed class RangeReadTests : IDisposable
{
    private readonly ExcelEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    /// <summary>在工作表 Data 填入內容後開啟，回傳 workbookId。</summary>
    private string Open(Action<IXLWorksheet> fill, ExcelEnv? env = null)
    {
        var e = env ?? _env;
        return e.Files.Open(e.MakeWorkbook($"{Guid.NewGuid():N}.xlsx", fill)).WorkbookId;
    }

    private static void Sample(IXLWorksheet ws)
    {
        ws.Cell("A1").Value = "name";
        ws.Cell("B1").Value = "qty";
        ws.Cell("C1").Value = "ok";
        ws.Cell("A2").Value = "apple";
        ws.Cell("B2").Value = 3;
        ws.Cell("C2").Value = true;
        ws.Cell("A3").Value = "pear";
        ws.Cell("B3").Value = 4.5;
        // C3 留空
    }

    // ---- 基本讀取 ----

    [Fact]
    public void Reads_a_rectangle_with_each_value_type_and_blanks_as_null()
    {
        var id = Open(ws =>
        {
            Sample(ws);
            ws.Cell("A4").Value = new DateTime(2026, 10, 3);
            ws.Cell("B4").Value = new DateTime(2026, 10, 3, 14, 5, 9);
        });

        var data = _env.Ranges.ReadRange(id, "Data", "A1:C4");

        Assert.Equal("Data", data.Sheet);
        Assert.Equal("A1:C4", data.Range);
        Assert.Equal("A1:C4", data.UsedRange);
        Assert.False(data.Truncated);
        Assert.Null(data.NextRange);
        Assert.Null(data.Formulas);
        Assert.Empty(data.CalculationWarnings);
        Assert.Equal(["name", "qty", "ok"], data.Values[0]);
        Assert.Equal(["apple", 3.0, true], data.Values[1]);
        Assert.Equal(["pear", 4.5, null], data.Values[2]);
        Assert.Equal(["2026-10-03", "2026-10-03T14:05:09", null], data.Values[3]);
    }

    [Fact]
    public void Reads_a_single_cell_as_one_by_one()
    {
        var id = Open(Sample);
        var data = _env.Ranges.ReadRange(id, "Data", "B2");

        Assert.Equal("B2", data.Range);
        Assert.Equal(3.0, Assert.Single(Assert.Single(data.Values)));
    }

    [Theory]
    [InlineData("a1:c2")]
    [InlineData("$A$1:$C$2")]
    [InlineData("C2:A1")]
    [InlineData("  A1:C2  ")]
    public void Accepts_the_tolerated_address_forms(string range) =>
        Assert.Equal("A1:C2", _env.Ranges.ReadRange(Open(Sample), "data", range).Range);

    [Fact]
    public void Range_is_limited_to_the_used_area_but_keeps_the_requested_start()
    {
        var id = Open(Sample);

        var data = _env.Ranges.ReadRange(id, "Data", "B2:Z1000");

        Assert.Equal("B2:C3", data.Range);
        Assert.Equal([[3.0, true], [4.5, null]], data.Values);
    }

    [Theory]
    [InlineData("A5:C9")]
    [InlineData("D1:F3")]
    [InlineData("Z100")]
    public void A_range_entirely_outside_the_used_area_is_empty(string range)
    {
        var id = Open(Sample);
        var data = _env.Ranges.ReadRange(id, "Data", range);

        Assert.Empty(data.Values);
        Assert.Equal(range, data.Range);
        Assert.Equal("A1:C3", data.UsedRange);
        Assert.False(data.Truncated);
    }

    [Fact]
    public void An_empty_sheet_reads_as_empty_with_no_used_range()
    {
        var id = _env.Files.Create(_env.Path("empty.xlsx")).WorkbookId;
        var data = _env.Ranges.ReadRange(id, "Sheet1", "A1:C3");

        Assert.Empty(data.Values);
        Assert.Null(data.UsedRange);
    }

    [Fact]
    public void Whole_columns_and_rows_are_limited_to_the_used_area()
    {
        var id = Open(Sample);

        var column = _env.Ranges.ReadRange(id, "Data", "B:B");
        var row = _env.Ranges.ReadRange(id, "Data", "2:2");
        var all = _env.Ranges.ReadRange(id, "Data", "A1:XFD1048576");

        Assert.Equal("B1:B3", column.Range);
        Assert.Equal([["qty"], [3.0], [4.5]], column.Values);
        Assert.Equal("A2:C2", row.Range);
        Assert.Equal(["apple", 3.0, true], row.Values[0]);
        Assert.Equal("A1:C3", all.Range);
        Assert.Equal(3, all.Values.Length);
    }

    [Fact]
    public void Merged_cells_only_have_a_value_in_the_top_left_cell()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "merged";
            ws.Range("A1:B2").Merge();
            ws.Cell("C1").Value = "x";
            ws.Cell("C3").Value = "y";
        });

        var data = _env.Ranges.ReadRange(id, "Data", "A1:C3");

        Assert.Equal(new object?[] { "merged", null, "x" }, data.Values[0]);
        Assert.Equal(new object?[] { null, null, null }, data.Values[1]);
        Assert.Equal(new object?[] { null, null, "y" }, data.Values[2]);
    }

    [Fact]
    public void A_range_that_cuts_through_a_merged_area_never_returns_cells_outside_the_request()
    {
        var id = Open(ws =>
        {
            ws.Cell("B2").Value = "top-left";
            ws.Range("B2:D4").Merge();
            ws.Cell("F6").Value = "far";
        });

        // 只要求合併區的一部分：不能崩潰，也不能多回傳
        var data = _env.Ranges.ReadRange(id, "Data", "C3:C3");

        Assert.Equal("C3", data.Range);
        Assert.Equal(new object?[] { null }, data.Values[0]);
        Assert.Equal("top-left", _env.Ranges.ReadRange(id, "Data", "B2").Values[0][0]);
    }

    // ---- 截斷與續讀 ----

    [Fact]
    public void Large_ranges_are_truncated_by_whole_rows_and_can_be_read_to_the_end_with_NextRange()
    {
        using var env = new ExcelEnv(maxCellsPerRead: 2000);
        var id = Open(ws =>
        {
            for (var r = 1; r <= 500; r++)
            {
                for (var c = 1; c <= 10; c++)
                {
                    ws.Cell(r, c).Value = (r * 100) + c;
                }
            }
        }, env);

        var first = env.Ranges.ReadRange(id, "Data", "A1:J500");
        Assert.True(first.Truncated);
        Assert.Equal("A1:J200", first.Range);
        Assert.Equal("A201:J500", first.NextRange);
        Assert.Equal(200, first.Values.Length);

        var rows = new List<object?[]>(first.Values);
        var chunks = 1;
        for (var current = first; current.NextRange is not null; chunks++)
        {
            current = env.Ranges.ReadRange(id, "Data", current.NextRange);
            rows.AddRange(current.Values);
        }

        Assert.Equal(3, chunks);
        Assert.Equal(500, rows.Count);
        for (var r = 0; r < 500; r++)
        {
            Assert.Equal(((r + 1) * 100) + 1.0, rows[r][0]);
            Assert.Equal(((r + 1) * 100) + 10.0, rows[r][9]);
        }
    }

    [Fact]
    public void A_range_that_exactly_fits_the_limit_is_not_truncated()
    {
        using var env = new ExcelEnv(maxCellsPerRead: 20);
        var id = Open(ws =>
        {
            for (var r = 1; r <= 4; r++)
            {
                for (var c = 1; c <= 5; c++)
                {
                    ws.Cell(r, c).Value = r;
                }
            }
        }, env);

        var data = env.Ranges.ReadRange(id, "Data", "A1:E4");

        Assert.False(data.Truncated);
        Assert.Null(data.NextRange);
        Assert.Equal(4, data.Values.Length);
    }

    [Fact]
    public void A_single_row_wider_than_the_limit_is_still_returned_whole_one_row_at_a_time()
    {
        using var env = new ExcelEnv(maxCellsPerRead: 5);
        var id = Open(ws =>
        {
            for (var r = 1; r <= 3; r++)
            {
                for (var c = 1; c <= 8; c++)
                {
                    ws.Cell(r, c).Value = (r * 10) + c;
                }
            }
        }, env);

        var first = env.Ranges.ReadRange(id, "Data", "A1:H3");

        Assert.True(first.Truncated);
        Assert.Equal("A1:H1", first.Range);
        Assert.Equal(8, first.Values[0].Length);
        Assert.Equal("A2:H3", first.NextRange);
    }

    [Fact]
    public void Truncating_a_whole_column_request_gives_a_cell_range_for_the_next_chunk()
    {
        using var env = new ExcelEnv(maxCellsPerRead: 10);
        var id = Open(ws =>
        {
            for (var r = 1; r <= 25; r++)
            {
                ws.Cell(r, 1).Value = r;
            }
        }, env);

        var first = env.Ranges.ReadRange(id, "Data", "A:A");

        Assert.Equal("A1:A10", first.Range);
        Assert.Equal("A11:A25", first.NextRange);
    }

    // ---- 公式與顯示文字 ----

    [Fact]
    public void IncludeFormulas_returns_formulas_only_for_formula_cells_and_values_are_computed()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = 5;
            ws.Cell("B1").FormulaA1 = "A1*2";
            ws.Cell("C1").Value = "plain";
        });

        var data = _env.Ranges.ReadRange(id, "Data", "A1:C1", new ReadOptions(IncludeFormulas: true));

        Assert.Equal([5.0, 10.0, "plain"], data.Values[0]);
        Assert.Equal(new string?[] { null, "=A1*2", null }, data.Formulas![0]);
    }

    [Fact]
    public void UseFormattedText_returns_the_displayed_text_with_blank_cells_as_null()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = 1234.5;
            ws.Cell("A1").Style.NumberFormat.Format = "#,##0.00";
            ws.Cell("B1").Value = 0.256;
            ws.Cell("B1").Style.NumberFormat.Format = "0.0%";
            ws.Cell("C1").Value = new DateTime(2026, 10, 3);
            ws.Cell("C1").Style.NumberFormat.Format = "yyyy/mm/dd";
            ws.Cell("D1").Value = "text";
            ws.Cell("E1").Value = true;
            ws.Cell("G1").Value = "after blank";
        });

        var data = _env.Ranges.ReadRange(id, "Data", "A1:G1", new ReadOptions(UseFormattedText: true));

        Assert.Equal(new object?[] { "1,234.50", "25.6%", "2026/10/03", "text", "TRUE", null, "after blank" }, data.Values[0]);
    }

    [Theory]
    [InlineData("fr-FR")]
    [InlineData("de-DE")]
    [InlineData("ar-SA")]
    public void Formatted_text_does_not_depend_on_the_users_culture(string culture)
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = 1234.5;
            ws.Cell("A1").Style.NumberFormat.Format = "#,##0.00";
            ws.Cell("B1").Value = new DateTime(2026, 10, 3);
            ws.Cell("B1").Style.NumberFormat.Format = "mmm d, yyyy";
        });

        var old = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo(culture);
            var data = _env.Ranges.ReadRange(id, "Data", "A1:B1", new ReadOptions(UseFormattedText: true));
            Assert.Equal(["1,234.50", "Oct 3, 2026"], data.Values[0]);
        }
        finally
        {
            CultureInfo.CurrentCulture = old;
        }
    }

    [Fact]
    public void Formula_errors_are_returned_as_error_text()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").FormulaA1 = "1/0";
            ws.Cell("A2").FormulaA1 = "NOSUCHFN(1)";
        });

        var data = _env.Ranges.ReadRange(id, "Data", "A1:A2");

        Assert.Equal([["#DIV/0!"], ["#NAME?"]], data.Values);
        Assert.Empty(data.CalculationWarnings);
    }

    [Fact]
    public void Formulas_that_cannot_be_calculated_fall_back_to_the_cached_value_and_warn()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "before";
            ws.Cell("B1").FormulaA1 = "SUM(A1:";   // 語法錯誤（ClosedXML 不會在設定時檢查）
            ws.Cell("C1").FormulaA1 = "A1 B1";     // 範圍交集：ClosedXML 0.105 解析不了
            ws.Cell("D1").Value = "after";
        });

        var data = _env.Ranges.ReadRange(id, "Data", "A1:D1");

        Assert.Equal("before", data.Values[0][0]);
        Assert.Equal("after", data.Values[0][3]);
        Assert.Equal(2, data.CalculationWarnings.Count);
        Assert.Contains(data.CalculationWarnings, w => w.StartsWith("B1：", StringComparison.Ordinal) && w.Contains("語法錯誤", StringComparison.Ordinal));
        Assert.Contains(data.CalculationWarnings, w => w.StartsWith("C1：", StringComparison.Ordinal) && w.Contains("快取值", StringComparison.Ordinal));
    }

    [Fact]
    public void Calculation_warnings_are_capped()
    {
        var id = Open(ws =>
        {
            for (var r = 1; r <= 30; r++)
            {
                ws.Cell(r, 1).FormulaA1 = "SUM(A1:";
            }
        });

        var data = _env.Ranges.ReadRange(id, "Data", "A1:A30");

        Assert.Equal(21, data.CalculationWarnings.Count);
        Assert.Contains("還有 10 個", data.CalculationWarnings[^1], StringComparison.Ordinal);
    }

    // ---- 副作用與錯誤 ----

    [Fact]
    public void Reading_does_not_change_the_workbook()
    {
        var id = Open(Sample);
        var before = _env.Sheets.GetSheetInfo(id, "Data").UsedRange;

        _env.Ranges.ReadRange(id, "Data", "A1:Z200");
        _env.Ranges.ReadRange(id, "Data", "A:Z");

        Assert.Equal(before, _env.Sheets.GetSheetInfo(id, "Data").UsedRange);
        Assert.False(_env.Files.ListOpen().Single().IsDirty);
    }

    [Fact]
    public void Read_only_workbooks_can_be_read()
    {
        File.Move(_env.MakeWorkbook("book.xlsx"), _env.Path("macro.xlsm"));
        var id = _env.Files.Open(_env.Path("macro.xlsm")).WorkbookId;

        Assert.Equal("apple", _env.Ranges.ReadRange(id, "Data", "A2").Values[0][0]);
    }

    [Fact]
    public void Errors_are_reported_with_the_right_codes()
    {
        var id = Open(Sample);

        Assert.Equal(ErrorCodes.SheetNotFound, Throws(() => _env.Ranges.ReadRange(id, "Nope", "A1")).Code);
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.ReadRange(id, "Data", "Sheet1!A1")).Code);
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.ReadRange(id, "Data", "")).Code);
        Assert.Equal(ErrorCodes.SessionNotFound, Throws(() => _env.Ranges.ReadRange("wb_nope", "Data", "A1")).Code);
    }
}
