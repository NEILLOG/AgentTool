using ClosedXML.Excel;
using ExcelTools.Core.Internal;
using ExcelTools.Core.Models;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Tests;

public sealed class RangeWriteTests : IDisposable
{
    private readonly ExcelEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    private string Open(Action<IXLWorksheet>? fill = null, ExcelEnv? env = null)
    {
        var e = env ?? _env;
        return e.Files.Open(e.MakeWorkbook($"{Guid.NewGuid():N}.xlsx", fill ?? (_ => { }))).WorkbookId;
    }

    private object?[][] Read(string id, string range) => _env.Ranges.ReadRange(id, "Data", range).Values;

    // ---- WriteRange ----

    [Fact]
    public void Writes_a_block_from_a_single_cell_and_reports_what_was_written()
    {
        var id = Open();

        var result = _env.Ranges.WriteRange(id, "Data", "B2", [["a", 1, true], ["b", 2.5, false]]);

        Assert.Equal("Data", result.Sheet);
        Assert.Equal("B2:D3", result.Range);
        Assert.Equal(6, result.CellsAffected);
        Assert.Empty(result.Warnings);
        Assert.Equal(new object?[] { "a", 1.0, true }, Read(id, "B2:D3")[0]);
        Assert.Equal(new object?[] { "b", 2.5, false }, Read(id, "B2:D3")[1]);
        Assert.True(_env.Files.ListOpen().Single().IsDirty);
    }

    [Fact]
    public void Writes_into_a_range_of_exactly_the_same_size()
    {
        var id = Open();
        var result = _env.Ranges.WriteRange(id, "Data", "A1:B2", [[1, 2], [3, 4]]);

        Assert.Equal("A1:B2", result.Range);
        Assert.Equal([[1.0, 2.0], [3.0, 4.0]], Read(id, "A1:B2"));
    }

    [Theory]
    [InlineData("A1:C3")]
    [InlineData("A1:A2")]
    [InlineData("A1:B1")]
    public void A_range_whose_size_differs_from_the_data_is_rejected_and_nothing_is_written(string range)
    {
        var id = Open();
        var ex = Throws(() => _env.Ranges.WriteRange(id, "Data", range, [[1, 2], [3, 4]]));

        Assert.Equal(ErrorCodes.InvalidRange, ex.Code);
        Assert.Contains("2 列 × 2 欄", ex.Message, StringComparison.Ordinal);
        Assert.Empty(Read(id, "A1:C3"));
    }

    [Theory]
    [InlineData("A:A")]
    [InlineData("1:1")]
    public void Whole_columns_and_rows_are_not_valid_write_targets(string range) =>
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.WriteRange(Open(), "Data", range, [[1]])).Code);

    [Fact]
    public void Strings_stay_text_dates_and_numbers_are_typed_and_null_clears()
    {
        var id = Open(ws => ws.Cell("D1").Value = "will be cleared");

        _env.Ranges.WriteRange(id, "Data", "A1", [["00123", "2026-10-03", "2026-10-03T14:05:09", 7, null, "-", "中文 😀"]]);
        _env.Ranges.WriteRange(id, "Data", "D1", [[null]]);

        Assert.Equal(new object?[] { "00123", "2026-10-03", "2026-10-03T14:05:09", null, null, "-", "中文 😀" }, Read(id, "A1:G1")[0]);
        // 日期字串被轉成真正的日期（讀回仍是 ISO 字串），而 00123 保持文字
        var types = _env.Sessions.Use(id, false, s => new[] { "A1", "B1", "C1" }.Select(a => s.Workbook.Worksheet(1).Cell(a).DataType).ToArray());
        Assert.Equal([XLDataType.Text, XLDataType.DateTime, XLDataType.DateTime], types);
    }

    [Fact]
    public void Rows_may_have_different_lengths_and_missing_cells_are_left_alone()
    {
        var id = Open(ws =>
        {
            ws.Cell("A2").Value = "keep A2";
            ws.Cell("C2").Value = "keep C2";
        });

        var result = _env.Ranges.WriteRange(id, "Data", "A1", [["a", "b", "c"], ["x"]]);

        Assert.Equal("A1:C2", result.Range);
        Assert.Equal(new object?[] { "a", "b", "c" }, Read(id, "A1:C2")[0]);
        Assert.Equal(new object?[] { "x", null, "keep C2" }, Read(id, "A1:C2")[1]);
    }

    [Fact]
    public void A_null_row_is_skipped()
    {
        var id = Open(ws => ws.Cell("A2").Value = "keep");
        _env.Ranges.WriteRange(id, "Data", "A1", [["first"], null!, ["third"]]);

        Assert.Equal([["first"], ["keep"], ["third"]], Read(id, "A1:A3"));
    }

    [Fact]
    public void Writing_keeps_the_existing_formatting_of_the_cells()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "old";
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.Yellow;
        });

        _env.Ranges.WriteRange(id, "Data", "A1", [["new"]]);

        var cell = _env.Sessions.Use(id, false, s => (s.Workbook.Worksheet(1).Cell("A1").Style.Font.Bold, s.Workbook.Worksheet(1).Cell("A1").Style.Fill.BackgroundColor));
        Assert.True(cell.Bold);
        Assert.Equal(XLColor.Yellow, cell.BackgroundColor);
        Assert.Equal("new", Read(id, "A1")[0][0]);
    }

    // ---- 邊界與大小 ----

    [Fact]
    public void Writing_beyond_the_sheet_edge_is_rejected_and_nothing_is_written()
    {
        var id = Open();

        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.WriteRange(id, "Data", "XFD1", [[1, 2]])).Code);
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.WriteRange(id, "Data", "A1048576", [[1], [2]])).Code);
        Assert.Empty(Read(id, "A1:XFD1048576"));
    }

    [Fact]
    public void Writing_exactly_at_the_last_cell_works()
    {
        var id = Open();
        _env.Ranges.WriteRange(id, "Data", "XFD1048576", [["end"]]);
        Assert.Equal("end", Read(id, "XFD1048576")[0][0]);
    }

    [Theory]
    [InlineData(0)]
    public void Empty_data_is_rejected(int rows)
    {
        var id = Open();
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Ranges.WriteRange(id, "Data", "A1", new object?[rows][])).Code);
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Ranges.WriteRange(id, "Data", "A1", [[], []])).Code);
    }

    [Fact]
    public void Data_over_the_write_limit_is_rejected_with_advice_to_split()
    {
        using var env = new ExcelEnv(maxCellsPerWrite: 10);
        var id = Open(env: env);
        var rows = Enumerable.Range(0, 4).Select(_ => new object?[] { 1, 2, 3 }).ToArray();

        var ex = Throws(() => env.Ranges.WriteRange(id, "Data", "A1", rows));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("分成多次", ex.Hint, StringComparison.Ordinal);
        env.Ranges.WriteRange(id, "Data", "A1", rows[..3]); // 剛好 9 格可以
    }

    // ---- 全有或全無 ----

    [Fact]
    public void One_bad_value_means_nothing_at_all_is_written_and_the_error_names_the_cell()
    {
        var id = Open(ws => ws.Cell("A1").Value = "original");

        var ex = Throws(() => _env.Ranges.WriteRange(id, "Data", "A1", [["new 1", "new 2"], ["new 3", double.NaN]]));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.StartsWith("B2：", ex.Message, StringComparison.Ordinal);
        Assert.Equal([["original"]], Read(id, "A1:B2"));
    }

    [Fact]
    public void An_unsupported_value_type_is_rejected_before_any_write()
    {
        var id = Open();
        var ex = Throws(() => _env.Ranges.WriteRange(id, "Data", "A1", [["fine", Guid.NewGuid()]]));

        Assert.StartsWith("B1：", ex.Message, StringComparison.Ordinal);
        Assert.Empty(Read(id, "A1:B1"));
    }

    // ---- 公式 ----

    [Fact]
    public void Formulas_are_written_as_formulas_and_computed()
    {
        var id = Open();
        _env.Ranges.WriteRange(id, "Data", "A1", [[5, "=A1*2", "=SUM(A1:B1)"]]);

        var data = _env.Ranges.ReadRange(id, "Data", "A1:C1", new ReadOptions(IncludeFormulas: true));

        Assert.Equal(new object?[] { 5.0, 10.0, 15.0 }, data.Values[0]);
        Assert.Equal(new string?[] { null, "=A1*2", "=SUM(A1:B1)" }, data.Formulas![0]);
    }

    [Theory]
    [InlineData("=SUM(A1:")]
    [InlineData("=1+")]
    [InlineData("=(1+2")]
    [InlineData("=SUM(1,2))")]
    [InlineData("=\"unterminated")]
    [InlineData("=IF(A1>1,\"x\"")]
    public void A_formula_with_a_syntax_error_is_rejected_naming_the_cell_and_nothing_is_written(string formula)
    {
        var id = Open();

        var ex = Throws(() => _env.Ranges.WriteRange(id, "Data", "A1", [["first", 1], [2, formula]]));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("B2", ex.Message, StringComparison.Ordinal);
        Assert.Contains(formula[1..], ex.Message, StringComparison.Ordinal);
        Assert.Contains("沒有寫入任何資料", ex.Hint, StringComparison.Ordinal);
        Assert.Empty(Read(id, "A1:B2"));
    }

    [Theory]
    [InlineData("=SUM(A1:A3)")]
    [InlineData("=IF(A1>1,\"x\",\"y\")")]
    [InlineData("=VLOOKUP(A1,Sheet2!A:B,2,FALSE)")]
    [InlineData("=A1&\" \"&B1")]
    [InlineData("=SUMIFS(B:B,A:A,\">0\")")]
    [InlineData("='Other Sheet'!A1+1")]
    [InlineData("=NOSUCHFUNCTION(1)")]
    [InlineData("=-A1^2")]
    [InlineData("=TODAY()")]
    public void Valid_formulas_are_accepted(string formula)
    {
        var id = Open();
        _env.Ranges.WriteRange(id, "Data", "D1", [[formula]]);
        Assert.Equal(formula, _env.Ranges.ReadRange(id, "Data", "D1", new ReadOptions(IncludeFormulas: true)).Formulas![0][0]);
    }

    [Fact]
    public void Formula_parsing_can_be_turned_off_so_text_starting_with_equals_is_kept()
    {
        var id = Open();
        _env.Ranges.WriteRange(id, "Data", "A1", [["=SUM(A1:"]], new WriteOptions(ParseFormulas: false));

        Assert.Equal("=SUM(A1:", Read(id, "A1")[0][0]);
        Assert.Null(_env.Ranges.ReadRange(id, "Data", "A1", new ReadOptions(IncludeFormulas: true)).Formulas![0][0]);
    }

    [Fact]
    public void Date_parsing_can_be_turned_off()
    {
        var id = Open();
        _env.Ranges.WriteRange(id, "Data", "A1", [["2026-10-03"]], new WriteOptions(ParseIsoDates: false));

        var type = _env.Sessions.Use(id, false, s => s.Workbook.Worksheet(1).Cell("A1").DataType);
        Assert.Equal(XLDataType.Text, type);
    }

    [Theory]
    [InlineData("=1/0", "#DIV/0!")]
    [InlineData("=NOSUCHFUNCTION(1)", "#NAME?")]
    [InlineData("=Nope!A1", "#REF!")]
    public void Formulas_that_evaluate_to_errors_produce_warnings_but_are_still_written(string formula, string error)
    {
        var id = Open();

        var result = _env.Ranges.WriteRange(id, "Data", "B2", [[formula]]);

        var warning = Assert.Single(result.Warnings);
        Assert.StartsWith("B2：", warning, StringComparison.Ordinal);
        Assert.Contains(error, warning, StringComparison.Ordinal);
        Assert.Equal(error, Read(id, "B2")[0][0]);
    }

    [Fact]
    public void A_circular_reference_produces_a_warning_but_is_still_written()
    {
        var id = Open();

        var result = _env.Ranges.WriteRange(id, "Data", "A1", [["=A1+1"]]);

        var warning = Assert.Single(result.Warnings);
        Assert.Contains("循環參照", warning, StringComparison.Ordinal);
        Assert.Equal("=A1+1", _env.Ranges.ReadRange(id, "Data", "A1", new ReadOptions(IncludeFormulas: true)).Formulas![0][0]);
    }

    [Fact]
    public void Healthy_formulas_produce_no_warnings() =>
        Assert.Empty(_env.Ranges.WriteRange(Open(), "Data", "A1", [[1, "=A1+1", "=SUM(A1:B1)"]]).Warnings);

    [Fact]
    public void Writing_a_value_over_a_formula_removes_the_formula_and_vice_versa()
    {
        var id = Open(ws => ws.Cell("A1").FormulaA1 = "1+1");

        _env.Ranges.WriteRange(id, "Data", "A1", [["plain"]]);
        Assert.Null(_env.Ranges.ReadRange(id, "Data", "A1", new ReadOptions(IncludeFormulas: true)).Formulas![0][0]);

        _env.Ranges.WriteRange(id, "Data", "A1", [["=2+3"]]);
        Assert.Equal(5.0, Read(id, "A1")[0][0]);
    }

    // ---- 存檔往返與合法性 ----

    [Fact]
    public void Written_data_survives_save_and_reopen_and_the_file_is_valid_open_xml()
    {
        var id = Open();
        var path = _env.Files.ListOpen().Single().Path;
        _env.Ranges.WriteRange(id, "Data", "A1", [["name", "qty", "total", "when"], ["apple", 3, "=B2*2", "2026-10-03"], ["pear", 4.5, "=B3*2", "2026-10-04T08:30:00"], [null, null, "=SUM(C2:C3)", "中文 😀"]]);
        _env.Files.Save(id);
        _env.Files.Close(id);

        var reopened = _env.Files.Open(path).WorkbookId;
        var data = _env.Ranges.ReadRange(reopened, "Data", "A1:D4", new ReadOptions(IncludeFormulas: true));

        Assert.Equal(new object?[] { "apple", 3.0, 6.0, "2026-10-03" }, data.Values[1]);
        Assert.Equal(new object?[] { "pear", 4.5, 9.0, "2026-10-04T08:30:00" }, data.Values[2]);
        Assert.Equal(new object?[] { null, null, 15.0, "中文 😀" }, data.Values[3]);
        Assert.Equal("=SUM(C2:C3)", data.Formulas![3][2]);
        Assert.Empty(ExcelEnv.ValidateXlsx(path));
    }

    [Fact]
    public void Read_only_workbooks_refuse_writes()
    {
        File.Move(_env.MakeWorkbook("book.xlsx"), _env.Path("macro.xlsm"));
        var id = _env.Files.Open(_env.Path("macro.xlsm")).WorkbookId;

        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Ranges.WriteRange(id, "Data", "A1", [[1]])).Code);
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Ranges.AppendRows(id, "Data", [[1]])).Code);
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Ranges.ClearRange(id, "Data", "A1")).Code);
    }

    [Fact]
    public void Errors_use_the_right_codes()
    {
        var id = Open();
        Assert.Equal(ErrorCodes.SheetNotFound, Throws(() => _env.Ranges.WriteRange(id, "Nope", "A1", [[1]])).Code);
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.WriteRange(id, "Data", "nonsense", [[1]])).Code);
        Assert.Equal(ErrorCodes.SessionNotFound, Throws(() => _env.Ranges.WriteRange("wb_nope", "Data", "A1", [[1]])).Code);
    }

    // ---- AppendRows ----

    [Fact]
    public void Append_to_an_empty_sheet_starts_at_A1()
    {
        var id = Open();
        var result = _env.Ranges.AppendRows(id, "Data", [["h1", "h2"], [1, 2]]);

        Assert.Equal("A1:B2", result.Range);
        Assert.Equal([["h1", "h2"], [1.0, 2.0]], Read(id, "A1:B2"));
    }

    [Fact]
    public void Append_adds_below_the_last_row_and_keeps_going_on_repeated_calls()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "name";
            ws.Cell("B1").Value = "qty";
            ws.Cell("A2").Value = "apple";
            ws.Cell("B2").Value = 3;
        });

        var first = _env.Ranges.AppendRows(id, "Data", [["pear", 4]]);
        var second = _env.Ranges.AppendRows(id, "Data", [["fig", 5], ["kiwi", 6]]);

        Assert.Equal("A3:B3", first.Range);
        Assert.Equal("A4:B5", second.Range);
        Assert.Equal(5, _env.Ranges.ReadRange(id, "Data", "A:A").Values.Length);
        Assert.Equal(new object?[] { "kiwi", 6.0 }, Read(id, "A5:B5")[0]);
    }

    [Fact]
    public void Append_aligns_with_the_first_used_column_of_the_data()
    {
        var id = Open(ws =>
        {
            ws.Cell("C3").Value = "h";
            ws.Cell("D3").Value = "h2";
        });

        var result = _env.Ranges.AppendRows(id, "Data", [["x", "y"]]);

        Assert.Equal("C4:D4", result.Range);
    }

    [Fact]
    public void Append_can_start_at_an_explicit_column()
    {
        var id = Open(ws => ws.Cell("A1").Value = "h");
        Assert.Equal("E2:F2", _env.Ranges.AppendRows(id, "Data", [["x", "y"]], startColumn: "e").Range);
    }

    [Fact]
    public void Append_with_formulas_and_validation_is_all_or_nothing()
    {
        var id = Open(ws => ws.Cell("A1").Value = 1);

        var ex = Throws(() => _env.Ranges.AppendRows(id, "Data", [[1, "=A2*2"], [2, "=SUM("]]));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Equal([[1.0]], _env.Ranges.ReadRange(id, "Data", "A1:B10").Values);
    }

    [Fact]
    public void Append_below_the_last_possible_row_is_rejected()
    {
        var id = Open(ws => ws.Cell("A1048576").Value = "last");
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.AppendRows(id, "Data", [[1]])).Code);
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("")]
    [InlineData("1")]
    [InlineData("XFE")]
    [InlineData("A:A")]
    public void Append_with_an_invalid_start_column_is_rejected(string column) =>
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.AppendRows(Open(), "Data", [[1]], startColumn: column)).Code);

    // ---- ClearRange ----

    private string OpenFormatted() => Open(ws =>
    {
        ws.Cell("A1").Value = "x";
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("B1").FormulaA1 = "1+1";
        ws.Cell("A2").Value = "y";
        ws.Cell("C3").Style.Fill.BackgroundColor = XLColor.Red; // 只有格式
    });

    private bool IsBold(string id, string address) =>
        _env.Sessions.Use(id, false, s => s.Workbook.Worksheet(1).Cell(address).Style.Font.Bold);

    [Fact]
    public void Clear_contents_keeps_the_formatting()
    {
        var id = OpenFormatted();

        var result = _env.Ranges.ClearRange(id, "Data", "A1:B2");

        Assert.Equal("A1:B2", result.Range);
        Assert.Equal(4, result.CellsAffected);
        Assert.Empty(Read(id, "A1:B2")); // 內容都清掉了，工作表剩下的有內容區域不再涵蓋這裡
        Assert.True(IsBold(id, "A1"));
    }

    [Fact]
    public void Clear_formats_keeps_the_content()
    {
        var id = OpenFormatted();
        _env.Ranges.ClearRange(id, "Data", "A1", ClearMode.Formats);

        Assert.Equal("x", Read(id, "A1")[0][0]);
        Assert.False(IsBold(id, "A1"));
    }

    [Fact]
    public void Clear_all_removes_content_and_formatting()
    {
        var id = OpenFormatted();
        _env.Ranges.ClearRange(id, "Data", "A1", ClearMode.All);

        Assert.Null(Read(id, "A1")[0][0]); // 工作表別處還有內容，所以仍讀得到這一格，只是空的
        Assert.False(IsBold(id, "A1"));
    }

    [Fact]
    public void Clear_removes_formulas()
    {
        var id = OpenFormatted();
        _env.Ranges.ClearRange(id, "Data", "B1");

        var stillFormula = _env.Sessions.Use(id, false, s => s.Workbook.Worksheet(1).Cell("B1").HasFormula);
        Assert.False(stillFormula);
        Assert.Equal("A1:A2", _env.Sheets.GetSheetInfo(id, "Data").UsedRange); // 公式清掉後 B 欄不再有內容
    }

    [Theory]
    [InlineData("A:A", "A1:A3")]
    [InlineData("1:1", "A1:C1")]
    [InlineData("A1:XFD1048576", "A1:C3")]
    public void Clearing_whole_columns_rows_or_sheet_is_limited_to_the_used_area(string range, string cleared)
    {
        var id = OpenFormatted();
        Assert.Equal(cleared, _env.Ranges.ClearRange(id, "Data", range).Range);
    }

    [Fact]
    public void Clearing_the_whole_sheet_is_fast()
    {
        var id = OpenFormatted();
        var sw = System.Diagnostics.Stopwatch.StartNew();
        _env.Ranges.ClearRange(id, "Data", "A1:XFD1048576", ClearMode.All);
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(5), $"took {sw.Elapsed}");
        Assert.Null(_env.Sheets.GetSheetInfo(id, "Data").UsedRange);
    }

    [Fact]
    public void Clearing_outside_the_used_area_does_nothing()
    {
        var id = OpenFormatted();
        var result = _env.Ranges.ClearRange(id, "Data", "J10:K20");

        Assert.Equal(0, result.CellsAffected);
        Assert.Equal("x", Read(id, "A1")[0][0]);
    }

    [Fact]
    public void Clear_on_an_unknown_sheet_or_bad_range_reports_errors()
    {
        var id = OpenFormatted();
        Assert.Equal(ErrorCodes.SheetNotFound, Throws(() => _env.Ranges.ClearRange(id, "Nope", "A1")).Code);
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.ClearRange(id, "Data", "zzzz")).Code);
    }
}

public class FormulaValidatorTests
{
    [Fact]
    public void Reports_the_address_and_formula_of_the_bad_cell()
    {
        using var validator = new FormulaValidator();
        validator.Validate("A1", "SUM(1,2)");

        var ex = Assert.Throws<OfficeToolException>(() => validator.Validate("C7", "SUM(A1:"));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("C7", ex.Message, StringComparison.Ordinal);
        Assert.Contains("=SUM(A1:", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("SUM(A:A)")]
    [InlineData("SUM(1:3)")]
    [InlineData("50%")]
    [InlineData("{1,2,3}")]
    [InlineData("SUM({1,2;3,4})")]
    [InlineData("'It''s'!A1")]
    [InlineData("SUM(Sheet1:Sheet3!A1)")]
    [InlineData("SUMPRODUCT((A1:A10>5)*(B1:B10))")]
    [InlineData("XLOOKUP(A1,B:B,C:C)")]
    [InlineData("LET(x,1,x+1)")]
    [InlineData("FILTER(A1:A10,B1:B10>0)")]
    [InlineData("Table1[[#This Row],[Col]]")]
    [InlineData("A1#")]
    [InlineData("A1:INDEX(B:B,5)")]
    [InlineData("COUNTIF(A:A,\">=\"&B1)")]
    [InlineData("\"quoted \"\"inner\"\" text\"")]
    [InlineData("SUM(A1,,B1)")]
    public void Real_world_valid_syntax_is_never_rejected(string formula)
    {
        using var validator = new FormulaValidator();
        validator.Validate("A1", formula);
    }

    [Fact]
    public void Can_be_reused_after_a_failure()
    {
        using var validator = new FormulaValidator();
        Assert.Throws<OfficeToolException>(() => validator.Validate("A1", "1+"));
        validator.Validate("A2", "1+2"); // 失敗後仍可繼續使用
    }

    [Fact]
    public void Self_referencing_or_unsupported_but_syntactically_valid_formulas_are_not_rejected()
    {
        using var validator = new FormulaValidator();
        validator.Validate("A1", "ZZ1000+1"); // 暫存格自己
        validator.Validate("A2", "NOSUCHFN(1)");
        validator.Validate("A3", "Nope!A1");
    }
}
