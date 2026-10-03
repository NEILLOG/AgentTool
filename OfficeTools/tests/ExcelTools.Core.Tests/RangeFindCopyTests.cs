using ClosedXML.Excel;
using ExcelTools.Core.Internal;
using ExcelTools.Core.Models;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Tests;

public sealed class RangeFindTests : IDisposable
{
    private readonly ExcelEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    private string OpenFixture()
    {
        var path = _env.MakeBook("find.xlsx", wb =>
        {
            var a = wb.AddWorksheet("Data");
            a.Cell("A1").Value = "Hello World";
            a.Cell("B1").Value = "hello";
            a.Cell("C1").Value = 1234.5;
            a.Cell("C1").Style.NumberFormat.Format = "#,##0.00";
            a.Cell("A2").Value = new DateTime(2026, 10, 3);
            a.Cell("B2").FormulaA1 = "C1*2";
            a.Cell("C2").Value = true;
            a.Cell("A3").Value = "HELLO";
            a.Cell("B3").Value = "say hello there";
            var b = wb.AddWorksheet("Notes");
            b.Cell("A1").Value = "hello again";
            b.Cell("B5").Value = "nothing";
        });
        return _env.Files.Open(path).WorkbookId;
    }

    private static string[] Addresses(FindResult r) => r.Matches.Select(m => $"{m.Sheet}!{m.Address}").ToArray();

    [Fact]
    public void Finds_substrings_case_insensitively_across_all_sheets_in_tab_and_row_major_order()
    {
        var id = OpenFixture();
        var result = _env.Ranges.Find(id, "hello");

        Assert.Equal(["Data!A1", "Data!B1", "Data!A3", "Data!B3", "Notes!A1"], Addresses(result));
        Assert.False(result.Truncated);
        Assert.All(result.Matches, m => Assert.Equal("Text", m.MatchedIn));
    }

    [Fact]
    public void MatchCase_distinguishes_case()
    {
        var id = OpenFixture();
        Assert.Equal(["Data!B1", "Data!B3", "Notes!A1"], Addresses(_env.Ranges.Find(id, "hello", findOptions: new FindOptions(MatchCase: true))));
    }

    [Fact]
    public void WholeCell_requires_the_entire_content_to_match()
    {
        var id = OpenFixture();
        Assert.Equal(["Data!B1", "Data!A3"], Addresses(_env.Ranges.Find(id, "hello", sheet: "Data", findOptions: new FindOptions(WholeCell: true))));
        Assert.Equal(["Data!B1"], Addresses(_env.Ranges.Find(id, "hello", sheet: "Data", findOptions: new FindOptions(WholeCell: true, MatchCase: true))));
    }

    [Fact]
    public void Numbers_match_by_displayed_text_and_by_raw_value()
    {
        var id = OpenFixture();

        var displayed = Assert.Single(_env.Ranges.Find(id, "1,234.50").Matches);
        var raw = Assert.Single(_env.Ranges.Find(id, "1234.5").Matches, m => m.Address == "C1");

        Assert.Equal("Data", displayed.Sheet);
        Assert.Equal("C1", displayed.Address);
        Assert.Equal("Text", displayed.MatchedIn);
        Assert.Equal(1234.5, displayed.Value);
        Assert.Equal("Value", raw.MatchedIn);
    }

    [Fact]
    public void Dates_booleans_and_computed_values_are_searchable()
    {
        var id = OpenFixture();

        Assert.Equal(["Data!A2"], Addresses(_env.Ranges.Find(id, "2026-10-03")));
        Assert.Equal(["Data!C2"], Addresses(_env.Ranges.Find(id, "true", sheet: "Data", findOptions: new FindOptions(WholeCell: true))));
        Assert.Equal(["Data!B2"], Addresses(_env.Ranges.Find(id, "2469", sheet: "Data"))); // C1*2 的計算結果
    }

    [Fact]
    public void Formula_text_is_only_searched_when_asked_and_matches_report_the_formula()
    {
        var id = OpenFixture();

        Assert.Empty(_env.Ranges.Find(id, "C1*").Matches);

        var hit = Assert.Single(_env.Ranges.Find(id, "C1*", findOptions: new FindOptions(SearchFormulas: true)).Matches);
        Assert.Equal("B2", hit.Address);
        Assert.Equal("Formula", hit.MatchedIn);
        Assert.Equal("=C1*2", hit.Formula);
        Assert.Equal(2469.0, hit.Value);
    }

    [Fact]
    public void Sheet_and_range_limits_are_applied()
    {
        var id = OpenFixture();

        Assert.Equal(["Notes!A1"], Addresses(_env.Ranges.Find(id, "hello", sheet: "notes")));
        Assert.Equal(["Data!B1", "Data!B3"], Addresses(_env.Ranges.Find(id, "hello", sheet: "Data", range: "B:B")));
        Assert.Equal(["Data!A1"], Addresses(_env.Ranges.Find(id, "hello", sheet: "Data", range: "A1:A2")));
        Assert.Empty(_env.Ranges.Find(id, "hello", sheet: "Data", range: "Z1:Z9").Matches);
    }

    [Fact]
    public void Results_are_capped_and_flagged_as_truncated()
    {
        var path = _env.MakeBook("many.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("Data");
            for (var r = 1; r <= 250; r++)
            {
                ws.Cell(r, 1).Value = $"row {r}";
            }
        });
        var id = _env.Files.Open(path).WorkbookId;

        var capped = _env.Ranges.Find(id, "row");
        var small = _env.Ranges.Find(id, "row", findOptions: new FindOptions(MaxResults: 5));
        var exact = _env.Ranges.Find(id, "row", findOptions: new FindOptions(MaxResults: 250));

        Assert.Equal(100, capped.Matches.Count);
        Assert.True(capped.Truncated);
        Assert.Equal(["Data!A1", "Data!A2", "Data!A3", "Data!A4", "Data!A5"], Addresses(small));
        Assert.Equal(250, exact.Matches.Count);
        Assert.False(exact.Truncated);
    }

    [Fact]
    public void Formulas_that_cannot_be_calculated_do_not_break_the_search()
    {
        var path = _env.MakeBook("broken.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "needle";
            ws.Cell("B1").FormulaA1 = "SUM(A1:";
        });
        var id = _env.Files.Open(path).WorkbookId;

        var result = _env.Ranges.Find(id, "needle", findOptions: new FindOptions(SearchFormulas: true));
        var byFormula = _env.Ranges.Find(id, "SUM(", findOptions: new FindOptions(SearchFormulas: true));

        Assert.Equal(["Data!A1"], Addresses(result));
        Assert.Single(result.CalculationWarnings);
        Assert.Equal(["Data!B1"], Addresses(byFormula));
    }

    [Fact]
    public void Searching_does_not_change_the_workbook()
    {
        var id = OpenFixture();
        _env.Ranges.Find(id, "hello");
        _env.Ranges.Find(id, "x", sheet: "Data", range: "A1:Z100");

        Assert.False(_env.Files.ListOpen().Single().IsDirty);
        Assert.Equal("A1:C3", _env.Sheets.GetSheetInfo(id, "Data").UsedRange);
    }

    [Fact]
    public void Invalid_arguments_are_rejected()
    {
        var id = OpenFixture();

        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Ranges.Find(id, "")).Code);
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Ranges.Find(id, "x", range: "A1:B2")).Code);
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Ranges.Find(id, "x", findOptions: new FindOptions(MaxResults: 0))).Code);
        Assert.Equal(ErrorCodes.SheetNotFound, Throws(() => _env.Ranges.Find(id, "x", sheet: "Nope")).Code);
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.Find(id, "x", sheet: "Data", range: "zz")).Code);
        Assert.Equal(ErrorCodes.SessionNotFound, Throws(() => _env.Ranges.Find("wb_nope", "x")).Code);
    }

    [Fact]
    public void Searching_an_empty_workbook_finds_nothing()
    {
        var id = _env.Files.Create(_env.Path("empty.xlsx")).WorkbookId;
        Assert.Empty(_env.Ranges.Find(id, "x").Matches);
    }
}

public sealed class RangeCopyTests : IDisposable
{
    private readonly ExcelEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    private string Open(Action<IXLWorkbook> build) => _env.Files.Open(_env.MakeBook($"{Guid.NewGuid():N}.xlsx", build)).WorkbookId;

    private string Formula(string id, string sheet, string address) =>
        _env.Sessions.Use(id, false, s => s.Workbook.Worksheet(sheet).Cell(address).FormulaA1);

    private object?[] Column(string id, string sheet, string range) =>
        _env.Ranges.ReadRange(id, sheet, range).Values.Select(r => r[0]).ToArray();

    private string[] SheetNames(string id) => _env.Sheets.ListSheets(id).Select(s => s.Name).ToArray();

    private string OpenFixture() => Open(wb =>
    {
        var d = wb.AddWorksheet("Data");
        for (var r = 1; r <= 5; r++)
        {
            d.Cell(r, 1).Value = r;
        }

        d.Cell("B1").FormulaA1 = "A1*10";
        d.Cell("C1").FormulaA1 = "$A$1+A$1+$A1";
        d.Cell("E1").Value = "styled";
        d.Cell("E1").Style.Font.Bold = true;
        d.Cell("E1").Style.Fill.BackgroundColor = XLColor.Yellow;
        d.Cell("G1").Value = "m";
        d.Range("G1:H2").Merge();
        wb.AddWorksheet("Other");
    });

    // ---- 全部複製 ----

    [Fact]
    public void Copies_values_and_moves_relative_formula_references_like_excel()
    {
        var id = OpenFixture();

        var result = _env.Ranges.CopyRange(id, "Data", "B1:C1", "B20");

        Assert.Equal("Data", result.Sheet);
        Assert.Equal("B20:C20", result.Range);
        Assert.Equal(2, result.CellsAffected);
        Assert.Empty(result.Warnings);
        Assert.Equal("A20*10", Formula(id, "Data", "B20"));            // 相對參照位移 19 列
        Assert.Equal("$A$1+A$1+$A20", Formula(id, "Data", "C20"));    // 絕對與混合參照維持
        Assert.Equal("A1*10", Formula(id, "Data", "B1"));              // 來源不動
    }

    [Fact]
    public void Copies_formatting_and_merged_cells()
    {
        var id = OpenFixture();
        _env.Ranges.CopyRange(id, "Data", "E1", "E10");
        _env.Ranges.CopyRange(id, "Data", "G1:H2", "G10");

        var (bold, fill, merged) = _env.Sessions.Use(id, false, s =>
        {
            var ws = s.Workbook.Worksheet("Data");
            return (ws.Cell("E10").Style.Font.Bold, ws.Cell("E10").Style.Fill.BackgroundColor, ws.MergedRanges.Select(m => m.RangeAddress.ToStringRelative(false)).ToArray());
        });
        Assert.True(bold);
        Assert.Equal(XLColor.Yellow, fill);
        Assert.Contains("G10:H11", merged);
        Assert.Equal("styled", _env.Ranges.ReadRange(id, "Data", "E10").Values[0][0]);
    }

    [Fact]
    public void Copies_to_another_sheet()
    {
        var id = OpenFixture();

        var result = _env.Ranges.CopyRange(id, "Data", "A1:A3", "C2", targetSheet: "other");

        Assert.Equal("Other", result.Sheet);
        Assert.Equal("C2:C4", result.Range);
        Assert.Equal(new object?[] { 1.0, 2.0, 3.0 }, Column(id, "Other", "C2:C4"));
        Assert.Equal(new object?[] { 1.0, 2.0, 3.0, 4.0, 5.0 }, Column(id, "Data", "A1:A5")); // 來源不動
    }

    [Fact]
    public void Whole_column_sources_are_limited_to_the_used_area()
    {
        var id = OpenFixture();
        var result = _env.Ranges.CopyRange(id, "Data", "A:A", "J1");
        Assert.Equal("J1:J5", result.Range);
    }

    // ---- 重疊 ----

    [Fact]
    public void Overlapping_copy_uses_a_snapshot_of_the_source_like_excel()
    {
        var id = OpenFixture();

        var result = _env.Ranges.CopyRange(id, "Data", "A1:A5", "A3");

        Assert.Equal("A3:A7", result.Range);
        Assert.Equal(new object?[] { 1.0, 2.0, 1.0, 2.0, 3.0, 4.0, 5.0 }, Column(id, "Data", "A1:A7"));
        Assert.Equal(["Data", "Other"], SheetNames(id)); // 暫存工作表已移除
    }

    [Fact]
    public void Overlapping_copy_upwards_also_uses_a_snapshot()
    {
        var id = OpenFixture();
        _env.Ranges.CopyRange(id, "Data", "A3:A5", "A1");
        Assert.Equal(new object?[] { 3.0, 4.0, 5.0, 4.0, 5.0 }, Column(id, "Data", "A1:A5"));
    }

    [Fact]
    public void Overlapping_copy_of_formulas_applies_the_net_offset()
    {
        var id = Open(wb =>
        {
            var d = wb.AddWorksheet("Data");
            for (var r = 1; r <= 4; r++)
            {
                d.Cell(r, 1).Value = r;
                d.Cell(r, 2).FormulaA1 = $"A{r}*10";
            }
        });

        _env.Ranges.CopyRange(id, "Data", "B1:B3", "B2");

        Assert.Equal("A1*10", Formula(id, "Data", "B1"));
        Assert.Equal("A2*10", Formula(id, "Data", "B2")); // 來源 B1 的公式位移 1 列
        Assert.Equal("A3*10", Formula(id, "Data", "B3"));
        Assert.Equal("A4*10", Formula(id, "Data", "B4"));
    }

    [Fact]
    public void The_temporary_sheet_is_removed_even_when_the_copy_fails_midway()
    {
        var id = OpenFixture();
        // 重疊且目標超出邊界會在動工作表前就被擋下，這裡驗證不留任何暫存工作表
        Throws(() => _env.Ranges.CopyRange(id, "Data", "A1:A5", "A1048575"));
        Assert.Equal(["Data", "Other"], SheetNames(id));
    }

    // ---- 只貼值 ----

    [Fact]
    public void ValuesOnly_pastes_computed_values_without_formulas_or_formatting()
    {
        var id = OpenFixture();

        _env.Ranges.CopyRange(id, "Data", "B1:C1", "B20", mode: CopyMode.ValuesOnly);
        _env.Ranges.CopyRange(id, "Data", "E1", "E20", mode: CopyMode.ValuesOnly);

        Assert.Equal(new object?[] { 10.0, 3.0 }, _env.Ranges.ReadRange(id, "Data", "B20:C20").Values[0]);
        Assert.Null(_env.Ranges.ReadRange(id, "Data", "B20:C20", new ReadOptions(IncludeFormulas: true)).Formulas![0][0]);
        var bold = _env.Sessions.Use(id, false, s => s.Workbook.Worksheet("Data").Cell("E20").Style.Font.Bold);
        Assert.False(bold);
        Assert.Equal("styled", _env.Ranges.ReadRange(id, "Data", "E20").Values[0][0]);
    }

    [Fact]
    public void ValuesOnly_handles_overlap_and_blank_source_cells()
    {
        var id = Open(wb =>
        {
            var d = wb.AddWorksheet("Data");
            d.Cell("A1").Value = 1;
            d.Cell("A3").Value = 3; // A2 留空
            d.Cell("B1").Value = "keep me";
            d.Cell("B2").Value = "overwritten by a blank";
        });

        _env.Ranges.CopyRange(id, "Data", "A1:A3", "B1", mode: CopyMode.ValuesOnly);
        Assert.Equal(new object?[] { 1.0, null, 3.0 }, Column(id, "Data", "B1:B3"));
    }

    [Fact]
    public void ValuesOnly_dates_and_errors_come_through()
    {
        var id = Open(wb =>
        {
            var d = wb.AddWorksheet("Data");
            d.Cell("A1").Value = new DateTime(2026, 10, 3);
            d.Cell("A2").FormulaA1 = "1/0";
        });

        _env.Ranges.CopyRange(id, "Data", "A1:A2", "C1", mode: CopyMode.ValuesOnly);

        Assert.Equal(new object?[] { "2026-10-03", "#DIV/0!" }, Column(id, "Data", "C1:C2"));
    }

    // ---- 警告與錯誤 ----

    [Fact]
    public void Overwriting_existing_content_is_reported()
    {
        var id = OpenFixture();

        var result = _env.Ranges.CopyRange(id, "Data", "A1:A2", "A4");

        var warning = Assert.Single(result.Warnings);
        Assert.Contains("A4:A5", warning, StringComparison.Ordinal);
        Assert.Contains("2 個", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void Copying_into_empty_cells_gives_no_warning() =>
        Assert.Empty(_env.Ranges.CopyRange(OpenFixture(), "Data", "A1", "J10").Warnings);

    [Fact]
    public void A_source_with_nothing_in_it_is_rejected()
    {
        var id = OpenFixture();
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Ranges.CopyRange(id, "Data", "Z1:Z9", "A10")).Code);
    }

    [Theory]
    [InlineData("A1:A5", "A1048575")]
    [InlineData("A1:E1", "XFC1")]
    public void Pasting_beyond_the_sheet_edge_is_rejected_and_changes_nothing(string source, string target)
    {
        var id = OpenFixture();
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.CopyRange(id, "Data", source, target)).Code);
        Assert.Equal(1.0, _env.Ranges.ReadRange(id, "Data", "A1").Values[0][0]);
    }

    [Theory]
    [InlineData("A1:B2")]
    [InlineData("A:A")]
    [InlineData("")]
    public void The_target_must_be_a_single_cell(string target) =>
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.CopyRange(OpenFixture(), "Data", "A1", target)).Code);

    [Fact]
    public void Errors_use_the_right_codes_and_read_only_workbooks_refuse_copies()
    {
        var id = OpenFixture();
        Assert.Equal(ErrorCodes.SheetNotFound, Throws(() => _env.Ranges.CopyRange(id, "Nope", "A1", "B1")).Code);
        Assert.Equal(ErrorCodes.SheetNotFound, Throws(() => _env.Ranges.CopyRange(id, "Data", "A1", "B1", targetSheet: "Nope")).Code);
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.CopyRange(id, "Data", "bad", "B1")).Code);

        File.Move(_env.MakeWorkbook("ro.xlsx"), _env.Path("macro.xlsm"));
        var ro = _env.Files.Open(_env.Path("macro.xlsm")).WorkbookId;
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Ranges.CopyRange(ro, "Data", "A1", "B1")).Code);
    }

    [Fact]
    public void Copied_data_survives_save_and_the_file_is_valid_open_xml()
    {
        var id = OpenFixture();
        var path = _env.Files.ListOpen().Single().Path;
        _env.Ranges.CopyRange(id, "Data", "A1:C5", "A1", targetSheet: "Other");
        _env.Ranges.CopyRange(id, "Data", "A1:A5", "A3"); // 重疊（經由暫存工作表）
        _env.Files.Save(id);

        Assert.Empty(ExcelEnv.ValidateXlsx(path));
        Assert.Equal(["Data", "Other"], SheetNames(_env.Files.Open(path).WorkbookId));
    }
}

public class FormulaReferencesTests
{
    private static string[] Refs(string formula, string formulaSheet = "Data", string target = "Data") =>
        FormulaReferences.Find(formula, formulaSheet, target).Select(r => r.ToString()).ToArray();

    [Theory]
    [InlineData("A5*2", "A5")]
    [InlineData("$A$5+A$5+$A5", "A5,A5,A5")]
    [InlineData("SUM(A1:A9)", "A1:A9")]
    [InlineData("SUM(A:A)", "A:A")]
    [InlineData("SUM(1:3)", "1:3")]
    [InlineData("SUM($B$2:$C$4)+D1", "B2:C4,D1")]
    [InlineData("IF(A1>B1,C1,D1)", "A1,B1,C1,D1")]
    [InlineData("Data!A5+1", "A5")]
    [InlineData("data!A5", "A5")]
    [InlineData("'Data'!A5", "A5")]
    [InlineData("SUM(Data!A1:A9)", "A1:A9")]
    public void Finds_references_to_the_formulas_own_sheet_or_a_qualified_one(string formula, string expected) =>
        Assert.Equal(expected.Split(','), Refs(formula));

    [Theory]
    [InlineData("Other!A5")]
    [InlineData("'Other Sheet'!A5:B6")]
    [InlineData("\"A5\"")]
    [InlineData("\"say \"\"A5\"\" now\"&\"x\"")]
    [InlineData("LOG10(100)")]
    [InlineData("ATAN2(1,2)")]
    [InlineData("Table1[Col]")]
    [InlineData("T1[[#This Row],[Col]]")]
    [InlineData("MyData!A5")]
    [InlineData("TRUE")]
    [InlineData("1+2")]
    [InlineData("1.5E3")]
    public void Ignores_other_sheets_text_literals_functions_and_names(string formula) =>
        Assert.Empty(Refs(formula));

    [Fact]
    public void An_unqualified_reference_belongs_to_the_formulas_sheet_only()
    {
        Assert.Empty(Refs("A5", formulaSheet: "Report", target: "Data"));
        Assert.Equal(["A5"], Refs("A5", formulaSheet: "Report", target: "Report"));
        Assert.Empty(FormulaReferences.Find("A5", formulaSheet: null, "Data")); // 已定義名稱裡的參照一定有前綴
    }

    [Fact]
    public void Quoted_sheet_names_with_spaces_and_apostrophes_are_matched()
    {
        Assert.Equal(["A5"], Refs("'My Data'!A5", target: "My Data"));
        Assert.Equal(["A5"], Refs("'It''s'!A5", target: "It's"));
        Assert.Equal(["A5"], Refs("銷售!A5", target: "銷售"));
    }

    [Fact]
    public void Mixed_formulas_report_only_the_target_sheets_references()
    {
        Assert.Equal(["A1:A3"], Refs("SUM(Data!A1:A3)+Other!B2+'Third'!C3", formulaSheet: "Report"));
    }
}
