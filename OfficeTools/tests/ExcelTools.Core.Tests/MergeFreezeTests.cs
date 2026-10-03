using ClosedXML.Excel;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Tests;

public sealed class MergeTests : IDisposable
{
    private readonly ExcelEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    private string Open(Action<IXLWorksheet>? fill = null) =>
        _env.Files.Open(_env.MakeWorkbook($"{Guid.NewGuid():N}.xlsx", fill ?? (ws => { ws.Cell("A1").Value = "title"; }))).WorkbookId;

    private string[] Merged(string id) => _env.Sheets.GetSheetInfo(id, "Data").MergedRanges.Order().ToArray();

    // ---- Merge ----

    [Fact]
    public void Merges_a_range_and_keeps_the_top_left_content()
    {
        var id = Open();

        var result = _env.Formats.Merge(id, "Data", "A1:C1");

        Assert.Equal("Data", result.Sheet);
        Assert.Equal(["A1:C1"], result.Ranges);
        Assert.Empty(result.Warnings);
        Assert.Equal(["A1:C1"], Merged(id));
        Assert.Equal("title", _env.Ranges.ReadRange(id, "Data", "A1").Values[0][0]);
        Assert.True(_env.Files.ListOpen().Single().IsDirty);
    }

    [Theory]
    [InlineData("c3:a1", "A1:C3")]
    [InlineData("$A$1:$C$3", "A1:C3")]
    public void Accepts_the_tolerated_address_forms(string input, string expected)
    {
        var id = Open();
        Assert.Equal([expected], _env.Formats.Merge(id, "Data", input).Ranges);
    }

    [Fact]
    public void Merging_over_other_cells_with_content_is_refused_naming_them()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "keep";
            ws.Cell("B1").Value = "would be lost";
            ws.Cell("A2").Value = "also lost";
        });

        var ex = Throws(() => _env.Formats.Merge(id, "Data", "A1:C2"));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("2 個", ex.Message, StringComparison.Ordinal);
        Assert.Contains("A2", ex.Message, StringComparison.Ordinal);
        Assert.Contains("B1", ex.Message, StringComparison.Ordinal);
        Assert.Contains("discardOtherValues", ex.Hint, StringComparison.Ordinal);
        Assert.Empty(Merged(id));
        Assert.Equal("would be lost", _env.Ranges.ReadRange(id, "Data", "B1").Values[0][0]); // 沒被動到
        Assert.False(_env.Files.ListOpen().Single().IsDirty);
    }

    [Fact]
    public void With_discardOtherValues_the_merge_goes_ahead_and_reports_what_was_cleared()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "keep";
            ws.Cell("B1").Value = "lost";
            ws.Cell("A2").Value = 5;
        });

        var result = _env.Formats.Merge(id, "Data", "A1:B2", discardOtherValues: true);

        var warning = Assert.Single(result.Warnings);
        Assert.Contains("2 個", warning, StringComparison.Ordinal);
        Assert.Equal(["A1:B2"], Merged(id));
        Assert.Equal("keep", _env.Ranges.ReadRange(id, "Data", "A1").Values[0][0]);
        Assert.Equal("A1:A1", _env.Sheets.GetSheetInfo(id, "Data").UsedRange); // 只剩左上角有內容
    }

    [Fact]
    public void Many_discarded_cells_are_summarized_in_the_refusal()
    {
        var id = Open(ws =>
        {
            for (var r = 1; r <= 10; r++)
            {
                ws.Cell(r, 1).Value = r;
            }
        });

        var ex = Throws(() => _env.Formats.Merge(id, "Data", "A1:A10"));

        Assert.Contains("9 個", ex.Message, StringComparison.Ordinal);
        Assert.Contains("等共 9 格", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Blank_cells_and_formatting_do_not_count_as_content()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "t";
            ws.Cell("B1").Style.Fill.BackgroundColor = XLColor.Yellow;
        });

        Assert.Empty(_env.Formats.Merge(id, "Data", "A1:C1").Warnings);
    }

    [Fact]
    public void Merging_the_same_range_again_is_a_harmless_no_op()
    {
        var id = Open();
        _env.Formats.Merge(id, "Data", "A1:C1");

        var again = _env.Formats.Merge(id, "Data", "A1:C1");

        Assert.Equal(["A1:C1"], again.Ranges);
        Assert.Empty(again.Warnings);
        Assert.Equal(["A1:C1"], Merged(id));
    }

    [Fact]
    public void Merging_over_existing_merges_replaces_them_and_says_so()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "t";
            ws.Range("A1:B2").Merge();
            ws.Range("D1:E1").Merge();
        });

        var result = _env.Formats.Merge(id, "Data", "A1:C3");

        var warning = Assert.Single(result.Warnings);
        Assert.Contains("A1:B2", warning, StringComparison.Ordinal);
        Assert.DoesNotContain("D1:E1", warning, StringComparison.Ordinal);
        Assert.Equal(["A1:C3", "D1:E1"], Merged(id));
    }

    [Fact]
    public void Partially_overlapping_an_existing_merge_also_replaces_it()
    {
        var id = Open(ws => ws.Range("A1:B2").Merge());

        var result = _env.Formats.Merge(id, "Data", "B2:C3");

        Assert.Contains("A1:B2", Assert.Single(result.Warnings), StringComparison.Ordinal);
        Assert.Equal(["B2:C3"], Merged(id));
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("A1:A1")]
    public void A_single_cell_cannot_be_merged(string range) =>
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Formats.Merge(Open(), "Data", range)).Code);

    [Theory]
    [InlineData("A:C")]
    [InlineData("1:3")]
    public void Whole_columns_and_rows_cannot_be_merged(string range) =>
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Formats.Merge(Open(), "Data", range)).Code);

    [Fact]
    public void A_range_over_the_limit_is_rejected()
    {
        using var env = new ExcelEnv(maxCellsPerFormat: 10);
        var id = env.Files.Open(env.MakeWorkbook("big.xlsx", ws => ws.Cell("A1").Value = "t")).WorkbookId;

        Assert.Equal(ErrorCodes.InvalidRange, Assert.Throws<OfficeToolException>(() => env.Formats.Merge(id, "Data", "A1:C4")).Code);
        env.Formats.Merge(id, "Data", "A1:B5"); // 剛好 10 格
    }

    [Theory]
    [InlineData("A2:A3")]   // 完全在表格內
    [InlineData("B1:C1")]   // 部分重疊
    [InlineData("A1:C5")]   // 包含整個表格
    public void Merging_into_a_table_is_refused(string range)
    {
        var path = _env.MakeBook("table.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "h1";
            ws.Cell("B1").Value = "h2";
            ws.Cell("A2").Value = 1;
            ws.Cell("B2").Value = 2;
            ws.Range("A1:B2").CreateTable("T1");
        });
        var id = _env.Files.Open(path).WorkbookId;

        var ex = Throws(() => _env.Formats.Merge(id, "Data", range));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("表格", ex.Message, StringComparison.Ordinal);
        Assert.Empty(Merged(id));
    }

    [Fact]
    public void Merging_next_to_a_table_is_fine()
    {
        var path = _env.MakeBook("table.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "h";
            ws.Cell("A2").Value = 1;
            ws.Range("A1:A2").CreateTable("T1");
        });
        var id = _env.Files.Open(path).WorkbookId;

        _env.Formats.Merge(id, "Data", "C1:E1");

        Assert.Equal(["C1:E1"], Merged(id));
    }

    // ---- Unmerge ----

    [Fact]
    public void Unmerge_with_the_exact_range_keeps_the_top_left_content()
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "title";
            ws.Range("A1:C1").Merge();
        });

        var result = _env.Formats.Unmerge(id, "Data", "A1:C1");

        Assert.Equal(["A1:C1"], result.Ranges);
        Assert.Empty(Merged(id));
        Assert.Equal("title", _env.Ranges.ReadRange(id, "Data", "A1").Values[0][0]);
    }

    [Theory]
    [InlineData("B1")]        // 只碰到合併區內的一格
    [InlineData("A1:A1")]
    [InlineData("C1:Z9")]     // 只碰到邊緣
    [InlineData("A1:XFD10")]  // 範圍比合併區大
    [InlineData("A:A")]
    [InlineData("1:1")]
    public void Unmerge_cancels_the_whole_merge_when_the_range_touches_any_of_its_cells(string range)
    {
        var id = Open(ws =>
        {
            ws.Cell("A1").Value = "title";
            ws.Range("A1:C2").Merge();
        });

        var result = _env.Formats.Unmerge(id, "Data", range);

        Assert.Equal(["A1:C2"], result.Ranges);
        Assert.Empty(Merged(id));
    }

    [Fact]
    public void Unmerge_only_touches_merges_that_intersect_the_range()
    {
        var id = Open(ws =>
        {
            ws.Range("A1:B1").Merge();
            ws.Range("A3:B3").Merge();
            ws.Range("D1:E1").Merge();
        });

        var result = _env.Formats.Unmerge(id, "Data", "A2:B4");

        Assert.Equal(["A3:B3"], result.Ranges);
        Assert.Equal(["A1:B1", "D1:E1"], Merged(id));
    }

    [Fact]
    public void Unmerge_can_cancel_several_merges_at_once()
    {
        var id = Open(ws =>
        {
            ws.Range("A1:B1").Merge();
            ws.Range("A3:B3").Merge();
        });

        var result = _env.Formats.Unmerge(id, "Data", "A1:B3");

        Assert.Equal(["A1:B1", "A3:B3"], result.Ranges.Order());
        Assert.Empty(Merged(id));
    }

    [Fact]
    public void Unmerge_without_any_merge_returns_an_empty_list_and_does_not_fail()
    {
        var id = Open();
        Assert.Empty(_env.Formats.Unmerge(id, "Data", "A1:C3").Ranges);
        Assert.Empty(Merged(id));
    }

    [Fact]
    public void Merge_and_unmerge_round_trip_through_save_and_the_file_is_valid_open_xml()
    {
        var id = Open(ws => ws.Cell("A1").Value = "title");
        var path = _env.Files.ListOpen().Single().Path;
        _env.Formats.Merge(id, "Data", "A1:D1");
        _env.Formats.Merge(id, "Data", "A3:B4");
        _env.Formats.Unmerge(id, "Data", "A3");
        _env.Files.Save(id);
        _env.Files.Close(id);

        var reopened = _env.Files.Open(path).WorkbookId;

        Assert.Empty(ExcelEnv.ValidateXlsx(path));
        Assert.Equal(["A1:D1"], Merged(reopened));
    }

    [Fact]
    public void Errors_use_the_right_codes_and_read_only_workbooks_refuse_merging()
    {
        var id = Open();
        Assert.Equal(ErrorCodes.SheetNotFound, Throws(() => _env.Formats.Merge(id, "Nope", "A1:B1")).Code);
        Assert.Equal(ErrorCodes.SheetNotFound, Throws(() => _env.Formats.Unmerge(id, "Nope", "A1")).Code);
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Formats.Merge(id, "Data", "zz")).Code);
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Formats.Unmerge(id, "Data", "")).Code);
        Assert.Equal(ErrorCodes.SessionNotFound, Throws(() => _env.Formats.Merge("wb_nope", "Data", "A1:B1")).Code);

        File.Move(_env.MakeWorkbook("ro.xlsx"), _env.Path("macro.xlsm"));
        var ro = _env.Files.Open(_env.Path("macro.xlsm")).WorkbookId;
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Formats.Merge(ro, "Data", "A1:B1")).Code);
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Formats.Unmerge(ro, "Data", "A1")).Code);
    }
}

public sealed class FreezePanesTests : IDisposable
{
    private readonly ExcelEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    private string Open() => _env.Files.Open(_env.MakeWorkbook($"{Guid.NewGuid():N}.xlsx")).WorkbookId;

    private (int Rows, int Columns) Frozen(string id)
    {
        var info = _env.Sheets.GetSheetInfo(id, "Data");
        return (info.FrozenRows, info.FrozenColumns);
    }

    [Theory]
    [InlineData(1, 0)]
    [InlineData(0, 2)]
    [InlineData(3, 4)]
    [InlineData(1000, 100)]
    public void Freezes_the_requested_rows_and_columns(int rows, int columns)
    {
        var id = Open();

        var result = _env.Formats.FreezePanes(id, "Data", rows, columns);

        Assert.Equal("Data", result.Sheet);
        Assert.Equal((rows, columns), (result.FrozenRows, result.FrozenColumns));
        Assert.Equal((rows, columns), Frozen(id));
        Assert.True(_env.Files.ListOpen().Single().IsDirty);
    }

    [Fact]
    public void Each_call_sets_both_directions_exactly_and_zero_zero_unfreezes()
    {
        var id = Open();

        _env.Formats.FreezePanes(id, "Data", 2, 3);
        _env.Formats.FreezePanes(id, "Data", 1, 0); // 欄數 0 要真的取消欄的凍結（不是維持舊值）
        Assert.Equal((1, 0), Frozen(id));

        _env.Formats.FreezePanes(id, "Data", 0, 2);
        Assert.Equal((0, 2), Frozen(id));

        _env.Formats.FreezePanes(id, "Data", 0, 0);
        Assert.Equal((0, 0), Frozen(id));
    }

    [Theory]
    [InlineData(-1, 0)]
    [InlineData(0, -1)]
    [InlineData(1001, 0)]
    [InlineData(0, 101)]
    public void Out_of_range_values_are_rejected_and_change_nothing(int rows, int columns)
    {
        var id = Open();
        _env.Formats.FreezePanes(id, "Data", 1, 1);

        var ex = Throws(() => _env.Formats.FreezePanes(id, "Data", rows, columns));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("取消凍結", ex.Hint, StringComparison.Ordinal);
        Assert.Equal((1, 1), Frozen(id));
    }

    [Fact]
    public void Frozen_panes_survive_save_and_the_file_is_valid_open_xml()
    {
        var id = Open();
        var path = _env.Files.ListOpen().Single().Path;
        _env.Formats.FreezePanes(id, "Data", 1, 2);
        _env.Files.Save(id);
        _env.Files.Close(id);

        var reopened = _env.Files.Open(path).WorkbookId;

        Assert.Empty(ExcelEnv.ValidateXlsx(path));
        Assert.Equal((1, 2), Frozen(reopened));
    }

    [Fact]
    public void Unfreezing_leaves_a_valid_file()
    {
        var id = Open();
        var path = _env.Files.ListOpen().Single().Path;
        _env.Formats.FreezePanes(id, "Data", 2, 2);
        _env.Formats.FreezePanes(id, "Data", 0, 0);
        _env.Files.Save(id);

        Assert.Empty(ExcelEnv.ValidateXlsx(path));
    }

    [Fact]
    public void Errors_use_the_right_codes_and_read_only_workbooks_refuse_freezing()
    {
        var id = Open();
        Assert.Equal(ErrorCodes.SheetNotFound, Throws(() => _env.Formats.FreezePanes(id, "Nope", 1, 0)).Code);
        Assert.Equal(ErrorCodes.SessionNotFound, Throws(() => _env.Formats.FreezePanes("wb_nope", "Data", 1, 0)).Code);

        File.Move(_env.MakeWorkbook("ro.xlsx"), _env.Path("macro.xlsm"));
        var ro = _env.Files.Open(_env.Path("macro.xlsm")).WorkbookId;
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Formats.FreezePanes(ro, "Data", 1, 0)).Code);
    }
}
