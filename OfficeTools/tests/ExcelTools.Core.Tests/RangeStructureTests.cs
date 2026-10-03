using ClosedXML.Excel;
using ExcelTools.Core.Models;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Tests;

public sealed class RangeStructureTests : IDisposable
{
    private readonly ExcelEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    /// <summary>
    /// Data：A1:A9 = 1..9、A10 =SUM(A1:A9)、B1 =A5*2、C1 =$A$5+A$5+$A5、合併 D3:E4、已定義名稱 Five = Data!$A$5；
    /// Report：A1 =Data!A5+1、A2 =SUM(Data!A1:A9)。
    /// </summary>
    private string OpenFixture()
    {
        var path = _env.MakeBook("book.xlsx", wb =>
        {
            var d = wb.AddWorksheet("Data");
            for (var r = 1; r <= 9; r++)
            {
                d.Cell(r, 1).Value = r;
            }

            d.Cell("A10").FormulaA1 = "SUM(A1:A9)";
            d.Cell("B1").FormulaA1 = "A5*2";
            d.Cell("C1").FormulaA1 = "$A$5+A$5+$A5";
            d.Range("D3:E4").Merge();
            wb.DefinedNames.Add("Five", "Data!$A$5");
            var report = wb.AddWorksheet("Report");
            report.Cell("A1").FormulaA1 = "Data!A5+1";
            report.Cell("A2").FormulaA1 = "SUM(Data!A1:A9)";
        });
        return _env.Files.Open(path).WorkbookId;
    }

    private string Formula(string id, string sheet, string address) =>
        _env.Sessions.Use(id, false, s => s.Workbook.Worksheet(sheet).Cell(address).FormulaA1);

    private string[] Merged(string id) =>
        _env.Sessions.Use(id, false, s => s.Workbook.Worksheet("Data").MergedRanges.Select(m => m.RangeAddress.ToStringRelative(false)).ToArray());

    private string NameRefersTo(string id) =>
        _env.Sessions.Use(id, false, s => s.Workbook.DefinedNames.Single(n => n.Name == "Five").RefersTo);

    private object?[] Column(string id, string range) =>
        _env.Ranges.ReadRange(id, "Data", range).Values.Select(r => r[0]).ToArray();

    private bool IsDirty() => _env.Files.ListOpen().Single().IsDirty;

    // ---- InsertRows ----

    [Fact]
    public void InsertRows_shifts_content_down_and_adjusts_every_reference()
    {
        var id = OpenFixture();

        var result = _env.Ranges.InsertRows(id, "Data", 5, 2);

        Assert.Equal("Data", result.Sheet);
        Assert.Equal("5:6", result.Range);
        Assert.Equal("A1:C12", result.UsedRange); // 合併的 D3:E4 沒有內容，不算在有內容的範圍內
        Assert.Empty(result.Warnings);
        Assert.Equal(new object?[] { 4.0, null, null, 5.0, 6.0 }, Column(id, "A4:A8"));
        Assert.Equal("SUM(A1:A11)", Formula(id, "Data", "A12"));
        Assert.Equal("A7*2", Formula(id, "Data", "B1"));
        Assert.Equal("$A$7+A$7+$A7", Formula(id, "Data", "C1"));
        Assert.Equal("Data!A7+1", Formula(id, "Report", "A1"));
        Assert.Equal("SUM(Data!A1:A11)", Formula(id, "Report", "A2"));
        Assert.Contains("$A$7", NameRefersTo(id), StringComparison.Ordinal);
        Assert.Equal(["D3:E4"], Merged(id)); // 在合併區上方插入，不受影響
        Assert.True(IsDirty());
    }

    [Fact]
    public void InsertRows_above_a_merged_range_moves_it()
    {
        var id = OpenFixture();
        _env.Ranges.InsertRows(id, "Data", 1);
        Assert.Equal(["D4:E5"], Merged(id));
    }

    [Fact]
    public void InsertRows_values_and_computed_results_stay_consistent()
    {
        var id = OpenFixture();
        _env.Ranges.InsertRows(id, "Data", 5, 2);

        Assert.Equal(10.0, _env.Ranges.ReadRange(id, "Data", "B1").Values[0][0]); // 仍然是原本的 5 * 2
        Assert.Equal(45.0, _env.Ranges.ReadRange(id, "Data", "A12").Values[0][0]);
        Assert.Equal(6.0, _env.Ranges.ReadRange(id, "Report", "A1").Values[0][0]);
    }

    [Fact]
    public void InsertRows_below_the_data_is_harmless()
    {
        var id = OpenFixture();
        var result = _env.Ranges.InsertRows(id, "Data", 500, 3);

        Assert.Equal("A1:C10", result.UsedRange);
        Assert.Equal("SUM(A1:A9)", Formula(id, "Data", "A10"));
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(-3, 1)]
    [InlineData(1048577, 1)]
    [InlineData(5, 0)]
    [InlineData(5, -1)]
    [InlineData(1048576, 2)]
    public void InsertRows_with_invalid_arguments_is_rejected_and_changes_nothing(int row, int count)
    {
        var id = OpenFixture();
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.InsertRows(id, "Data", row, count)).Code);
        Assert.False(IsDirty());
        Assert.Equal(5.0, Column(id, "A5:A5")[0]);
    }

    [Fact]
    public void InsertRows_that_would_push_content_off_the_sheet_is_refused()
    {
        var path = _env.MakeBook("edge.xlsx", wb =>
        {
            var d = wb.AddWorksheet("Data");
            d.Cell("A1").Value = "top";
            d.Cell("A1048570").Value = "bottom";
        });
        var id = _env.Files.Open(path).WorkbookId;

        var ex = Throws(() => _env.Ranges.InsertRows(id, "Data", 2, 7));

        Assert.Equal(ErrorCodes.InvalidRange, ex.Code);
        Assert.Contains("擠出", ex.Message, StringComparison.Ordinal);
        Assert.Equal("bottom", _env.Ranges.ReadRange(id, "Data", "A1048570").Values[0][0]);
        Assert.False(IsDirty());

        _env.Ranges.InsertRows(id, "Data", 2, 6); // 剛好放得下
        Assert.Equal("bottom", _env.Ranges.ReadRange(id, "Data", "A1048576").Values[0][0]);
    }

    // ---- InsertColumns / DeleteColumns ----

    [Fact]
    public void InsertColumns_shifts_content_right_and_adjusts_references()
    {
        var id = OpenFixture();

        var result = _env.Ranges.InsertColumns(id, "Data", "A");

        Assert.Equal("A:A", result.Range);
        Assert.Equal(5.0, _env.Ranges.ReadRange(id, "Data", "B5").Values[0][0]);
        Assert.Equal("B5*2", Formula(id, "Data", "C1"));
        Assert.Equal("Data!B5+1", Formula(id, "Report", "A1"));
        Assert.Equal("SUM(Data!B1:B9)", Formula(id, "Report", "A2"));
        Assert.Equal(["E3:F4"], Merged(id));
        Assert.Contains("$B$5", NameRefersTo(id), StringComparison.Ordinal);
    }

    [Fact]
    public void InsertColumns_with_a_count_and_lowercase_dollar_column()
    {
        var id = OpenFixture();
        Assert.Equal("C:E", _env.Ranges.InsertColumns(id, "Data", "$c", 3).Range);
    }

    [Theory]
    [InlineData("")]
    [InlineData("A1")]
    [InlineData("1")]
    [InlineData("XFE")]
    [InlineData("A:A")]
    public void Column_operations_reject_invalid_column_names(string column)
    {
        var id = OpenFixture();
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.InsertColumns(id, "Data", column)).Code);
        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.DeleteColumns(id, "Data", column)).Code);
    }

    [Fact]
    public void InsertColumns_that_would_push_content_off_the_sheet_is_refused()
    {
        var path = _env.MakeBook("wide.xlsx", wb => wb.AddWorksheet("Data").Cell("XFD1").Value = "edge");
        var id = _env.Files.Open(path).WorkbookId;

        Assert.Equal(ErrorCodes.InvalidRange, Throws(() => _env.Ranges.InsertColumns(id, "Data", "A")).Code);
        Assert.Equal("edge", _env.Ranges.ReadRange(id, "Data", "XFD1").Values[0][0]);
    }

    // ---- DeleteRows ----

    [Fact]
    public void DeleteRows_shifts_content_up_and_shrinks_ranges()
    {
        var id = OpenFixture();

        var result = _env.Ranges.DeleteRows(id, "Data", 2, 2);

        Assert.Equal("2:3", result.Range);
        Assert.Equal(new object?[] { 1.0, 4.0, 5.0 }, Column(id, "A1:A3"));
        Assert.Equal("SUM(A1:A7)", Formula(id, "Data", "A8"));
        Assert.Equal("A3*2", Formula(id, "Data", "B1")); // 原本指向第 5 列，現在是第 3 列
        Assert.Equal("SUM(Data!A1:A7)", Formula(id, "Report", "A2"));
        Assert.Empty(result.Warnings); // 沒有任何參照整個落在被刪的列內
    }

    [Fact]
    public void DeleteRows_warns_about_references_to_the_deleted_cells_because_closedxml_silently_repoints_them()
    {
        var id = OpenFixture();

        var result = _env.Ranges.DeleteRows(id, "Data", 5);

        var warning = Assert.Single(result.Warnings);
        Assert.Contains("4 處", warning, StringComparison.Ordinal);
        Assert.Contains("Data!B1", warning, StringComparison.Ordinal);
        Assert.Contains("Data!C1", warning, StringComparison.Ordinal);
        Assert.Contains("Report!A1", warning, StringComparison.Ordinal);
        Assert.Contains("已定義名稱 Five", warning, StringComparison.Ordinal);
        Assert.Contains("#REF!", warning, StringComparison.Ordinal);
        Assert.DoesNotContain("Report!A2", warning, StringComparison.Ordinal); // SUM(A1:A9) 只是縮小，不是失效
        Assert.DoesNotContain("Data!A10", warning, StringComparison.Ordinal);
        Assert.Equal("SUM(Data!A1:A8)", Formula(id, "Report", "A2"));
    }

    [Fact]
    public void DeleteRows_does_not_count_formulas_that_live_inside_the_deleted_rows()
    {
        var id = OpenFixture();
        var result = _env.Ranges.DeleteRows(id, "Data", 10); // A10 的公式本身被刪除

        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void DeleteRows_warns_when_a_whole_range_reference_falls_inside_the_deleted_rows()
    {
        var path = _env.MakeBook("range.xlsx", wb =>
        {
            var d = wb.AddWorksheet("Data");
            for (var r = 1; r <= 8; r++)
            {
                d.Cell(r, 1).Value = r;
            }

            d.Cell("C1").FormulaA1 = "SUM(A5:A6)";
            d.Cell("C2").FormulaA1 = "SUM(A5:A8)";
        });
        var id = _env.Files.Open(path).WorkbookId;

        var warning = Assert.Single(_env.Ranges.DeleteRows(id, "Data", 5, 2).Warnings);

        Assert.Contains("1 處", warning, StringComparison.Ordinal);
        Assert.Contains("Data!C1", warning, StringComparison.Ordinal);
        Assert.DoesNotContain("Data!C2", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void DeleteRows_warning_recognizes_quoted_sheet_names_and_ignores_text_that_only_looks_like_a_reference()
    {
        var path = _env.MakeBook("quoted.xlsx", wb =>
        {
            var d = wb.AddWorksheet("My Data");
            d.Cell("A1").Value = 1;
            d.Cell("A5").Value = 5;
            var o = wb.AddWorksheet("Other");
            o.Cell("A1").FormulaA1 = "'My Data'!A5";            // 要警告
            o.Cell("A2").FormulaA1 = "\"'My Data'!A5\"&\"x\"";  // 字串內的文字，不是參照
            o.Cell("A3").FormulaA1 = "A5";                      // Other 自己的 A5
            o.Cell("A4").FormulaA1 = "LOG10(100)";              // 函式名稱長得像儲存格
        });
        var id = _env.Files.Open(path).WorkbookId;

        var warning = Assert.Single(_env.Ranges.DeleteRows(id, "My Data", 5).Warnings);

        Assert.Contains("1 處", warning, StringComparison.Ordinal);
        Assert.Contains("Other!A1", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void DeleteRows_beyond_the_used_area_changes_nothing_visible()
    {
        var id = OpenFixture();
        var result = _env.Ranges.DeleteRows(id, "Data", 500, 10);

        Assert.Equal("A1:C10", result.UsedRange);
        Assert.Empty(result.Warnings);
    }

    // ---- DeleteColumns ----

    [Fact]
    public void DeleteColumns_shifts_left_and_warns_about_references_to_the_deleted_column()
    {
        var id = OpenFixture();

        var result = _env.Ranges.DeleteColumns(id, "Data", "A");

        Assert.Equal("A:A", result.Range);
        var warning = Assert.Single(result.Warnings);
        Assert.Contains("5 處", warning, StringComparison.Ordinal);
        Assert.Contains("Data!B1", warning, StringComparison.Ordinal);
        Assert.Contains("Report!A2", warning, StringComparison.Ordinal);
        Assert.Contains("已定義名稱 Five", warning, StringComparison.Ordinal);
        Assert.Equal(["C3:D4"], Merged(id));
    }

    [Fact]
    public void DeleteColumns_that_nothing_references_gives_no_warning()
    {
        var id = OpenFixture();
        Assert.Empty(_env.Ranges.DeleteColumns(id, "Data", "D", 2).Warnings);
    }

    // ---- 表格、資料驗證與復原 ----

    private string OpenTable()
    {
        var path = _env.MakeBook("table.xlsx", wb =>
        {
            var d = wb.AddWorksheet("Data");
            d.Cell("A1").Value = "h1";
            d.Cell("B1").Value = "h2";
            for (var r = 2; r <= 6; r++)
            {
                d.Cell(r, 1).Value = r;
                d.Cell(r, 2).Value = r * 10;
            }

            d.Range("A1:B6").CreateTable("T1");
        });
        return _env.Files.Open(path).WorkbookId;
    }

    private string TableRange(string id) =>
        _env.Sessions.Use(id, false, s => s.Workbook.Worksheet("Data").Tables.Single().RangeAddress.ToStringRelative(false));

    [Fact]
    public void Inserting_and_deleting_rows_inside_a_table_resizes_it_and_the_file_stays_valid()
    {
        var id = OpenTable();
        var path = _env.Files.ListOpen().Single().Path;

        _env.Ranges.InsertRows(id, "Data", 4);
        Assert.Equal("A1:B7", TableRange(id));
        _env.Ranges.DeleteRows(id, "Data", 3);
        Assert.Equal("A1:B6", TableRange(id));
        _env.Files.Save(id);

        Assert.Empty(ExcelEnv.ValidateXlsx(path));
    }

    [Fact]
    public void Deleting_every_row_of_a_table_is_rolled_back_with_UNSAFE_OPERATION()
    {
        var id = OpenTable();

        var ex = Throws(() => _env.Ranges.DeleteRows(id, "Data", 1, 6));

        Assert.Equal(ErrorCodes.UnsafeOperation, ex.Code);
        Assert.Contains("已復原", ex.Message, StringComparison.Ordinal);
        Assert.Equal("A1:B6", TableRange(id));
        Assert.Equal(3.0, _env.Ranges.ReadRange(id, "Data", "A3").Values[0][0]);
        Assert.False(IsDirty()); // 復原後不是有變更
    }

    [Fact]
    public void After_a_rolled_back_operation_the_workbook_keeps_working_and_can_be_saved()
    {
        var id = OpenTable();
        var path = _env.Files.ListOpen().Single().Path;
        Throws(() => _env.Ranges.DeleteRows(id, "Data", 1, 6));

        _env.Ranges.WriteRange(id, "Data", "D1", [["after rollback"]]);
        _env.Ranges.InsertRows(id, "Data", 3);
        _env.Files.Save(id);

        Assert.Empty(ExcelEnv.ValidateXlsx(path));
        Assert.Equal("after rollback", ExcelEnv.ReadFromDisk(path, "D1"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Inserting_or_deleting_the_first_row_on_a_sheet_with_data_validation_never_leaves_a_corrupt_session(bool insert)
    {
        // 這組資料在「剛用 API 在記憶體中建立」的活頁簿上會讓 ClosedXML 產生空的 sqref（壞檔）；
        // 但 session 裡的活頁簿一律是從檔案載入的，實測不會發生。這個測試確保無論結果是成功或被復原，session 都不會壞掉。
        var path = _env.MakeBook("dv.xlsx", wb =>
        {
            var d = wb.AddWorksheet("Data");
            d.Cell("A1").Value = "h1";
            d.Cell("B1").Value = "h2";
            for (var r = 2; r <= 6; r++)
            {
                d.Cell(r, 1).Value = r;
                d.Cell(r, 2).Value = r * 10;
            }

            d.Cell("D1").Value = "x";
            d.Cell("D2").Value = 5;
            d.Cell("E1").Value = "e";
            d.Range("E2:E6").CreateDataValidation().WholeNumber.Between(1, 10);
        });
        var id = _env.Files.Open(path).WorkbookId;

        try
        {
            _ = insert ? _env.Ranges.InsertRows(id, "Data", 1) : _env.Ranges.DeleteRows(id, "Data", 1);
        }
        catch (OfficeToolException ex)
        {
            Assert.Equal(ErrorCodes.UnsafeOperation, ex.Code);
            Assert.False(IsDirty());
        }

        _env.Files.Save(id);
        Assert.Empty(ExcelEnv.ValidateXlsx(path));
    }

    [Fact]
    public void Structure_changes_survive_save_and_reopen_and_the_file_is_valid()
    {
        var id = OpenFixture();
        var path = _env.Files.ListOpen().Single().Path;
        _env.Ranges.InsertRows(id, "Data", 5, 2);
        _env.Ranges.InsertColumns(id, "Data", "B");
        _env.Ranges.DeleteRows(id, "Data", 1);
        _env.Files.Save(id);
        _env.Files.Close(id);

        var reopened = _env.Files.Open(path).WorkbookId;

        Assert.Empty(ExcelEnv.ValidateXlsx(path));
        Assert.Equal("SUM(A1:A10)", Formula(reopened, "Data", "A11"));
    }

    // ---- 共通錯誤 ----

    [Fact]
    public void Structure_operations_report_errors_with_the_right_codes()
    {
        var id = OpenFixture();
        Assert.Equal(ErrorCodes.SheetNotFound, Throws(() => _env.Ranges.InsertRows(id, "Nope", 1)).Code);
        Assert.Equal(ErrorCodes.SheetNotFound, Throws(() => _env.Ranges.DeleteColumns(id, "Nope", "A")).Code);
        Assert.Equal(ErrorCodes.SessionNotFound, Throws(() => _env.Ranges.DeleteRows("wb_nope", "Data", 1)).Code);
        Assert.False(IsDirty());
    }

    [Fact]
    public void Structure_operations_refuse_read_only_workbooks()
    {
        File.Move(_env.MakeWorkbook("book.xlsx"), _env.Path("macro.xlsm"));
        var id = _env.Files.Open(_env.Path("macro.xlsm")).WorkbookId;

        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Ranges.InsertRows(id, "Data", 1)).Code);
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Ranges.DeleteRows(id, "Data", 1)).Code);
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Ranges.InsertColumns(id, "Data", "A")).Code);
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Ranges.DeleteColumns(id, "Data", "A")).Code);
    }
}
