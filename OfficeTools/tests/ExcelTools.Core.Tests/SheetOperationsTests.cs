using ClosedXML.Excel;
using ExcelTools.Core.Operations;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Tests;

public sealed class SheetOperationsTests : IDisposable
{
    private readonly ExcelEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    /// <summary>Data（A1=5, A2=A1*2）、Report（引用 Data）、Hidden（隱藏）三張。</summary>
    private string OpenReportBook(string file = "book.xlsx")
    {
        var path = _env.MakeBook(file, wb =>
        {
            var data = wb.AddWorksheet("Data");
            data.Cell("A1").Value = 5;
            data.Cell("A2").FormulaA1 = "A1*2";
            var report = wb.AddWorksheet("Report");
            report.Cell("A1").FormulaA1 = "Data!A1+1";
            report.Cell("A2").FormulaA1 = "'Data'!A2";
            wb.AddWorksheet("Hidden").Visibility = XLWorksheetVisibility.Hidden;
        });
        return _env.Files.Open(path).WorkbookId;
    }

    private IReadOnlyList<string> Names(string id) => _env.Sheets.ListSheets(id).Select(s => s.Name).ToList();

    private object? Formula(string id, string sheet, string address) =>
        _env.Sessions.Use(id, false, s => s.Workbook.Worksheet(sheet).Cell(address).FormulaA1);

    // ---- ListSheets / GetSheetInfo ----

    [Fact]
    public void ListSheets_returns_tab_order_with_used_range_and_hidden_flag()
    {
        var id = OpenReportBook();
        var sheets = _env.Sheets.ListSheets(id);

        Assert.Equal(["Data", "Report", "Hidden"], sheets.Select(s => s.Name));
        Assert.Equal([1, 2, 3], sheets.Select(s => s.Index));
        Assert.Equal("A1:A2", sheets[0].UsedRange);
        Assert.Null(sheets[2].UsedRange);
        Assert.Equal([false, false, true], sheets.Select(s => s.Hidden));
    }

    [Fact]
    public void ListSheets_follows_tab_order_not_creation_order_after_a_move()
    {
        var id = OpenReportBook();
        _env.Sheets.Move(id, "Hidden", 1);

        Assert.Equal(["Hidden", "Data", "Report"], Names(id));
        Assert.Equal(["Hidden", "Data", "Report"], _env.Files.ListOpen().Count == 1 ? Names(id) : []);
    }

    [Fact]
    public void GetSheetInfo_describes_merges_freeze_tables_and_filter()
    {
        var path = _env.MakeBook("info.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "h1";
            ws.Cell("B1").Value = "h2";
            ws.Cell("A2").Value = 1;
            ws.Cell("B2").Value = 2;
            ws.Range("D1:E2").Merge();
            ws.SheetView.FreezeRows(1);
            ws.SheetView.FreezeColumns(2);
            ws.Range("A1:B2").CreateTable("Sales");
        });
        var id = _env.Files.Open(path).WorkbookId;

        var info = _env.Sheets.GetSheetInfo(id, "Data");

        Assert.Equal("Data", info.Name);
        Assert.Equal(1, info.Index);
        Assert.Equal("Visible", info.Visibility);
        Assert.Equal("A1:B2", info.UsedRange); // 已使用範圍以「有內容」為準，空的合併儲存格不算
        Assert.Equal(["D1:E2"], info.MergedRanges);
        Assert.Equal(1, info.MergedRangeCount);
        Assert.Equal(1, info.FrozenRows);
        Assert.Equal(2, info.FrozenColumns);
        Assert.Equal(["Sales"], info.TableNames);
    }

    [Fact]
    public void GetSheetInfo_reports_hidden_visibility_and_caps_the_merged_range_list()
    {
        var path = _env.MakeBook("merged.xlsx", wb =>
        {
            var ws = wb.AddWorksheet("M");
            for (var r = 1; r <= 60; r++)
            {
                ws.Range(r, 1, r, 2).Merge();
            }

            wb.AddWorksheet("H").Visibility = XLWorksheetVisibility.Hidden;
        });
        var id = _env.Files.Open(path).WorkbookId;

        var info = _env.Sheets.GetSheetInfo(id, "M");

        Assert.Equal(60, info.MergedRangeCount);
        Assert.Equal(SheetOperations.MaxMergedRangesListed, info.MergedRanges.Count);
        Assert.Equal("Hidden", _env.Sheets.GetSheetInfo(id, "H").Visibility);
    }

    [Theory]
    [InlineData("data")]
    [InlineData("DATA")]
    [InlineData("  Data  ")]
    public void Sheet_names_are_matched_case_insensitively_like_excel(string name)
    {
        var id = OpenReportBook();
        Assert.Equal("Data", _env.Sheets.GetSheetInfo(id, name).Name);
    }

    [Theory]
    [InlineData("Nope")]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void Unknown_sheet_returns_SHEET_NOT_FOUND_listing_the_existing_sheets_in_tab_order(string? name)
    {
        var id = OpenReportBook();
        _env.Sheets.Move(id, "Hidden", 1);

        var ex = Throws(() => _env.Sheets.GetSheetInfo(id, name!));

        Assert.Equal(ErrorCodes.SheetNotFound, ex.Code);
        Assert.Equal("現有的工作表：Hidden、Data、Report", ex.Hint);
    }

    [Fact]
    public void Every_operation_reports_SHEET_NOT_FOUND_for_unknown_sheets()
    {
        var id = OpenReportBook();
        foreach (var act in new Action[]
                 {
                     () => _env.Sheets.GetSheetInfo(id, "x"),
                     () => _env.Sheets.Rename(id, "x", "y"),
                     () => _env.Sheets.Copy(id, "x"),
                     () => _env.Sheets.Delete(id, "x"),
                     () => _env.Sheets.Move(id, "x", 1),
                 })
        {
            Assert.Equal(ErrorCodes.SheetNotFound, Throws(act).Code);
        }
    }

    // ---- Add ----

    [Fact]
    public void Add_appends_by_default_and_returns_the_new_order()
    {
        var id = OpenReportBook();

        var result = _env.Sheets.Add(id, "Summary");

        Assert.Equal(["Data", "Report", "Hidden", "Summary"], result.Sheets.Select(s => s.Name));
        Assert.Empty(result.Warnings);
        Assert.True(_env.Files.ListOpen().Single().IsDirty);
    }

    [Theory]
    [InlineData(1, new[] { "New", "Data", "Report", "Hidden" })]
    [InlineData(2, new[] { "Data", "New", "Report", "Hidden" })]
    [InlineData(4, new[] { "Data", "Report", "Hidden", "New" })]
    public void Add_at_a_position(int position, string[] expected)
    {
        var id = OpenReportBook();
        Assert.Equal(expected, _env.Sheets.Add(id, "New", position).Sheets.Select(s => s.Name));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(5)]
    [InlineData(-1)]
    public void Add_at_an_invalid_position_is_rejected_and_changes_nothing(int position)
    {
        var id = OpenReportBook();
        var ex = Throws(() => _env.Sheets.Add(id, "New", position));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("1 到 4", ex.Hint);
        Assert.Equal(["Data", "Report", "Hidden"], Names(id));
    }

    [Theory]
    [InlineData("Data")]
    [InlineData("data")]
    [InlineData(" DATA ")]
    public void Add_with_a_duplicate_name_is_rejected_case_insensitively(string name)
    {
        var id = OpenReportBook();
        var ex = Throws(() => _env.Sheets.Add(id, name));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("Data、Report、Hidden", ex.Hint);
    }

    [Theory]
    [InlineData("")]
    [InlineData("a/b")]
    [InlineData("a:b")]
    [InlineData("[x]")]
    [InlineData("'x'")]
    [InlineData("這個名稱超過三十一個字元的限制所以應該被拒絕才對喔喔喔喔喔喔喔喔")]
    public void Add_with_an_invalid_name_is_rejected(string name) =>
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Sheets.Add(OpenReportBook(), name)).Code);

    [Fact]
    public void Add_accepts_a_name_of_exactly_31_characters_and_chinese_names()
    {
        var id = OpenReportBook();
        _env.Sheets.Add(id, new string('x', 31));
        _env.Sheets.Add(id, "銷售報表");
        Assert.Contains("銷售報表", Names(id));
    }

    // ---- Rename ----

    [Fact]
    public void Rename_updates_formulas_in_other_sheets()
    {
        var id = OpenReportBook();

        var result = _env.Sheets.Rename(id, "Data", "Sales");

        Assert.Equal(["Sales", "Report", "Hidden"], result.Sheets.Select(s => s.Name));
        Assert.Equal("Sales!A1+1", Formula(id, "Report", "A1"));
        Assert.Equal("Sales!A2", Formula(id, "Report", "A2"));
        Assert.Equal(6.0, _env.Cell(id, "A1", sheet: 2));
    }

    [Fact]
    public void Rename_to_a_name_that_needs_quoting_keeps_formulas_valid()
    {
        var id = OpenReportBook();
        _env.Sheets.Rename(id, "Data", "Q1 Sales");

        Assert.Equal("'Q1 Sales'!A1+1", Formula(id, "Report", "A1"));
        Assert.Equal(6.0, _env.Cell(id, "A1", sheet: 2));
    }

    [Fact]
    public void Rename_changes_only_the_case_of_the_same_sheet()
    {
        var id = OpenReportBook();
        Assert.Equal("DATA", _env.Sheets.Rename(id, "Data", "DATA").Sheets[0].Name);
        Assert.Equal("Data", _env.Sheets.Rename(id, "data", "Data").Sheets[0].Name);
    }

    [Fact]
    public void Rename_to_an_existing_name_is_rejected_and_changes_nothing()
    {
        var id = OpenReportBook();
        var ex = Throws(() => _env.Sheets.Rename(id, "Data", "report"));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Equal(["Data", "Report", "Hidden"], Names(id));
        Assert.Equal("Data!A1+1", Formula(id, "Report", "A1"));
    }

    [Fact]
    public void Rename_to_an_invalid_name_is_rejected() =>
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Sheets.Rename(OpenReportBook(), "Data", "a/b")).Code);

    [Fact]
    public void Rename_updates_defined_names_that_point_at_the_sheet()
    {
        var path = _env.MakeBook("names.xlsx", wb =>
        {
            var data = wb.AddWorksheet("Data");
            data.Cell("A1").Value = 1;
            wb.DefinedNames.Add("Total", "Data!$A$1");
            wb.AddWorksheet("Other");
        });
        var id = _env.Files.Open(path).WorkbookId;

        _env.Sheets.Rename(id, "Data", "Sales");

        var refersTo = _env.Sessions.Use(id, false, s => s.Workbook.DefinedNames.Single(n => n.Name == "Total").RefersTo);
        Assert.Contains("Sales", refersTo, StringComparison.Ordinal);
        Assert.DoesNotContain("Data", refersTo, StringComparison.Ordinal);
    }

    // ---- Copy ----

    [Fact]
    public void Copy_defaults_to_a_numbered_name_right_after_the_source()
    {
        var id = OpenReportBook();

        var first = _env.Sheets.Copy(id, "Data");
        var second = _env.Sheets.Copy(id, "Data");

        Assert.Equal(["Data", "Data (2)", "Report", "Hidden"], first.Sheets.Select(s => s.Name));
        Assert.Equal(["Data", "Data (3)", "Data (2)", "Report", "Hidden"], second.Sheets.Select(s => s.Name));
    }

    [Fact]
    public void Copy_default_name_is_truncated_to_fit_31_characters()
    {
        var long31 = new string('x', 31);
        var path = _env.MakeBook("long.xlsx", wb => wb.AddWorksheet(long31));
        var id = _env.Files.Open(path).WorkbookId;

        var name = _env.Sheets.Copy(id, long31).Sheets[1].Name;

        Assert.Equal(31, name.Length);
        Assert.EndsWith(" (2)", name, StringComparison.Ordinal);
    }

    [Fact]
    public void Copy_with_a_name_and_position()
    {
        var id = OpenReportBook();

        var result = _env.Sheets.Copy(id, "Report", "Report backup", position: 1);

        Assert.Equal(["Report backup", "Data", "Report", "Hidden"], result.Sheets.Select(s => s.Name));
    }

    [Fact]
    public void Copy_keeps_values_and_formulas_and_is_independent_of_the_source()
    {
        var id = OpenReportBook();
        _env.Sheets.Copy(id, "Data", "Data copy");

        Assert.Equal(5.0, _env.Cell(id, "A1", sheet: 2));
        Assert.Equal("A1*2", Formula(id, "Data copy", "A2"));
        Assert.Equal(10.0, _env.Cell(id, "A2", sheet: 2));

        _env.SetCell(id, "A1", 100, sheet: 2);
        Assert.Equal(5.0, _env.Cell(id, "A1", sheet: 1));
        Assert.Equal(100.0, _env.Cell(id, "A1", sheet: 2));
    }

    [Fact]
    public void Copy_with_a_duplicate_name_is_rejected() =>
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Sheets.Copy(OpenReportBook(), "Data", "REPORT")).Code);

    [Fact]
    public void Copy_at_an_invalid_position_is_rejected_and_changes_nothing()
    {
        var id = OpenReportBook();
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Sheets.Copy(id, "Data", position: 9)).Code);
        Assert.Equal(["Data", "Report", "Hidden"], Names(id));
    }

    // ---- Delete ----

    [Fact]
    public void Delete_removes_the_sheet()
    {
        var id = OpenReportBook();
        var result = _env.Sheets.Delete(id, "Hidden");

        Assert.Equal(["Data", "Report"], result.Sheets.Select(s => s.Name));
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void Delete_warns_about_formulas_that_referenced_the_deleted_sheet()
    {
        var id = OpenReportBook();

        var result = _env.Sheets.Delete(id, "Data");

        var warning = Assert.Single(result.Warnings);
        Assert.Contains("2 處", warning, StringComparison.Ordinal);
        Assert.Contains("Report!A1", warning, StringComparison.Ordinal);
        Assert.Contains("Report!A2", warning, StringComparison.Ordinal);
        Assert.Contains("#REF!", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void Delete_warns_about_defined_names_and_quoted_references()
    {
        var path = _env.MakeBook("refs.xlsx", wb =>
        {
            wb.AddWorksheet("Q1 Sales").Cell("A1").Value = 1;
            var other = wb.AddWorksheet("Other");
            other.Cell("A1").FormulaA1 = "'Q1 Sales'!A1";
            other.Cell("A2").FormulaA1 = "SUM('Q1 Sales'!A1:A3)";
            wb.DefinedNames.Add("Total", "'Q1 Sales'!$A$1");
        });
        var id = _env.Files.Open(path).WorkbookId;

        var warning = Assert.Single(_env.Sheets.Delete(id, "Q1 Sales").Warnings);

        Assert.Contains("3 處", warning, StringComparison.Ordinal);
        Assert.Contains("已定義名稱 Total", warning, StringComparison.Ordinal);
    }

    [Fact]
    public void Delete_does_not_warn_about_names_that_merely_contain_the_sheet_name()
    {
        var path = _env.MakeBook("lookalike.xlsx", wb =>
        {
            wb.AddWorksheet("Data").Cell("A1").Value = 1;
            var other = wb.AddWorksheet("Other");
            other.Cell("A1").FormulaA1 = "MyData!A1";      // 另一張叫 MyData 的表
            other.Cell("A2").FormulaA1 = "A1+1";           // 一般公式
            other.Cell("A3").FormulaA1 = "\"Data!\"&A1";   // 字串裡的文字不是參照，但無法與參照區分；可接受誤報
            wb.AddWorksheet("MyData").Cell("A1").Value = 2;
        });
        var id = _env.Files.Open(path).WorkbookId;

        var warnings = _env.Sheets.Delete(id, "Data").Warnings;

        // MyData! 不該被當成 Data!；字串常數內的 Data! 在字面上與參照相同，允許被列入
        Assert.DoesNotContain(warnings, w => w.Contains("Other!A1", StringComparison.Ordinal));
    }

    [Fact]
    public void Delete_does_not_count_references_from_the_sheet_to_itself()
    {
        var path = _env.MakeBook("self.xlsx", wb =>
        {
            var data = wb.AddWorksheet("Data");
            data.Cell("A1").Value = 1;
            data.Cell("A2").FormulaA1 = "Data!A1+1";
            wb.AddWorksheet("Other");
        });
        var id = _env.Files.Open(path).WorkbookId;

        Assert.Empty(_env.Sheets.Delete(id, "Data").Warnings);
    }

    [Fact]
    public void Delete_of_the_only_sheet_is_refused()
    {
        var id = _env.Files.Create(_env.Path("one.xlsx")).WorkbookId;
        var ex = Throws(() => _env.Sheets.Delete(id, "Sheet1"));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Equal(["Sheet1"], Names(id));
    }

    [Fact]
    public void Delete_of_the_last_visible_sheet_is_refused_even_if_hidden_sheets_remain()
    {
        var id = OpenReportBook();
        _env.Sheets.Delete(id, "Data");

        var ex = Throws(() => _env.Sheets.Delete(id, "Report"));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("可見", ex.Message, StringComparison.Ordinal);
        Assert.Equal(["Report", "Hidden"], Names(id));
    }

    // ---- Move ----

    [Theory]
    [InlineData("Hidden", 1, new[] { "Hidden", "Data", "Report" })]
    [InlineData("Data", 3, new[] { "Report", "Hidden", "Data" })]
    [InlineData("Data", 2, new[] { "Report", "Data", "Hidden" })]
    [InlineData("Report", 2, new[] { "Data", "Report", "Hidden" })]
    public void Move_to_a_position(string sheet, int position, string[] expected)
    {
        var id = OpenReportBook();
        Assert.Equal(expected, _env.Sheets.Move(id, sheet, position).Sheets.Select(s => s.Name));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(4)]
    public void Move_out_of_range_is_rejected(int position)
    {
        var id = OpenReportBook();
        var ex = Throws(() => _env.Sheets.Move(id, "Data", position));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("1 到 3", ex.Hint, StringComparison.Ordinal);
    }

    // ---- 存檔 ----

    [Fact]
    public void All_changes_survive_save_and_reopen_in_the_same_order()
    {
        var path = _env.MakeBook("book.xlsx", wb =>
        {
            wb.AddWorksheet("Data").Cell("A1").Value = 5;
            wb.AddWorksheet("Report").Cell("A1").FormulaA1 = "Data!A1+1";
            wb.AddWorksheet("Hidden").Visibility = XLWorksheetVisibility.Hidden;
        });
        var id = _env.Files.Open(path).WorkbookId;

        _env.Sheets.Add(id, "Summary", 1);
        _env.Sheets.Rename(id, "Data", "Sales");
        _env.Sheets.Copy(id, "Sales", "Sales copy");
        _env.Sheets.Move(id, "Report", 2);
        var expected = Names(id);
        _env.Files.Save(id);
        _env.Files.Close(id);

        var reopened = _env.Files.Open(path);

        Assert.Equal(expected, reopened.Sheets.Select(s => s.Name));
        Assert.Equal(["Summary", "Report", "Sales", "Sales copy", "Hidden"], expected);
        Assert.Equal("Sales!A1+1", Formula(reopened.WorkbookId, "Report", "A1"));
        Assert.True(reopened.Sheets.Single(s => s.Name == "Hidden").Hidden);
    }

    [Fact]
    public void Files_saved_after_sheet_operations_are_valid_open_xml()
    {
        var id = OpenReportBook();
        var path = _env.Files.ListOpen().Single().Path;

        _env.Sheets.Add(id, "Summary", 1);
        _env.Sheets.Rename(id, "Data", "Q1 Sales");
        _env.Sheets.Copy(id, "Q1 Sales");
        _env.Sheets.Move(id, "Report", 2);
        _env.Sheets.Delete(id, "Summary");
        _env.Files.Save(id);

        Assert.Empty(ExcelEnv.ValidateXlsx(path));
    }

    [Fact]
    public void The_validator_itself_detects_a_broken_workbook()
    {
        var path = _env.MakeWorkbook("broken.xlsx");
        using (var zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/workbook.xml")!;
            string xml;
            using (var reader = new StreamReader(entry.Open()))
            {
                xml = reader.ReadToEnd();
            }

            Assert.Contains("<x:sheets>", xml, StringComparison.Ordinal); // ClosedXML 輸出帶 x: 前綴；沒替換成功就不是有效的對照
            entry.Delete();
            using var writer = new StreamWriter(zip.CreateEntry("xl/workbook.xml").Open());
            writer.Write(xml.Replace("<x:sheets>", "<x:bogus /><x:sheets>", StringComparison.Ordinal));
        }

        Assert.NotEmpty(ExcelEnv.ValidateXlsx(path));
    }

    [Fact]
    public void Newly_created_workbooks_are_valid_open_xml()
    {
        var info = _env.Files.Create(_env.Path("fresh.xlsx"));
        Assert.Empty(ExcelEnv.ValidateXlsx(info.Path));
    }

    [Fact]
    public void Sheet_changes_on_a_read_only_workbook_are_refused()
    {
        File.Move(_env.MakeWorkbook("book.xlsx"), _env.Path("macro.xlsm"));
        var id = _env.Files.Open(_env.Path("macro.xlsm")).WorkbookId;

        foreach (var act in new Action[]
                 {
                     () => _env.Sheets.Add(id, "X"),
                     () => _env.Sheets.Rename(id, "Data", "X"),
                     () => _env.Sheets.Copy(id, "Data"),
                     () => _env.Sheets.Delete(id, "Data"),
                     () => _env.Sheets.Move(id, "Data", 1),
                 })
        {
            Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(act).Code);
        }

        Assert.Equal(["Data"], Names(id)); // 唯讀仍可列出
    }

    [Fact]
    public void Read_only_operations_do_not_mark_the_workbook_dirty()
    {
        var id = OpenReportBook();
        _env.Sheets.ListSheets(id);
        _env.Sheets.GetSheetInfo(id, "Data");

        Assert.False(_env.Files.ListOpen().Single().IsDirty);
    }

    [Fact]
    public void A_rejected_operation_that_fails_validation_leaves_the_workbook_unchanged()
    {
        var id = OpenReportBook();
        Throws(() => _env.Sheets.Rename(id, "Data", "Report"));
        Throws(() => _env.Sheets.Add(id, "Data"));

        Assert.Equal(["Data", "Report", "Hidden"], Names(id));
        Assert.Equal("Data!A1+1", Formula(id, "Report", "A1"));
    }
}
