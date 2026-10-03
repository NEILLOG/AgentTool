using ClosedXML.Excel;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Tests;

public sealed class FileOperationsTests : IDisposable
{
    private readonly ExcelEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    // ---- Create ----

    [Fact]
    public void Create_writes_a_file_and_returns_a_clean_session()
    {
        var info = _env.Files.Create(_env.Path("new.xlsx"));

        Assert.True(File.Exists(_env.Path("new.xlsx")));
        Assert.StartsWith("wb_", info.WorkbookId, StringComparison.Ordinal);
        Assert.False(info.IsDirty);
        Assert.False(info.ReadOnly);
        Assert.Equal("Sheet1", Assert.Single(info.Sheets).Name);
        using var wb = new XLWorkbook(_env.Path("new.xlsx"));
        Assert.Equal("Sheet1", wb.Worksheet(1).Name);
    }

    [Fact]
    public void Create_uses_the_given_sheet_name() =>
        Assert.Equal("銷售", Assert.Single(_env.Files.Create(_env.Path("new.xlsx"), "銷售").Sheets).Name);

    [Fact]
    public void Create_over_an_existing_file_returns_FILE_EXISTS_and_leaves_it_alone()
    {
        var path = _env.MakeWorkbook("a.xlsx");
        var before = File.ReadAllBytes(path);

        Assert.Equal(ErrorCodes.FileExists, Throws(() => _env.Files.Create(path)).Code);
        Assert.Equal(before, File.ReadAllBytes(path));
    }

    [Fact]
    public void Create_over_an_existing_file_with_AllowOverwrite_makes_a_backup()
    {
        using var env = new ExcelEnv(allowOverwrite: true);
        var path = env.MakeWorkbook("a.xlsx");

        env.Files.Create(path);

        Assert.Single(Directory.GetFiles(env.Root, "a.backup-*.xlsx"));
        Assert.Equal("Sheet1", ExcelEnvSheetName(path));
    }

    private static string ExcelEnvSheetName(string path)
    {
        using var wb = new XLWorkbook(path);
        return wb.Worksheet(1).Name;
    }

    [Theory]
    [InlineData("new.xls")]
    [InlineData("new.xlsm")]
    [InlineData("new.csv")]
    public void Create_only_allows_xlsx(string name) =>
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Files.Create(_env.Path(name))).Code);

    [Fact]
    public void Create_in_a_missing_folder_returns_FILE_NOT_FOUND() =>
        Assert.Equal(ErrorCodes.FileNotFound, Throws(() => _env.Files.Create(_env.Path("nope/new.xlsx"))).Code);

    [Fact]
    public void Create_outside_the_roots_is_rejected() =>
        Assert.Equal(ErrorCodes.PathNotAllowed, Throws(() => _env.Files.Create(_env.Outside("new.xlsx"))).Code);

    [Theory]
    [InlineData("")]
    [InlineData("a/b")]
    [InlineData("a:b")]
    [InlineData("[x]")]
    [InlineData("'quoted'")]
    [InlineData("這個名稱超過三十一個字元的限制所以應該被拒絕才對喔喔喔喔喔喔喔喔")]
    public void Create_rejects_invalid_sheet_names(string name) =>
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Files.Create(_env.Path("new.xlsx"), name)).Code);

    // ---- Open ----

    [Fact]
    public void Open_returns_sheet_summaries_with_used_range()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        var info = _env.Files.Open(path);

        var sheet = Assert.Single(info.Sheets);
        Assert.Equal("Data", sheet.Name);
        Assert.Equal("A1:B2", sheet.UsedRange);
        Assert.False(sheet.Hidden);
        Assert.False(info.AlreadyOpen);
        Assert.Empty(info.PreservationWarnings);
        Assert.Equal("apple", _env.Cell(info.WorkbookId, "A2"));
    }

    [Fact]
    public void Open_of_an_empty_sheet_has_no_used_range()
    {
        var info = _env.Files.Create(_env.Path("empty.xlsx"));
        Assert.Null(info.Sheets[0].UsedRange);
    }

    [Fact]
    public void Opening_the_same_path_twice_returns_the_existing_session()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        var first = _env.Files.Open(path);
        var second = _env.Files.Open(path);

        Assert.Equal(first.WorkbookId, second.WorkbookId);
        Assert.True(second.AlreadyOpen);
        Assert.Single(_env.Files.ListOpen());
    }

    [Fact]
    public void Open_does_not_keep_the_file_locked()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        _env.Files.Open(path);

        // 獨占開啟會在檔案仍被持有時失敗（Windows 為共享違規，Unix 為 flock）
        using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        Assert.True(exclusive.CanWrite);
    }

    [Fact]
    public void Open_of_an_exclusively_locked_file_returns_FILE_LOCKED()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        Assert.Equal(ErrorCodes.FileLocked, Throws(() => _env.Files.Open(path)).Code);
    }

    [Fact]
    public void Open_missing_file_returns_FILE_NOT_FOUND() =>
        Assert.Equal(ErrorCodes.FileNotFound, Throws(() => _env.Files.Open(_env.Path("missing.xlsx"))).Code);

    [Theory]
    [InlineData("old.xls")]
    [InlineData("doc.docx")]
    [InlineData("slides.pptx")]
    [InlineData("report.pdf")]
    public void Open_rejects_formats_the_excel_tools_cannot_handle(string name)
    {
        File.WriteAllText(_env.Path(name), "x");
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Files.Open(_env.Path(name))).Code);
    }

    [Fact]
    public void Xlsm_opens_read_only()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        File.Move(path, _env.Path("macro.xlsm"));

        var info = _env.Files.Open(_env.Path("macro.xlsm"));

        Assert.True(info.ReadOnly);
        Assert.Equal("apple", _env.Cell(info.WorkbookId, "A2"));
    }

    [Fact]
    public void Password_protected_or_legacy_files_return_PASSWORD_PROTECTED()
    {
        byte[] ole = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0];
        File.WriteAllBytes(_env.Path("locked.xlsx"), ole);
        Assert.Equal(ErrorCodes.PasswordProtected, Throws(() => _env.Files.Open(_env.Path("locked.xlsx"))).Code);
    }

    [Theory]
    [InlineData("not a zip at all")]
    [InlineData("")]
    public void Garbage_returns_CORRUPT_FILE(string content)
    {
        File.WriteAllText(_env.Path("bad.xlsx"), content);
        Assert.Equal(ErrorCodes.CorruptFile, Throws(() => _env.Files.Open(_env.Path("bad.xlsx"))).Code);
    }

    [Fact]
    public void Truncated_zip_returns_CORRUPT_FILE()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        var bytes = File.ReadAllBytes(path);
        File.WriteAllBytes(path, bytes[..(bytes.Length / 2)]);
        Assert.Equal(ErrorCodes.CorruptFile, Throws(() => _env.Files.Open(path)).Code);
    }

    [Fact]
    public void Zip_without_a_workbook_part_returns_CORRUPT_FILE()
    {
        var path = _env.Path("empty-zip.xlsx");
        using (var zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Create))
        {
            zip.CreateEntry("hello.txt");
        }

        Assert.Equal(ErrorCodes.CorruptFile, Throws(() => _env.Files.Open(path)).Code);
    }

    [Fact]
    public void Zip_bomb_is_rejected_by_uncompressed_size()
    {
        using var env = new ExcelEnv(maxUncompressedMb: 1);
        var path = env.MakeWorkbook("book.xlsx");
        using (var zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Update))
        {
            using var s = zip.CreateEntry("xl/filler.bin").Open();
            s.Write(new byte[3 * 1024 * 1024]); // 壓縮後只有幾 KB
        }

        Assert.True(new FileInfo(path).Length < 100_000);
        Assert.Equal(ErrorCodes.FileTooLarge, Throws(() => env.Files.Open(path)).Code);
    }

    // ---- 資料保全 ----

    [Theory]
    [InlineData("xl/charts/chart1.xml", "圖表")]
    [InlineData("xl/pivotTables/pivotTable1.xml", "樞紐分析表")]
    [InlineData("xl/pivotCache/pivotCacheDefinition1.xml", "樞紐分析表")]
    [InlineData("xl/vbaProject.bin", "VBA 巨集")]
    [InlineData("xl/slicers/slicer1.xml", "交叉分析篩選器")]
    [InlineData("xl/ctrlProps/ctrlProp1.xml", "表單控制項")]
    [InlineData("xl/embeddings/oleObject1.bin", "內嵌物件")]
    [InlineData("xl/threadedComments/threadedComment1.xml", "討論串註解")]
    public void Open_reports_content_that_clearedxml_would_lose(string part, string label)
    {
        var path = _env.MakeWorkbook("book.xlsx");
        ExcelEnv.AddEntries(path, (part, "<x/>"));

        var info = _env.Files.Open(path);

        Assert.Contains(label, info.PreservationWarnings);
    }

    [Fact]
    public void Drawing_shapes_are_reported_but_pictures_are_not()
    {
        var shapes = _env.MakeWorkbook("shapes.xlsx");
        ExcelEnv.AddEntries(shapes, ("xl/drawings/drawing1.xml", "<xdr:wsDr><xdr:twoCellAnchor><xdr:sp macro=\"\"><xdr:spPr/></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"));
        var pictures = _env.MakeWorkbook("pictures.xlsx");
        ExcelEnv.AddEntries(pictures, ("xl/drawings/drawing1.xml", "<xdr:wsDr><xdr:twoCellAnchor><xdr:pic><xdr:spPr/></xdr:pic></xdr:twoCellAnchor></xdr:wsDr>"));

        Assert.Contains("圖形 / 文字方塊", _env.Files.Open(shapes).PreservationWarnings);
        Assert.Empty(_env.Files.Open(pictures).PreservationWarnings);
    }

    [Fact]
    public void Save_is_refused_when_the_original_has_content_that_would_be_lost()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        ExcelEnv.AddEntries(path, ("xl/charts/chart1.xml", "<x/>"), ("xl/pivotTables/p.xml", "<x/>"));
        var before = File.ReadAllBytes(path);
        var id = _env.Files.Open(path).WorkbookId;
        _env.SetCell(id, "A2", "changed");

        var ex = Throws(() => _env.Files.Save(id));

        Assert.Equal(ErrorCodes.UnsafeToOverwrite, ex.Code);
        Assert.Contains("圖表", ex.Message);
        Assert.Contains("save_as", ex.Hint);
        Assert.Equal(before, File.ReadAllBytes(path));
        Assert.Empty(Directory.GetFiles(_env.Root, "*.backup-*"));
    }

    [Fact]
    public void SaveAs_lets_a_workbook_with_lossy_content_be_saved_and_then_Save_works()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        ExcelEnv.AddEntries(path, ("xl/charts/chart1.xml", "<x/>"));
        var original = File.ReadAllBytes(path);
        var id = _env.Files.Open(path).WorkbookId;
        _env.SetCell(id, "A2", "changed");

        var result = _env.Files.SaveAs(id, _env.Path("copy.xlsx"));

        Assert.Equal(_env.Resolved("copy.xlsx"), result.Path);
        Assert.Contains("圖表", result.PreservationWarnings);
        Assert.Equal(original, File.ReadAllBytes(path));
        Assert.Equal("changed", ExcelEnv.ReadFromDisk(_env.Path("copy.xlsx"), "A2"));

        _env.SetCell(id, "A2", "again");
        _env.Files.Save(id);
        Assert.Equal("again", ExcelEnv.ReadFromDisk(_env.Path("copy.xlsx"), "A2"));
        Assert.Equal(original, File.ReadAllBytes(path));
    }

    // ---- Save ----

    [Fact]
    public void Save_writes_changes_and_backs_up_the_original_once()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        var original = File.ReadAllBytes(path);
        var id = _env.Files.Open(path).WorkbookId;

        _env.SetCell(id, "A2", "changed");
        Assert.True(_env.Files.ListOpen().Single().IsDirty);
        var first = _env.Files.Save(id);

        Assert.Equal(_env.Resolved("book.xlsx"), first.Path);
        Assert.NotNull(first.BackupPath);
        Assert.Equal(original, File.ReadAllBytes(first.BackupPath!));
        Assert.Equal("changed", ExcelEnv.ReadFromDisk(path, "A2"));
        Assert.False(_env.Files.ListOpen().Single().IsDirty);

        _env.SetCell(id, "A2", "changed again");
        var second = _env.Files.Save(id);

        Assert.Equal(first.BackupPath, second.BackupPath);
        Assert.Single(Directory.GetFiles(_env.Root, "*.backup-*"));
        Assert.Equal("changed again", ExcelEnv.ReadFromDisk(path, "A2"));
    }

    [Fact]
    public void Saving_many_times_in_a_row_keeps_working_and_keeps_every_change()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        var id = _env.Files.Open(path).WorkbookId;

        for (var i = 1; i <= 5; i++)
        {
            _env.SetCell(id, "A2", $"v{i}");
            _env.Files.Save(id);
            Assert.Equal($"v{i}", ExcelEnv.ReadFromDisk(path, "A2"));
        }

        // Save 與 SaveAs 交錯
        _env.Files.SaveAs(id, _env.Path("copy.xlsx"));
        _env.SetCell(id, "A2", "after copy");
        _env.Files.Save(id);
        Assert.Equal("after copy", ExcelEnv.ReadFromDisk(_env.Path("copy.xlsx"), "A2"));
        Assert.Equal("v5", ExcelEnv.ReadFromDisk(path, "A2"));
        Assert.Equal("after copy", _env.Cell(id, "A2"));
    }

    [Fact]
    public void Save_leaves_no_temp_files_behind()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        var id = _env.Files.Open(path).WorkbookId;
        _env.SetCell(id, "A1", "x");
        _env.Files.Save(id);

        Assert.Empty(Directory.GetFiles(_env.Root, "*.tmp.xlsx", SearchOption.AllDirectories));
        Assert.Empty(Directory.GetFiles(_env.Root, ".*", SearchOption.AllDirectories));
    }

    [Fact]
    public void Save_failure_keeps_the_original_file_intact()
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        var path = _env.MakeWorkbook("book.xlsx");
        var original = File.ReadAllBytes(path);
        var id = _env.Files.Open(path).WorkbookId;
        _env.SetCell(id, "A2", "changed");

        var mode = File.GetUnixFileMode(_env.Root);
        File.SetUnixFileMode(_env.Root, UnixFileMode.UserRead | UnixFileMode.UserExecute); // 資料夾唯讀，寫不出暫存檔
        try
        {
            Assert.Throws<OfficeToolException>(() => _env.Files.Save(id));
        }
        finally
        {
            File.SetUnixFileMode(_env.Root, mode);
        }

        Assert.Equal(original, File.ReadAllBytes(path));
        Assert.True(_env.Files.ListOpen().Single().IsDirty);
    }

    [Fact]
    public void Save_of_a_read_only_xlsm_is_refused()
    {
        File.Move(_env.MakeWorkbook("book.xlsx"), _env.Path("macro.xlsm"));
        var id = _env.Files.Open(_env.Path("macro.xlsm")).WorkbookId;

        var ex = Throws(() => _env.Files.Save(id));

        Assert.Equal(ErrorCodes.UnsupportedFormat, ex.Code);
        Assert.Contains("save_as", ex.Hint);
    }

    [Fact]
    public void Modifying_a_read_only_workbook_is_refused()
    {
        File.Move(_env.MakeWorkbook("book.xlsx"), _env.Path("macro.xlsm"));
        var id = _env.Files.Open(_env.Path("macro.xlsm")).WorkbookId;
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.SetCell(id, "A1", "x")).Code);
    }

    // ---- SaveAs ----

    [Fact]
    public void SaveAs_writes_a_new_file_and_the_session_follows_it()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        var original = File.ReadAllBytes(path);
        var id = _env.Files.Open(path).WorkbookId;
        _env.SetCell(id, "A2", "changed");

        var result = _env.Files.SaveAs(id, _env.Path("copy.xlsx"));

        Assert.Equal(_env.Resolved("copy.xlsx"), result.Path);
        Assert.Null(result.BackupPath);
        Assert.Equal(original, File.ReadAllBytes(path));
        var entry = _env.Files.ListOpen().Single();
        Assert.Equal(_env.Resolved("copy.xlsx"), entry.Path);
        Assert.False(entry.IsDirty);
    }

    [Fact]
    public void SaveAs_onto_an_existing_file_returns_FILE_EXISTS()
    {
        var id = _env.Files.Open(_env.MakeWorkbook("a.xlsx")).WorkbookId;
        _env.MakeWorkbook("b.xlsx");
        Assert.Equal(ErrorCodes.FileExists, Throws(() => _env.Files.SaveAs(id, _env.Path("b.xlsx"))).Code);
    }

    [Fact]
    public void SaveAs_onto_an_existing_file_with_AllowOverwrite_backs_it_up()
    {
        using var env = new ExcelEnv(allowOverwrite: true);
        var id = env.Files.Open(env.MakeWorkbook("a.xlsx")).WorkbookId;
        var target = env.MakeWorkbook("b.xlsx", ws => ws.Cell("A1").Value = "old b");

        var result = env.Files.SaveAs(id, target);

        Assert.NotNull(result.BackupPath);
        Assert.Equal("old b", ExcelEnv.ReadFromDisk(result.BackupPath!, "A1"));
        Assert.Equal("name", ExcelEnv.ReadFromDisk(target, "A1"));
    }

    [Fact]
    public void SaveAs_onto_a_path_open_in_another_session_is_refused()
    {
        using var env = new ExcelEnv(allowOverwrite: true); // 排除 FILE_EXISTS 來自 AllowOverwrite 的情況
        var a = env.Files.Open(env.MakeWorkbook("a.xlsx")).WorkbookId;
        env.Files.Open(env.MakeWorkbook("b.xlsx"));

        var ex = Throws(() => env.Files.SaveAs(a, env.Path("b.xlsx")));

        Assert.Equal(ErrorCodes.FileExists, ex.Code);
        Assert.Contains("另一個活頁簿", ex.Message);
    }

    [Theory]
    [InlineData("copy.xls")]
    [InlineData("copy.xlsm")]
    [InlineData("copy.docx")]
    public void SaveAs_only_writes_xlsx(string name)
    {
        var id = _env.Files.Open(_env.MakeWorkbook("a.xlsx")).WorkbookId;
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Files.SaveAs(id, _env.Path(name))).Code);
    }

    [Fact]
    public void SaveAs_outside_the_roots_is_rejected()
    {
        var id = _env.Files.Open(_env.MakeWorkbook("a.xlsx")).WorkbookId;
        Assert.Equal(ErrorCodes.PathNotAllowed, Throws(() => _env.Files.SaveAs(id, _env.Outside("copy.xlsx"))).Code);
    }

    [Fact]
    public void SaveAs_turns_a_read_only_xlsm_into_an_editable_xlsx()
    {
        File.Move(_env.MakeWorkbook("book.xlsx"), _env.Path("macro.xlsm"));
        var id = _env.Files.Open(_env.Path("macro.xlsm")).WorkbookId;

        _env.Files.SaveAs(id, _env.Path("editable.xlsx"));

        Assert.False(_env.Files.ListOpen().Single().ReadOnly);
        _env.SetCell(id, "A2", "now editable");
        _env.Files.Save(id);
        Assert.Equal("now editable", ExcelEnv.ReadFromDisk(_env.Path("editable.xlsx"), "A2"));
    }

    // ---- Close / ListOpen ----

    [Fact]
    public void Close_with_unsaved_changes_is_refused_unless_discarded()
    {
        var id = _env.Files.Open(_env.MakeWorkbook("book.xlsx")).WorkbookId;
        _env.SetCell(id, "A1", "dirty");

        Assert.Equal(ErrorCodes.UnsavedChanges, Throws(() => _env.Files.Close(id)).Code);
        Assert.Single(_env.Files.ListOpen());

        _env.Files.Close(id, discardChanges: true);
        Assert.Empty(_env.Files.ListOpen());
        Assert.Equal("name", ExcelEnv.ReadFromDisk(_env.Path("book.xlsx"), "A1"));
    }

    [Fact]
    public void Closing_a_clean_workbook_works_and_the_id_is_gone()
    {
        var id = _env.Files.Open(_env.MakeWorkbook("book.xlsx")).WorkbookId;
        _env.Files.Close(id);

        Assert.Equal(ErrorCodes.SessionNotFound, Throws(() => _env.Files.Save(id)).Code);
        Assert.Equal(ErrorCodes.SessionNotFound, Throws(() => _env.Files.Close(id)).Code);
    }

    [Theory]
    [InlineData("wb_doesnotexist")]
    [InlineData("")]
    public void Unknown_workbook_id_returns_SESSION_NOT_FOUND(string id)
    {
        var ex = Throws(() => _env.Files.Save(id));
        Assert.Equal(ErrorCodes.SessionNotFound, ex.Code);
        Assert.Contains("open", ex.Hint);
    }

    [Fact]
    public void A_closed_path_can_be_reopened_with_a_new_id()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        var first = _env.Files.Open(path).WorkbookId;
        _env.Files.Close(first);

        var second = _env.Files.Open(path);

        Assert.NotEqual(first, second.WorkbookId);
        Assert.False(second.AlreadyOpen);
    }

    // ---- ListFiles ----

    [Fact]
    public void ListFiles_lists_only_excel_files_and_skips_office_lock_files()
    {
        _env.MakeWorkbook("b.xlsx");
        _env.MakeWorkbook("a.xlsx");
        File.WriteAllText(_env.Path("macro.xlsm"), "x");
        File.WriteAllText(_env.Path("~$a.xlsx"), "lock");
        File.WriteAllText(_env.Path("notes.txt"), "x");
        File.WriteAllText(_env.Path("old.xls"), "x");

        var listing = _env.Files.ListFiles(_env.Root);

        Assert.Equal(["a.xlsx", "b.xlsx", "macro.xlsm"], listing.Files.Select(f => f.Name));
        Assert.False(listing.Truncated);
        Assert.All(listing.Files, f => Assert.True(f.SizeBytes > 0));
    }

    [Fact]
    public void ListFiles_is_not_recursive_unless_asked()
    {
        _env.MakeWorkbook("top.xlsx");
        _env.MakeWorkbook("sub/inner.xlsx");

        Assert.Equal(["top.xlsx"], _env.Files.ListFiles(_env.Root).Files.Select(f => f.Name));
        Assert.Equal(["inner.xlsx", "top.xlsx"], _env.Files.ListFiles(_env.Root, recursive: true).Files.Select(f => f.Name).Order());
    }

    [Fact]
    public void ListFiles_truncates_at_maxEntries()
    {
        for (var i = 0; i < 5; i++)
        {
            _env.MakeWorkbook($"f{i}.xlsx");
        }

        var listing = _env.Files.ListFiles(_env.Root, maxEntries: 3);

        Assert.Equal(3, listing.Files.Count);
        Assert.True(listing.Truncated);
    }

    [Fact]
    public void ListFiles_outside_the_roots_or_missing_is_rejected()
    {
        Assert.Equal(ErrorCodes.PathNotAllowed, Throws(() => _env.Files.ListFiles(_env.Outside("x"))).Code);
        Assert.Equal(ErrorCodes.FileNotFound, Throws(() => _env.Files.ListFiles(_env.Path("missing"))).Code);
    }

    [Fact]
    public void ListFiles_hides_symlinks_that_point_outside_the_roots()
    {
        var secret = _env.Outside("secret.xlsx");
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("S");
            wb.SaveAs(secret);
        }

        try
        {
            File.CreateSymbolicLink(_env.Path("link.xlsx"), secret);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return; // 沒有建立 symlink 的權限（Windows）
        }

        _env.MakeWorkbook("real.xlsx");

        Assert.Equal(["real.xlsx"], _env.Files.ListFiles(_env.Root).Files.Select(f => f.Name));
    }
}
