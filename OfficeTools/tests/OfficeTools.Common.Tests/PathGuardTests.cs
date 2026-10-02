using OfficeTools.Common.Errors;
using OfficeTools.Common.Security;

namespace OfficeTools.Common.Tests;

public sealed class PathGuardTests : IDisposable
{
    private readonly TempDirectory _tmp = new();

    public void Dispose() => _tmp.Dispose();

    private PathGuard Guard(int maxMb = 50, bool allowOverwrite = false, params string[] roots) =>
        new(new OfficeToolsOptions
        {
            AllowedRoots = roots.Length == 0 ? [_tmp.Combine("root")] : roots,
            MaxFileSizeMb = maxMb,
            AllowOverwrite = allowOverwrite,
        });

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    // ---- 根目錄限制 ----

    [Fact]
    public void Path_inside_root_is_allowed()
    {
        var file = _tmp.CreateFile("root/sub/book.xlsx");
        Assert.EndsWith("book.xlsx", Guard().ResolveAllowed(file));
    }

    [Fact]
    public void Root_itself_is_allowed()
    {
        Directory.CreateDirectory(_tmp.Combine("root"));
        Guard().ResolveAllowed(_tmp.Combine("root"));
    }

    [Fact]
    public void Path_outside_root_is_rejected_and_hint_lists_roots()
    {
        var outside = _tmp.CreateFile("other/book.xlsx");
        var ex = Throws(() => Guard().ResolveAllowed(outside));
        Assert.Equal(ErrorCodes.PathNotAllowed, ex.Code);
        Assert.Contains("root", ex.Hint);
    }

    [Fact]
    public void Empty_AllowedRoots_rejects_everything()
    {
        var file = _tmp.CreateFile("root/book.xlsx");
        var guard = new PathGuard(new OfficeToolsOptions());
        var ex = Throws(() => guard.ResolveAllowed(file));
        Assert.Equal(ErrorCodes.PathNotAllowed, ex.Code);
        Assert.Contains("AllowedRoots", ex.Hint);
    }

    [Fact]
    public void Sibling_folder_sharing_the_root_prefix_is_rejected()
    {
        var sibling = _tmp.CreateFile("root-evil/book.xlsx");
        Assert.Equal(ErrorCodes.PathNotAllowed, Throws(() => Guard().ResolveAllowed(sibling)).Code);
    }

    [Fact]
    public void Dot_dot_traversal_out_of_root_is_rejected()
    {
        _tmp.CreateFile("other/secret.xlsx");
        var sneaky = _tmp.Combine("root", "..", "other", "secret.xlsx");
        Assert.Equal(ErrorCodes.PathNotAllowed, Throws(() => Guard().ResolveAllowed(sneaky)).Code);
    }

    [Fact]
    public void Dot_dot_that_stays_inside_root_is_allowed()
    {
        _tmp.CreateFile("root/book.xlsx");
        Directory.CreateDirectory(_tmp.Combine("root", "sub"));
        var path = _tmp.Combine("root", "sub", "..", "book.xlsx");
        Assert.EndsWith("book.xlsx", Guard().ResolveAllowed(path));
    }

    // ---- symlink ----

    [Fact]
    public void Symlink_inside_root_pointing_outside_is_rejected()
    {
        if (!CanCreateSymlinks())
        {
            return;
        }

        _tmp.CreateFile("other/secret.xlsx");
        Directory.CreateDirectory(_tmp.Combine("root"));
        Directory.CreateSymbolicLink(_tmp.Combine("root", "link"), _tmp.Combine("other"));

        var ex = Throws(() => Guard().ResolveAllowed(_tmp.Combine("root", "link", "secret.xlsx")));
        Assert.Equal(ErrorCodes.PathNotAllowed, ex.Code);
    }

    [Fact]
    public void File_symlink_inside_root_pointing_outside_is_rejected()
    {
        if (!CanCreateSymlinks())
        {
            return;
        }

        var secret = _tmp.CreateFile("other/secret.xlsx");
        Directory.CreateDirectory(_tmp.Combine("root"));
        File.CreateSymbolicLink(_tmp.Combine("root", "book.xlsx"), secret);

        Assert.Equal(ErrorCodes.PathNotAllowed, Throws(() => Guard().ResolveExistingFile(_tmp.Combine("root", "book.xlsx"))).Code);
    }

    [Fact]
    public void Root_given_as_symlink_matches_paths_through_the_real_location()
    {
        if (!CanCreateSymlinks())
        {
            return;
        }

        var file = _tmp.CreateFile("real/book.xlsx");
        var link = _tmp.Combine("link");
        Directory.CreateSymbolicLink(link, _tmp.Combine("real"));

        var guard = new PathGuard(new OfficeToolsOptions { AllowedRoots = [link] });
        guard.ResolveAllowed(file);
        guard.ResolveAllowed(Path.Combine(link, "book.xlsx"));
    }

    private bool CanCreateSymlinks()
    {
        try
        {
            Directory.CreateDirectory(_tmp.Combine("probe-target"));
            Directory.CreateSymbolicLink(_tmp.Combine("probe-link"), _tmp.Combine("probe-target"));
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false; // Windows 沒有權限建立 symlink 時略過
        }
    }

    // ---- 可疑路徑格式 ----

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("a\0b.xlsx")]
    [InlineData(@"\\server\share\book.xlsx")]
    [InlineData("//server/share/book.xlsx")]
    [InlineData(@"\\?\C:\data\book.xlsx")]
    [InlineData(@"C:\data\book.xlsx:hidden")]
    [InlineData("book.xlsx:stream")]
    public void Suspicious_paths_are_rejected(string path)
    {
        Assert.Equal(ErrorCodes.PathNotAllowed, Throws(() => Guard().ResolveAllowed(path)).Code);
    }

    [Fact]
    public void Windows_drive_letter_colon_alone_is_not_treated_as_stream()
    {
        // 驗證冒號檢查會放行開頭的磁碟機代號：這個路徑會因為不在根目錄內而被拒絕，
        // 但錯誤訊息不能是「不允許的冒號」
        var ex = Throws(() => Guard().ResolveAllowed(@"C:\data\book.xlsx"));
        Assert.DoesNotContain("冒號", ex.Message);
    }

    // ---- 讀取：副檔名、存在、大小 ----

    [Theory]
    [InlineData("old.xls")]
    [InlineData("old.doc")]
    [InlineData("old.ppt")]
    [InlineData("notes.txt")]
    [InlineData("noext")]
    public void Unsupported_extensions_are_rejected(string name)
    {
        var file = _tmp.CreateFile($"root/{name}");
        var ex = Throws(() => Guard().ResolveExistingFile(file));
        Assert.Equal(ErrorCodes.UnsupportedFormat, ex.Code);
        Assert.Contains("另存為新格式", ex.Hint);
    }

    [Theory]
    [InlineData("a.xlsx")]
    [InlineData("a.XLSX")]
    [InlineData("a.xlsm")]
    [InlineData("a.docx")]
    [InlineData("a.pptx")]
    [InlineData("a.pdf")]
    public void Whitelisted_extensions_are_accepted_case_insensitively(string name)
    {
        var file = _tmp.CreateFile($"root/{name}");
        Assert.Equal(Path.GetFileName(file), Path.GetFileName(Guard().ResolveExistingFile(file)));
    }

    [Fact]
    public void Missing_file_returns_FILE_NOT_FOUND()
    {
        Directory.CreateDirectory(_tmp.Combine("root"));
        var ex = Throws(() => Guard().ResolveExistingFile(_tmp.Combine("root", "missing.xlsx")));
        Assert.Equal(ErrorCodes.FileNotFound, ex.Code);
    }

    [Fact]
    public void Directory_with_a_file_extension_is_not_a_file()
    {
        Directory.CreateDirectory(_tmp.Combine("root", "folder.xlsx"));
        var ex = Throws(() => Guard().ResolveExistingFile(_tmp.Combine("root", "folder.xlsx")));
        Assert.Equal(ErrorCodes.FileNotFound, ex.Code);
    }

    [Fact]
    public void File_over_size_limit_returns_FILE_TOO_LARGE()
    {
        var file = _tmp.CreateFile("root/big.xlsx", length: 2 * 1024 * 1024);
        var ex = Throws(() => Guard(maxMb: 1).ResolveExistingFile(file));
        Assert.Equal(ErrorCodes.FileTooLarge, ex.Code);
    }

    [Fact]
    public void File_at_the_size_limit_is_accepted()
    {
        var file = _tmp.CreateFile("root/exact.xlsx", length: 1024 * 1024);
        Guard(maxMb: 1).ResolveExistingFile(file);
    }

    [Fact]
    public void Extension_is_checked_before_existence()
    {
        Directory.CreateDirectory(_tmp.Combine("root"));
        var ex = Throws(() => Guard().ResolveExistingFile(_tmp.Combine("root", "missing.xls")));
        Assert.Equal(ErrorCodes.UnsupportedFormat, ex.Code);
    }

    // ---- 寫入：覆寫政策 ----

    [Fact]
    public void Write_target_that_does_not_exist_is_accepted()
    {
        Directory.CreateDirectory(_tmp.Combine("root"));
        Guard().ResolveWriteTarget(_tmp.Combine("root", "new.xlsx"));
    }

    [Fact]
    public void Existing_write_target_returns_FILE_EXISTS_by_default()
    {
        var file = _tmp.CreateFile("root/book.xlsx");
        var ex = Throws(() => Guard().ResolveWriteTarget(file));
        Assert.Equal(ErrorCodes.FileExists, ex.Code);
    }

    [Fact]
    public void Existing_write_target_is_accepted_when_AllowOverwrite_is_true()
    {
        var file = _tmp.CreateFile("root/book.xlsx");
        Guard(allowOverwrite: true).ResolveWriteTarget(file);
    }

    [Fact]
    public void Write_target_outside_root_is_rejected_even_with_AllowOverwrite()
    {
        var outside = _tmp.Combine("other", "new.xlsx");
        Assert.Equal(ErrorCodes.PathNotAllowed, Throws(() => Guard(allowOverwrite: true).ResolveWriteTarget(outside)).Code);
    }

    [Fact]
    public void Write_target_with_unsupported_extension_is_rejected()
    {
        Directory.CreateDirectory(_tmp.Combine("root"));
        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => Guard().ResolveWriteTarget(_tmp.Combine("root", "new.exe"))).Code);
    }

    // ---- 備份 ----

    [Fact]
    public void CreateBackup_copies_the_file_next_to_the_original_with_a_timestamp()
    {
        var file = _tmp.CreateFile("root/book.xlsx");
        File.WriteAllText(file, "original");
        var guard = new PathGuard(
            new OfficeToolsOptions { AllowedRoots = [_tmp.Combine("root")] },
            new FixedTimeProvider(new DateTimeOffset(2026, 10, 3, 14, 5, 9, TimeSpan.Zero)));

        var backup = guard.CreateBackup(file);

        Assert.Equal("book.backup-20261003-140509.xlsx", Path.GetFileName(backup));
        Assert.Equal(Path.GetDirectoryName(file), Path.GetDirectoryName(backup));
        Assert.Equal("original", File.ReadAllText(backup));
        Assert.Equal("original", File.ReadAllText(file));
    }

    [Fact]
    public void CreateBackup_twice_in_the_same_second_does_not_overwrite()
    {
        var file = _tmp.CreateFile("root/book.xlsx");
        var guard = new PathGuard(
            new OfficeToolsOptions { AllowedRoots = [_tmp.Combine("root")] },
            new FixedTimeProvider(DateTimeOffset.UnixEpoch));

        var first = guard.CreateBackup(file);
        var second = guard.CreateBackup(file);

        Assert.NotEqual(first, second);
        Assert.True(File.Exists(first));
        Assert.True(File.Exists(second));
    }
}

public class OfficeToolExceptionTests
{
    [Fact]
    public void Exposes_code_and_hint()
    {
        var ex = new OfficeToolException(ErrorCodes.SheetNotFound, "msg", "hint");
        Assert.Equal("SHEET_NOT_FOUND", ex.Code);
        Assert.Equal("msg", ex.Message);
        Assert.Equal("hint", ex.Hint);
    }

    [Fact]
    public void Error_codes_match_the_names_documented_in_the_plan()
    {
        var codes = typeof(ErrorCodes).GetFields().Select(f => (string)f.GetRawConstantValue()!).ToHashSet();
        var planTable = File.ReadAllText(FindPlanFile("04-common-security-errors.md"));
        foreach (var code in codes)
        {
            Assert.Contains($"| `{code}` |", planTable);
        }

        Assert.Equal(15, codes.Count);
    }

    private static string FindPlanFile(string name)
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            var candidate = Path.Combine(dir.FullName, "plan", name);
            if (File.Exists(candidate))
            {
                return candidate;
            }

            dir = dir.Parent;
        }

        throw new FileNotFoundException($"找不到 plan/{name}");
    }
}
