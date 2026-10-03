// 這些測試刻意用同步等待：被測的鎖是同步的 SemaphoreSlim.Wait，要驗證的就是執行緒之間的阻塞行為
#pragma warning disable xUnit1031

using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Tests;

public sealed class SessionManagerTests : IDisposable
{
    private static readonly TimeSpan Idle = TimeSpan.FromMinutes(30);

    private readonly ExcelEnv _env = new(idleTimeout: Idle);

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    private string OpenBook(string name = "book.xlsx") => _env.Files.Open(_env.MakeWorkbook(name)).WorkbookId;

    // ---- 閒置逾時：沒有未存變更 ----

    [Fact]
    public void Clean_session_is_released_after_the_idle_timeout()
    {
        var id = OpenBook();

        _env.Time.Advance(Idle + TimeSpan.FromSeconds(1));

        Assert.Empty(_env.Files.ListOpen());
        Assert.Equal(ErrorCodes.SessionNotFound, Throws(() => _env.Cell(id, "A1")).Code);
        Assert.Empty(Directory.GetFiles(_env.Root, "*.autosave-*"));
    }

    [Fact]
    public void Session_just_inside_the_timeout_stays_open()
    {
        var id = OpenBook();
        _env.Time.Advance(Idle - TimeSpan.FromSeconds(1));
        Assert.Equal("name", _env.Cell(id, "A1"));
    }

    [Fact]
    public void Activity_resets_the_idle_clock()
    {
        var id = OpenBook();

        _env.Time.Advance(TimeSpan.FromMinutes(20));
        Assert.Equal("name", _env.Cell(id, "A1")); // 存取一次
        _env.Time.Advance(TimeSpan.FromMinutes(20)); // 距離開檔 40 分鐘，距離上次存取 20 分鐘

        Assert.Equal("name", _env.Cell(id, "A1"));
    }

    // ---- 閒置逾時：有未存變更，先備份再釋放 ----

    [Fact]
    public void Dirty_session_is_backed_up_then_released_and_reports_SESSION_EXPIRED_with_the_path()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        var originalBytes = File.ReadAllBytes(path);
        var id = _env.Files.Open(path).WorkbookId;
        _env.SetCell(id, "A2", "unsaved edit");

        _env.Time.Advance(Idle + TimeSpan.FromMinutes(1));
        Assert.Empty(_env.Files.ListOpen());

        var backup = Assert.Single(Directory.GetFiles(_env.Root, "book.autosave-*.xlsx"));
        Assert.Equal("unsaved edit", ExcelEnv.ReadFromDisk(backup, "A2"));
        Assert.Equal(originalBytes, File.ReadAllBytes(path)); // 原檔不動

        var ex = Throws(() => _env.Cell(id, "A1"));
        Assert.Equal(ErrorCodes.SessionExpired, ex.Code);
        Assert.Contains(Path.GetFileName(backup), ex.Hint);
        Assert.Contains("open", ex.Hint);

        var reopened = _env.Files.Open(backup);
        Assert.Equal("unsaved edit", _env.Cell(reopened.WorkbookId, "A2"));
    }

    [Fact]
    public void Backup_of_a_workbook_with_lossy_content_is_still_written_as_a_new_file()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        ExcelEnv.AddEntries(path, ("xl/charts/chart1.xml", "<x/>"));
        var id = _env.Files.Open(path).WorkbookId;
        _env.SetCell(id, "A2", "edit");

        _env.Time.Advance(Idle + TimeSpan.FromMinutes(1));
        _env.Files.ListOpen();

        Assert.Equal(ErrorCodes.SessionExpired, Throws(() => _env.Cell(id, "A1")).Code);
        Assert.Single(Directory.GetFiles(_env.Root, "book.autosave-*.xlsx"));
    }

    [Fact]
    public void Failed_backup_keeps_the_session_in_memory_and_flags_it_then_retries_later()
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        var id = OpenBook();
        _env.SetCell(id, "A2", "precious");

        var mode = File.GetUnixFileMode(_env.Root);
        File.SetUnixFileMode(_env.Root, UnixFileMode.UserRead | UnixFileMode.UserExecute);
        try
        {
            _env.Time.Advance(Idle + TimeSpan.FromMinutes(1));

            var entry = Assert.Single(_env.Files.ListOpen());
            Assert.True(entry.ExpiryBackupFailed);
            Assert.True(entry.IsDirty);
            Assert.Equal("precious", _env.Cell(id, "A2")); // 內容還在，仍可使用

            File.SetUnixFileMode(_env.Root, mode);
            _env.Time.Advance(Idle + TimeSpan.FromMinutes(1)); // 再過一個逾時週期
            Assert.Empty(_env.Files.ListOpen());
        }
        finally
        {
            File.SetUnixFileMode(_env.Root, mode);
        }

        var backup = Assert.Single(Directory.GetFiles(_env.Root, "book.autosave-*.xlsx"));
        Assert.Equal("precious", ExcelEnv.ReadFromDisk(backup, "A2"));
    }

    [Fact]
    public void Successful_save_clears_the_failed_backup_flag_and_prevents_a_backup()
    {
        var path = _env.MakeWorkbook("book.xlsx");
        var id = _env.Files.Open(path).WorkbookId;
        _env.SetCell(id, "A2", "saved");
        _env.Files.Save(id);

        _env.Time.Advance(Idle + TimeSpan.FromMinutes(1));

        Assert.Empty(_env.Files.ListOpen());
        Assert.Empty(Directory.GetFiles(_env.Root, "*.autosave-*"));
        Assert.Equal(ErrorCodes.SessionNotFound, Throws(() => _env.Cell(id, "A1")).Code);
    }

    [Fact]
    public void Two_expiries_in_the_same_second_get_distinct_backup_files()
    {
        var a = _env.MakeWorkbook("a.xlsx");
        var idA = _env.Files.Open(a).WorkbookId;
        _env.SetCell(idA, "A2", "first");
        _env.Time.Advance(Idle + TimeSpan.FromMinutes(1));
        _env.Files.ListOpen(); // 第一次逾時、備份、釋放

        var idA2 = _env.Files.Open(a).WorkbookId;
        _env.SetCell(idA2, "A2", "second");
        _env.Time.Advance(Idle + TimeSpan.FromMinutes(1));
        _env.Files.ListOpen();

        Assert.Equal(2, Directory.GetFiles(_env.Root, "a.autosave-*.xlsx").Length);
    }

    [Fact]
    public void Sweeping_while_a_session_is_in_use_does_not_release_it()
    {
        var id = OpenBook();

        var cell = _env.Sessions.Use(id, false, s =>
        {
            _env.Time.Advance(Idle * 2);
            _env.Sessions.SweepExpired(); // 同一執行緒、鎖被持有：必須略過這個 session
            return Internal.CellValueConverter.ToSerializable(s.Workbook.Worksheet(1).Cell("A1").Value);
        });

        Assert.Equal("name", cell);
        Assert.Single(_env.Files.ListOpen());
    }

    [Fact]
    public void Only_the_most_recent_expired_records_are_remembered()
    {
        var ids = new List<string>();
        for (var i = 0; i < 105; i++)
        {
            var id = _env.Files.Open(_env.MakeWorkbook($"b{i:000}.xlsx")).WorkbookId;
            _env.SetCell(id, "A1", $"edit {i}");
            ids.Add(id);
        }

        _env.Time.Advance(Idle + TimeSpan.FromMinutes(1));
        _env.Files.ListOpen();

        Assert.Equal(105, Directory.GetFiles(_env.Root, "*.autosave-*").Length); // 備份檔都在，只是忘了 ID 對應

        var codes = ids.Select(id => Throws(() => _env.Cell(id, "A1")).Code).ToList();
        Assert.Equal(100, codes.Count(c => c == ErrorCodes.SessionExpired));
        Assert.Equal(5, codes.Count(c => c == ErrorCodes.SessionNotFound));
    }

    [Fact]
    public void A_long_running_operation_does_not_count_as_idle_time()
    {
        var id = OpenBook();

        _env.Sessions.Use(id, false, _ => _env.Time.Advance(Idle * 2)); // 操作本身跑了很久

        Assert.Single(_env.Files.ListOpen());
        Assert.Equal("name", _env.Cell(id, "A1"));
    }

    // ---- 並行 ----

    [Fact]
    public void Concurrent_operations_on_one_session_never_overlap_and_lose_no_updates()
    {
        var id = OpenBook();
        _env.SetCell(id, "C1", 0);
        var inside = 0;
        var maxInside = 0;

        Parallel.For(0, 64, new ParallelOptions { MaxDegreeOfParallelism = 16 }, _ =>
            _env.Sessions.Use(id, true, s =>
            {
                var now = Interlocked.Increment(ref inside);
                InterlockedMax(ref maxInside, now);
                var cell = s.Workbook.Worksheet(1).Cell("C1");
                var value = cell.Value.GetNumber();
                Thread.Sleep(1);
                cell.Value = value + 1;
                Interlocked.Decrement(ref inside);
            }));

        Assert.Equal(1, maxInside);
        Assert.Equal(64.0, _env.Cell(id, "C1"));
    }

    [Fact]
    public void Different_sessions_do_not_block_each_other()
    {
        var a = OpenBook("a.xlsx");
        var b = OpenBook("b.xlsx");

        var bRanWhileAWasHeld = false;
        _env.Sessions.Use(a, false, _ =>
        {
            var task = Task.Run(() => _env.Sessions.Use(b, false, _ => bRanWhileAWasHeld = true));
            Assert.True(task.Wait(TimeSpan.FromSeconds(5)));
        });

        Assert.True(bRanWhileAWasHeld);
    }

    [Fact]
    public void Closing_while_another_thread_waits_makes_the_waiter_see_SESSION_NOT_FOUND()
    {
        var id = OpenBook();
        Task<OfficeToolException?>? waiter = null;

        _env.Sessions.Use(id, false, _ =>
        {
            waiter = Task.Run(() =>
            {
                try
                {
                    _env.Cell(id, "A1");
                    return null;
                }
                catch (OfficeToolException ex)
                {
                    return ex;
                }
            });
            Thread.Sleep(100); // 讓等待者排進鎖的佇列
            _env.Sessions.Remove(_env.Sessions.Snapshot().Single(s => s.Id == id));
        });

        var ex = waiter!.GetAwaiter().GetResult();
        Assert.Equal(ErrorCodes.SessionNotFound, ex?.Code);
    }

    private static void InterlockedMax(ref int location, int value)
    {
        int current;
        do
        {
            current = Volatile.Read(ref location);
            if (value <= current)
            {
                return;
            }
        }
        while (Interlocked.CompareExchange(ref location, value, current) != current);
    }
}
