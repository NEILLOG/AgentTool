using System.Collections.Concurrent;
using System.Globalization;
using ClosedXML.Excel;
using OfficeTools.Common.Errors;
using OfficeTools.Common.Security;

namespace ExcelTools.Core.Workspace;

/// <summary>
/// 管理開啟中的活頁簿：ID 對應、每個 session 一把鎖、閒置逾時。
/// 沒有背景計時器：每次存取時順便清理，主程式也可定期呼叫 <see cref="SweepExpired"/>。
/// </summary>
public sealed class WorkbookSessionManager : IDisposable
{
    private const int MaxExpiredRecords = 100;

    private readonly ConcurrentDictionary<string, WorkbookSession> _sessions = new();
    private readonly ConcurrentDictionary<string, string> _expired = new();
    private readonly ConcurrentQueue<string> _expiredOrder = new();
    private readonly PathGuard _guard;
    private readonly ExcelToolsOptions _options;
    private readonly TimeProvider _time;
    private readonly StringComparison _pathComparison;

    public WorkbookSessionManager(PathGuard guard, ExcelToolsOptions options, TimeProvider? timeProvider = null)
    {
        ArgumentNullException.ThrowIfNull(guard);
        ArgumentNullException.ThrowIfNull(options);
        _guard = guard;
        _options = options;
        _time = timeProvider ?? TimeProvider.System;
        _pathComparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
    }

    /// <summary>
    /// 處理閒置逾時的 session：沒有未存變更的直接釋放；有未存變更的先存成備份再釋放，
    /// 備份失敗則保留在記憶體並標記。
    /// </summary>
    public void SweepExpired()
    {
        var now = _time.GetUtcNow();
        foreach (var session in _sessions.Values)
        {
            if (now - session.LastAccess < _options.IdleTimeout || !session.Lock.Wait(0))
            {
                continue;
            }

            try
            {
                if (session.Closed || now - session.LastAccess < _options.IdleTimeout)
                {
                    continue;
                }

                if (!session.IsDirty)
                {
                    Remove(session);
                    continue;
                }

                try
                {
                    var backup = WriteAutosave(session, now);
                    RecordExpired(session.Id, backup);
                    Remove(session);
                }
                catch (Exception ex) when (ex is OfficeToolException or IOException or UnauthorizedAccessException)
                {
                    session.ExpiryBackupFailed = true;
                    session.LastAccess = now; // 過一個逾時週期再重試，避免每次存取都重複失敗
                }
            }
            finally
            {
                session.Lock.Release();
            }
        }
    }

    public void Dispose()
    {
        foreach (var session in _sessions.Values)
        {
            session.Closed = true;
            session.Workbook.Dispose();
        }

        _sessions.Clear();
    }

    internal WorkbookSession Register(string fullPath, XLWorkbook workbook, bool readOnly, IReadOnlyList<string> warnings)
    {
        var id = $"wb_{Guid.NewGuid():N}"[..11];
        var session = new WorkbookSession(id, fullPath, workbook, readOnly, warnings, _time.GetUtcNow());
        _sessions[id] = session;
        return session;
    }

    internal WorkbookSession? FindByPath(string fullPath)
    {
        SweepExpired();
        return _sessions.Values.FirstOrDefault(s => !s.Closed && string.Equals(s.Path, fullPath, _pathComparison));
    }

    internal IReadOnlyList<WorkbookSession> Snapshot()
    {
        SweepExpired();
        return _sessions.Values.Where(s => !s.Closed).OrderBy(s => s.LastAccess).ToList();
    }

    /// <summary>取得鎖後執行；<paramref name="mutates"/> 為 true 時會檢查唯讀，並在結束後標記為有未存變更。</summary>
    internal T Use<T>(string workbookId, bool mutates, Func<WorkbookSession, T> action)
    {
        SweepExpired();
        var session = Find(workbookId);

        session.Lock.Wait();
        try
        {
            if (session.Closed)
            {
                throw NotFound(workbookId);
            }

            if (mutates && session.ReadOnly)
            {
                throw new OfficeToolException(
                    ErrorCodes.UnsupportedFormat,
                    "這個活頁簿是唯讀的（.xlsm 第一版只能讀取）",
                    "請用 save_as 另存為 .xlsx 後再編輯；巨集不會保留");
            }

            session.LastAccess = _time.GetUtcNow();
            try
            {
                return action(session);
            }
            finally
            {
                if (mutates)
                {
                    session.IsDirty = true; // 操作中途失敗時活頁簿可能已部分變更，一律視為有變更
                }

                session.LastAccess = _time.GetUtcNow(); // 閒置從操作結束算起，長時間的操作不算閒置
            }
        }
        finally
        {
            session.Lock.Release();
        }
    }

    internal void Use(string workbookId, bool mutates, Action<WorkbookSession> action) =>
        Use<object?>(workbookId, mutates, s =>
        {
            action(s);
            return null;
        });

    /// <summary>必須在持有該 session 的鎖時呼叫。</summary>
    internal void Remove(WorkbookSession session)
    {
        session.Closed = true;
        _sessions.TryRemove(session.Id, out _);
        session.Workbook.Dispose();
    }

    private WorkbookSession Find(string workbookId)
    {
        if (workbookId is not null && _sessions.TryGetValue(workbookId, out var session) && !session.Closed)
        {
            return session;
        }

        throw NotFound(workbookId);
    }

    private OfficeToolException NotFound(string? workbookId)
    {
        if (workbookId is not null && _expired.TryGetValue(workbookId, out var backup))
        {
            return new OfficeToolException(
                ErrorCodes.SessionExpired,
                $"活頁簿 {workbookId} 閒置逾時已釋放，未存檔的變更已備份到：{backup}",
                $"請告知使用者有未存檔的變更已備份，並用 open 開啟備份檔接續：{backup}");
        }

        return new OfficeToolException(
            ErrorCodes.SessionNotFound,
            $"找不到活頁簿 {workbookId}",
            "workbookId 不存在，請重新 open");
    }

    private string WriteAutosave(WorkbookSession session, DateTimeOffset now)
    {
        var dir = Path.GetDirectoryName(session.Path)!;
        var stem = Path.GetFileNameWithoutExtension(session.Path);
        var stamp = now.ToLocalTime().ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);

        var target = Path.Combine(dir, $"{stem}.autosave-{stamp}.xlsx");
        for (var i = 2; File.Exists(target); i++)
        {
            target = Path.Combine(dir, $"{stem}.autosave-{stamp}-{i}.xlsx");
        }

        _guard.ResolveAllowed(target);
        WorkbookPackage.WriteAtomically(session.Workbook, target);
        return target;
    }

    private void RecordExpired(string workbookId, string backupPath)
    {
        _expired[workbookId] = backupPath;
        _expiredOrder.Enqueue(workbookId);
        while (_expiredOrder.Count > MaxExpiredRecords && _expiredOrder.TryDequeue(out var oldest))
        {
            _expired.TryRemove(oldest, out _);
        }
    }
}
