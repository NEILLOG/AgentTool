using ClosedXML.Excel;
using ExcelTools.Core.Internal;
using ExcelTools.Core.Models;
using ExcelTools.Core.Workspace;
using OfficeTools.Common;
using OfficeTools.Common.Errors;
using OfficeTools.Common.Security;

namespace ExcelTools.Core.Operations;

/// <summary>檔案與 session 層級的操作：列檔、建立、開啟、存檔、另存、關閉。</summary>
public sealed class FileOperations(PathGuard guard, OfficeToolsOptions options, WorkbookSessionManager sessions)
{
    private const string NewFileExtension = ".xlsx";

    private static readonly string[] ReadableExtensions = [".xlsx", ".xlsm"];

    public FileListing ListFiles(string directory, bool recursive = false, int maxEntries = 200)
    {
        var dir = guard.ResolveAllowed(directory);
        if (!Directory.Exists(dir))
        {
            throw new OfficeToolException(ErrorCodes.FileNotFound, $"找不到資料夾：{directory}", "請確認資料夾路徑");
        }

        var enumeration = new EnumerationOptions { RecurseSubdirectories = recursive, IgnoreInaccessible = true };
        var files = new List<FileEntry>();
        var truncated = false;

        foreach (var file in Directory.EnumerateFiles(dir, "*", enumeration).Order(StringComparer.OrdinalIgnoreCase))
        {
            var name = Path.GetFileName(file);
            if (name.StartsWith("~$", StringComparison.Ordinal)
                || !ReadableExtensions.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase))
            {
                continue; // 略過 Office 的鎖定檔與非 Excel 檔案
            }

            string resolved;
            try
            {
                resolved = guard.ResolveAllowed(file); // 指向允許範圍外的 symlink 不列出
            }
            catch (OfficeToolException)
            {
                continue;
            }

            if (files.Count >= maxEntries)
            {
                truncated = true;
                break;
            }

            var info = new FileInfo(resolved);
            files.Add(new FileEntry(name, file, info.Length, new DateTimeOffset(info.LastWriteTimeUtc, TimeSpan.Zero)));
        }

        return new FileListing(dir, files, truncated);
    }

    public WorkbookInfo Open(string path)
    {
        var full = guard.ResolveExistingFile(path);
        var extension = Path.GetExtension(full);
        if (!ReadableExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
        {
            throw new OfficeToolException(
                ErrorCodes.UnsupportedFormat,
                $"Excel 工具只能開啟 .xlsx / .xlsm：{Path.GetFileName(full)}",
                "Word、PPT、PDF 請改用文件讀取工具；.xls 請先用 Excel 另存為 .xlsx");
        }

        var existing = sessions.FindByPath(full);
        if (existing is not null)
        {
            return sessions.Use(existing.Id, mutates: false, s => Describe(s, alreadyOpen: true));
        }

        var bytes = WorkbookPackage.ReadShared(full);
        var warnings = WorkbookPackage.Inspect(bytes, options);
        var workbook = WorkbookPackage.Load(bytes);

        var session = sessions.Register(
            full,
            workbook,
            readOnly: extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase),
            warnings);
        return sessions.Use(session.Id, mutates: false, s => Describe(s, alreadyOpen: false));
    }

    public WorkbookInfo Create(string path, string? sheetName = null)
    {
        var full = ResolveNewWorkbookPath(path);
        var name = SheetNameRules.Validate(sheetName ?? "Sheet1");

        if (File.Exists(full) && options.BackupOnOverwrite)
        {
            guard.CreateBackup(full);
        }

        byte[] written;
        using (var created = new XLWorkbook())
        {
            created.AddWorksheet(name);
            written = WorkbookPackage.WriteAtomically(created, full);
        }

        var session = sessions.Register(full, WorkbookPackage.Load(written), readOnly: false, []);
        return sessions.Use(session.Id, mutates: false, s => Describe(s, alreadyOpen: false));
    }

    /// <summary>覆寫原檔。含 ClosedXML 無法保留內容的活頁簿會被拒絕；第一次存檔前會備份原檔。</summary>
    public SaveResult Save(string workbookId) =>
        sessions.Use(workbookId, mutates: false, s =>
        {
            if (s.ReadOnly)
            {
                throw new OfficeToolException(
                    ErrorCodes.UnsupportedFormat,
                    "這個活頁簿是唯讀的（.xlsm 第一版只能讀取）",
                    "請用 save_as 另存為 .xlsx");
            }

            if (s.PreservationWarnings.Count > 0)
            {
                throw new OfficeToolException(
                    ErrorCodes.UnsafeToOverwrite,
                    $"原檔含有無法保留的內容（{string.Join("、", s.PreservationWarnings)}），覆寫會讓這些內容遺失",
                    "請改用 save_as 存成新檔，原檔不會被動到");
            }

            if (s.BackupPath is null && File.Exists(s.Path))
            {
                s.BackupPath = guard.CreateBackup(s.Path);
            }

            s.ReloadFrom(WorkbookPackage.WriteAtomically(s.Workbook, s.Path));
            s.IsDirty = false;
            s.ExpiryBackupFailed = false;
            return new SaveResult(s.Path, s.BackupPath, []);
        });

    /// <summary>另存為新的 .xlsx，session 之後指向新檔。</summary>
    public SaveResult SaveAs(string workbookId, string path)
    {
        var full = ResolveNewWorkbookPath(path);

        var other = sessions.FindByPath(full);
        if (other is not null && other.Id != workbookId)
        {
            throw new OfficeToolException(
                ErrorCodes.FileExists,
                $"目標檔已在另一個活頁簿（{other.Id}）中開啟：{path}",
                "請改用其他檔名，或先關閉那個活頁簿");
        }

        return sessions.Use(workbookId, mutates: false, s =>
        {
            string? backup = null;
            if (File.Exists(full) && options.BackupOnOverwrite)
            {
                backup = guard.CreateBackup(full);
            }

            var lost = s.PreservationWarnings;
            s.ReloadFrom(WorkbookPackage.WriteAtomically(s.Workbook, full));

            s.Path = full;
            s.ReadOnly = false;
            s.PreservationWarnings = [];
            s.BackupPath = null;
            s.IsDirty = false;
            s.ExpiryBackupFailed = false;
            return new SaveResult(full, backup, lost);
        });
    }

    public void Close(string workbookId, bool discardChanges = false) =>
        sessions.Use(workbookId, mutates: false, s =>
        {
            if (s.IsDirty && !discardChanges)
            {
                throw new OfficeToolException(
                    ErrorCodes.UnsavedChanges,
                    $"活頁簿 {s.Id} 有未存檔的變更",
                    "請先 save / save_as，或帶 discardChanges 放棄變更後關閉");
            }

            sessions.Remove(s);
        });

    public IReadOnlyList<OpenWorkbookEntry> ListOpen() =>
        sessions.Snapshot()
            .Select(s => new OpenWorkbookEntry(s.Id, s.Path, s.IsDirty, s.ReadOnly, s.LastAccess, s.ExpiryBackupFailed))
            .ToList();

    private string ResolveNewWorkbookPath(string path)
    {
        var full = guard.ResolveWriteTarget(path);
        if (!Path.GetExtension(full).Equals(NewFileExtension, StringComparison.OrdinalIgnoreCase))
        {
            throw new OfficeToolException(
                ErrorCodes.UnsupportedFormat,
                $"新檔只能存成 {NewFileExtension}：{Path.GetFileName(full)}",
                $"請把副檔名改成 {NewFileExtension}");
        }

        if (!Directory.Exists(Path.GetDirectoryName(full)))
        {
            throw new OfficeToolException(ErrorCodes.FileNotFound, $"資料夾不存在：{Path.GetDirectoryName(full)}", "請改存到既有的資料夾");
        }

        return full;
    }

    private static WorkbookInfo Describe(WorkbookSession s, bool alreadyOpen) =>
        new(s.Id, s.Path, s.ReadOnly, s.IsDirty, alreadyOpen, SheetNameRules.Summarize(s.Workbook), s.PreservationWarnings);
}
