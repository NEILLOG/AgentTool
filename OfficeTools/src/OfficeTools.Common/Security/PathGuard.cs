using OfficeTools.Common.Errors;

namespace OfficeTools.Common.Security;

/// <summary>
/// 所有接收路徑的方法第一行都要過這裡：正規化、解析 symlink、限制在 AllowedRoots 內，
/// 並檢查副檔名、檔案大小與覆寫政策。
/// </summary>
public sealed class PathGuard
{
    private const long BytesPerMb = 1024L * 1024;
    private const int MaxLinkDepth = 16;

    private readonly OfficeToolsOptions _options;
    private readonly TimeProvider _time;
    private readonly StringComparison _comparison;
    private readonly string[] _roots;
    private readonly HashSet<string> _extensions;

    public PathGuard(OfficeToolsOptions options, TimeProvider? timeProvider = null)
    {
        ArgumentNullException.ThrowIfNull(options);
        _options = options;
        _time = timeProvider ?? TimeProvider.System;
        _comparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        _extensions = new HashSet<string>(options.AllowedExtensions, StringComparer.OrdinalIgnoreCase);
        // 根目錄也要解析 symlink（例如 macOS 的 /var → /private/var），才能和解析後的路徑比對
        _roots = options.AllowedRoots
            .Where(r => !string.IsNullOrWhiteSpace(r))
            .Select(r => ResolveLinks(Path.GetFullPath(r), 0))
            .ToArray();
    }

    /// <summary>檢查路徑（檔案或資料夾，不必存在）在允許範圍內，回傳解析後的完整路徑。</summary>
    public string ResolveAllowed(string path)
    {
        RejectSuspicious(path);
        var full = ResolveLinks(Path.GetFullPath(path), 0);
        if (!IsUnderAnyRoot(full))
        {
            throw NotAllowed($"路徑不在允許的範圍內：{path}");
        }

        return full;
    }

    /// <summary>讀取用：路徑允許、副檔名在白名單、檔案存在且未超過大小上限。</summary>
    public string ResolveExistingFile(string path)
    {
        var full = ResolveAllowed(path);
        ValidateExtension(full);

        var info = new FileInfo(full);
        if (!info.Exists)
        {
            throw new OfficeToolException(
                ErrorCodes.FileNotFound,
                $"找不到檔案：{path}",
                "請確認路徑，或先列出資料夾內容查看實際檔名");
        }

        if (info.Length > _options.MaxFileSizeMb * BytesPerMb)
        {
            throw new OfficeToolException(
                ErrorCodes.FileTooLarge,
                $"檔案大小 {info.Length / BytesPerMb} MB，超過上限 {_options.MaxFileSizeMb} MB：{path}",
                "請先將檔案分割或縮小後再處理");
        }

        return full;
    }

    /// <summary>寫入新檔用（Create / SaveAs）：路徑允許、副檔名在白名單，既有檔案依 AllowOverwrite 決定。</summary>
    public string ResolveWriteTarget(string path)
    {
        var full = ResolveAllowed(path);
        ValidateExtension(full);

        if (File.Exists(full) && !_options.AllowOverwrite)
        {
            throw new OfficeToolException(
                ErrorCodes.FileExists,
                $"檔案已存在：{path}",
                "請改用其他檔名，或明確要求覆寫");
        }

        return full;
    }

    /// <summary>
    /// 把既有檔案複製成同資料夾的 <c>名稱.backup-時間戳.副檔名</c>，回傳備份路徑。
    /// 是否需要備份由呼叫端決定（Excel Save 一律備份；SaveAs 依 BackupOnOverwrite）。
    /// </summary>
    public string CreateBackup(string fullPath)
    {
        var dir = Path.GetDirectoryName(fullPath)!;
        var stem = Path.GetFileNameWithoutExtension(fullPath);
        var ext = Path.GetExtension(fullPath);
        var stamp = _time.GetLocalNow().ToString("yyyyMMdd-HHmmss", System.Globalization.CultureInfo.InvariantCulture);

        var backup = Path.Combine(dir, $"{stem}.backup-{stamp}{ext}");
        for (var i = 2; File.Exists(backup); i++)
        {
            backup = Path.Combine(dir, $"{stem}.backup-{stamp}-{i}{ext}");
        }

        try
        {
            File.Copy(fullPath, backup);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new OfficeToolException(
                ErrorCodes.FileLocked,
                $"無法建立備份檔：{backup}",
                "請確認資料夾有寫入權限且空間足夠；為了不遺失原檔，這次操作已中止",
                ex);
        }

        return backup;
    }

    private void ValidateExtension(string fullPath)
    {
        var ext = Path.GetExtension(fullPath);
        if (!_extensions.Contains(ext))
        {
            throw new OfficeToolException(
                ErrorCodes.UnsupportedFormat,
                $"不支援的檔案格式「{ext}」：{Path.GetFileName(fullPath)}",
                $"支援的格式：{string.Join("、", _extensions.Order(StringComparer.OrdinalIgnoreCase))}。.xls / .doc / .ppt 請先用 Office 另存為新格式");
        }
    }

    private static void RejectSuspicious(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || path.Contains('\0'))
        {
            throw new OfficeToolException(ErrorCodes.PathNotAllowed, "路徑是空的或含有無效字元", "請提供完整的檔案路徑");
        }

        // UNC（\\server\share）與 \\?\ \\.\ 前綴
        if (path.StartsWith(@"\\", StringComparison.Ordinal) || path.StartsWith("//", StringComparison.Ordinal))
        {
            throw new OfficeToolException(ErrorCodes.PathNotAllowed, $"不允許網路路徑或裝置路徑：{path}", "請使用本機資料夾內的路徑");
        }

        // 除了開頭的磁碟機代號（C:），路徑中不應出現冒號（NTFS 替代資料流 file.xlsx:stream）
        var rest = path.Length >= 2 && char.IsAsciiLetter(path[0]) && path[1] == ':' ? path[2..] : path;
        if (rest.Contains(':'))
        {
            throw new OfficeToolException(ErrorCodes.PathNotAllowed, $"路徑含有不允許的冒號：{path}", "請移除路徑中的冒號（替代資料流不受支援）");
        }
    }

    /// <summary>逐層展開路徑中的 symlink / junction，讓比對根目錄時看到的是實際位置。路徑不必存在。</summary>
    private static string ResolveLinks(string fullPath, int depth)
    {
        if (depth > MaxLinkDepth)
        {
            throw new OfficeToolException(ErrorCodes.PathNotAllowed, $"符號連結層數過深：{fullPath}", "請使用實際的檔案路徑");
        }

        var root = Path.GetPathRoot(fullPath) ?? string.Empty;
        var current = root;
        var segments = fullPath[root.Length..].Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);

        foreach (var segment in segments)
        {
            current = Path.Combine(current, segment);
            FileSystemInfo info = Directory.Exists(current) ? new DirectoryInfo(current) : new FileInfo(current);
            if (info.LinkTarget is null)
            {
                continue;
            }

            FileSystemInfo? target;
            try
            {
                target = info.ResolveLinkTarget(returnFinalTarget: true);
            }
            catch (IOException ex)
            {
                throw new OfficeToolException(ErrorCodes.PathNotAllowed, $"無法解析符號連結：{current}", "請使用實際的檔案路徑", ex);
            }

            if (target is not null)
            {
                current = ResolveLinks(target.FullName, depth + 1);
            }
        }

        return current;
    }

    private bool IsUnderAnyRoot(string fullPath)
    {
        foreach (var root in _roots)
        {
            var trimmed = Path.TrimEndingDirectorySeparator(root);
            var prefix = trimmed.EndsWith(Path.DirectorySeparatorChar) ? trimmed : trimmed + Path.DirectorySeparatorChar;
            if (fullPath.Equals(trimmed, _comparison) || fullPath.StartsWith(prefix, _comparison))
            {
                return true;
            }
        }

        return false;
    }

    private OfficeToolException NotAllowed(string message) =>
        new(
            ErrorCodes.PathNotAllowed,
            message,
            _roots.Length == 0
                ? "尚未設定任何允許存取的資料夾，請先在設定中加入 AllowedRoots"
                : $"允許的根目錄：{string.Join("、", _roots)}");
}
