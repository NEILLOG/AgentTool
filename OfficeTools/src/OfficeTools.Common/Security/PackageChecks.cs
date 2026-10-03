using System.IO.Compression;
using OfficeTools.Common.Errors;

namespace OfficeTools.Common.Security;

/// <summary>Office 檔案（xlsx / docx / pptx 都是 zip 封裝）共用的讀檔與結構檢查。</summary>
public static class PackageChecks
{
    private static readonly byte[] OleSignature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    private static readonly byte[] ZipSignature = [0x50, 0x4B, 0x03, 0x04];

    /// <summary>讀進記憶體後立即放開檔案，使用者可繼續在 Office 中開啟與編輯。</summary>
    public static byte[] ReadShared(string fullPath)
    {
        try
        {
            using var fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var ms = new MemoryStream();
            fs.CopyTo(ms);
            return ms.ToArray();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new OfficeToolException(
                ErrorCodes.FileLocked,
                $"檔案無法存取（被其他程式鎖定或沒有權限）：{fullPath}",
                "檔案在 Office 中開啟，請關閉或改用 save_as",
                ex);
        }
    }

    /// <summary>
    /// 驗證是結構完整的 zip 封裝：不是加密檔（OLE 複合檔）、解壓縮後大小在上限內、包含必要的部件。
    /// 回傳所有部件名稱（供呼叫端偵測特定內容）。
    /// </summary>
    public static IReadOnlySet<string> ValidateZip(byte[] bytes, OfficeToolsOptions options, params string[] requiredParts)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        ArgumentNullException.ThrowIfNull(options);

        if (bytes.AsSpan().StartsWith(OleSignature))
        {
            throw new OfficeToolException(
                ErrorCodes.PasswordProtected,
                "檔案有密碼保護，或是舊版格式（.xls / .doc / .ppt）",
                "請使用者先在 Office 移除密碼，或另存為新格式（.xlsx / .docx / .pptx）後再處理");
        }

        if (!bytes.AsSpan().StartsWith(ZipSignature))
        {
            throw Corrupt("檔案不是有效的 Office 檔案（不是 zip 封裝）", null);
        }

        try
        {
            using var zip = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);

            var total = zip.Entries.Sum(e => e.Length);
            if (total > options.MaxUncompressedMb * 1024L * 1024)
            {
                throw new OfficeToolException(
                    ErrorCodes.FileTooLarge,
                    $"檔案解壓縮後約 {total / (1024 * 1024)} MB，超過上限 {options.MaxUncompressedMb} MB",
                    "請先將檔案分割或縮小後再處理");
            }

            var names = zip.Entries.Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (requiredParts.Any(p => !names.Contains(p)))
            {
                throw Corrupt("檔案缺少必要的部件，不是有效的 Office 檔案", null);
            }

            return names;
        }
        catch (InvalidDataException ex)
        {
            throw Corrupt("檔案結構損壞，無法讀取", ex);
        }
    }

    public static OfficeToolException Corrupt(string message, Exception? inner) =>
        new(ErrorCodes.CorruptFile, message, "請用 Office 開啟並修復後另存新檔", inner);
}
