using ClosedXML.Excel;

namespace ExcelTools.Core.Workspace;

/// <summary>一個開啟中的活頁簿。ClosedXML 不是執行緒安全，所有存取都必須持有 <see cref="Lock"/>。</summary>
internal sealed class WorkbookSession(string id, string path, XLWorkbook workbook, bool readOnly, IReadOnlyList<string> preservationWarnings, DateTimeOffset now)
{
    public string Id { get; } = id;

    public string Path { get; set; } = path;

    public XLWorkbook Workbook { get; private set; } = workbook;

    /// <summary>存檔後換成用剛寫出的內容重新載入的活頁簿（原因見 <see cref="WorkbookPackage.WriteAtomically"/>）。</summary>
    public void ReloadFrom(byte[] savedContent)
    {
        var fresh = WorkbookPackage.Load(savedContent);
        Workbook.Dispose();
        Workbook = fresh;
    }

    public SemaphoreSlim Lock { get; } = new(1, 1);

    public bool ReadOnly { get; set; } = readOnly;

    public IReadOnlyList<string> PreservationWarnings { get; set; } = preservationWarnings;

    public bool IsDirty { get; set; }

    public DateTimeOffset LastAccess { get; set; } = now;

    /// <summary>此 session 第一次 Save 前為原檔建立的備份；之後的 Save 不再重複備份。</summary>
    public string? BackupPath { get; set; }

    public bool ExpiryBackupFailed { get; set; }

    public volatile bool Closed;
}
