namespace ExcelTools.Core.Models;

public sealed record SheetSummary(int Index, string Name, string? UsedRange, bool Hidden);

/// <param name="WorkbookId">後續操作都帶這個 ID。</param>
/// <param name="ReadOnly">.xlsm 只能讀取；寫入與 Save 會被拒絕。</param>
/// <param name="AlreadyOpen">這個路徑先前已開啟，回傳的是既有 session。</param>
/// <param name="PreservationWarnings">ClosedXML 存檔時無法保留的內容；非空時 Save 會被拒絕，只能 SaveAs。</param>
public sealed record WorkbookInfo(
    string WorkbookId,
    string Path,
    bool ReadOnly,
    bool IsDirty,
    bool AlreadyOpen,
    IReadOnlyList<SheetSummary> Sheets,
    IReadOnlyList<string> PreservationWarnings);

/// <param name="BackupPath">覆寫前建立的備份；沒有覆寫既有檔案時為 null。</param>
/// <param name="PreservationWarnings">SaveAs 時，原檔中新檔不會包含的內容。</param>
public sealed record SaveResult(string Path, string? BackupPath, IReadOnlyList<string> PreservationWarnings);

/// <param name="ExpiryBackupFailed">閒置逾時時備份失敗，session 仍保留在記憶體中，請盡快存檔。</param>
public sealed record OpenWorkbookEntry(
    string WorkbookId,
    string Path,
    bool IsDirty,
    bool ReadOnly,
    DateTimeOffset LastAccessUtc,
    bool ExpiryBackupFailed);

public sealed record FileEntry(string Name, string Path, long SizeBytes, DateTimeOffset LastModifiedUtc);

public sealed record FileListing(string Directory, IReadOnlyList<FileEntry> Files, bool Truncated);

/// <param name="Visibility">Visible、Hidden 或 VeryHidden。</param>
/// <param name="MergedRanges">合併儲存格範圍，最多列出 <c>SheetOperations.MaxMergedRangesListed</c> 個。</param>
/// <param name="FrozenRows">凍結的列數（0 = 沒有凍結）。</param>
/// <param name="FrozenColumns">凍結的欄數（0 = 沒有凍結）。</param>
public sealed record SheetInfo(
    string Name,
    int Index,
    string Visibility,
    string? UsedRange,
    IReadOnlyList<string> MergedRanges,
    int MergedRangeCount,
    int FrozenRows,
    int FrozenColumns,
    IReadOnlyList<string> TableNames,
    bool HasAutoFilter);

/// <param name="Sheets">操作後的工作表清單（依分頁順序）。</param>
/// <param name="Warnings">需要使用者或 agent 留意的後果，例如刪除後有公式失效。</param>
public sealed record SheetChangeResult(IReadOnlyList<SheetSummary> Sheets, IReadOnlyList<string> Warnings);
