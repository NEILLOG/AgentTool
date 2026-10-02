namespace OfficeTools.Common;

public sealed class OfficeToolsOptions
{
    /// <summary>允許存取的根目錄。空 = 全部拒絕。</summary>
    public IReadOnlyList<string> AllowedRoots { get; init; } = [];

    /// <summary>副檔名白名單（含句點，不分大小寫）。</summary>
    public IReadOnlyList<string> AllowedExtensions { get; init; } = [".xlsx", ".xlsm", ".docx", ".pptx", ".pdf"];

    public int MaxFileSizeMb { get; init; } = 50;

    /// <summary>單一 XML 部件解壓後的字元上限（套給 Open XML SDK 的 OpenSettings.MaxCharactersInPart）。</summary>
    public long MaxCharactersInPart { get; init; } = 50_000_000;

    /// <summary>只管 Create / SaveAs 撞到既有檔案的情況；Excel Save 覆寫原檔不受此控制，但一律備份。</summary>
    public bool AllowOverwrite { get; init; }

    public bool BackupOnOverwrite { get; init; } = true;
}
