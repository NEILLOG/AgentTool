namespace DocumentReader.Core.Pdf;

/// <summary>OCR 引擎。Core 只認這個介面；Windows 實作放在另一個專案，沒有引擎時需要 OCR 的頁面只會被標記。</summary>
public interface IOcrEngine
{
    /// <summary>引擎是否可用；不可用時說明原因與安裝提示（例如缺少繁體中文語言套件）。</summary>
    bool IsAvailable(out string? reason);

    /// <summary>辨識一張 PNG 圖片，回傳純文字（一行一行，由上而下）。</summary>
    string Recognize(byte[] png);
}
