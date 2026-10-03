namespace DocumentReader.Core;

public sealed class DocumentReaderOptions
{
    /// <summary>單次讀取最多回傳的字元數；超過就截斷並回傳續讀位置。</summary>
    public int MaxCharsPerRead { get; init; } = 20_000;

    /// <summary>PDF 頁面轉圖片 / OCR 的解析度。</summary>
    public int OcrDpi { get; init; } = 200;

    /// <summary>解析結果的記憶體快取筆數上限（LRU）。</summary>
    public int CacheMaxEntries { get; init; } = 20;
}
