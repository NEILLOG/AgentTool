namespace DocumentReader.Core.Models;

/// <summary>Word 讀取選項。預設只輸出內文；頁首頁尾、註腳、註解會附在文件最後的「附錄」區段。</summary>
public sealed record WordReadOptions(bool IncludeHeadersFooters = false, bool IncludeFootnotes = false, bool IncludeComments = false);

/// <param name="SectionId">用於 ReadSection；s0 是第一個標題之前的前言，其餘依文件順序為 s1、s2…。</param>
/// <param name="Level">標題層級 1 到 9；前言為 0。</param>
/// <param name="Chars">這一節自己的字元數（不含子節）。</param>
/// <param name="TotalChars">含所有子節的字元數。</param>
public sealed record OutlineNode(string SectionId, int Level, string Title, int Chars, int TotalChars, IReadOnlyList<OutlineNode> Children);

/// <param name="Sections">標題樹的最上層節點；有前言時第一個是 s0。</param>
/// <param name="HasHeadings">文件有沒有任何標題；沒有時請改用 Read 依字數分頁。</param>
/// <param name="Notes">轉換時略過或簡化的內容說明（目錄、圖片數量等）。</param>
public sealed record DocumentOutline(string Path, int TotalChars, bool HasHeadings, IReadOnlyList<OutlineNode> Sections, IReadOnlyList<string> Notes);

/// <param name="Markdown">這一段的 Markdown。</param>
/// <param name="StartOffset">這一段在讀取範圍（整份文件或某一節）內的起點字元位置。</param>
/// <param name="EndOffset">終點（不含）。</param>
/// <param name="TotalChars">讀取範圍的總字元數。</param>
/// <param name="Truncated">範圍內還有沒讀到的內容。</param>
/// <param name="NextOffset">被截斷時，下一次呼叫要帶的 offset。</param>
/// <param name="SectionId">讀的是哪一節（整份文件為 null）。</param>
public sealed record ReadResult(
    string Markdown,
    int StartOffset,
    int EndOffset,
    int TotalChars,
    bool Truncated,
    int? NextOffset,
    string? SectionId,
    IReadOnlyList<string> Notes);

/// <summary>PPT 讀取選項。演講者備註預設輸出（附在各投影片最後，以引用區塊呈現）。</summary>
public sealed record PptReadOptions(bool IncludeNotes = true);

/// <param name="Number">投影片編號（從 1 開始，依簡報順序）。</param>
/// <param name="Title">標題預留位置的文字；沒有標題為空字串。</param>
/// <param name="Chars">這張投影片轉成 Markdown 後的字元數（含備註）。</param>
public sealed record SlideOutline(int Number, string Title, int Chars, bool Hidden, bool HasNotes, int Pictures, int Tables, int Charts);

/// <param name="Notes">轉換時略過或簡化的內容說明（圖片數量、SmartArt 等）。</param>
public sealed record PresentationOutline(string Path, int SlideCount, int TotalChars, IReadOnlyList<SlideOutline> Slides, IReadOnlyList<string> Notes);

/// <summary>PDF 讀取選項。</summary>
/// <param name="UseOcr">需要 OCR 的頁面（掃描檔、亂碼頁）在有可用的 OCR 引擎時自動辨識；關閉則只標記。</param>
public sealed record PdfReadOptions(bool UseOcr = true);

/// <param name="Chars">文字層抽出的字元數（OCR 結果不計）。</param>
/// <param name="NeedsOcr">文字層不可用（沒有文字的掃描頁，或亂碼比例高）。</param>
/// <param name="Tables">用框線偵測到的表格數。</param>
public sealed record PdfPageOutline(int Number, int Chars, int Images, int Tables, bool NeedsOcr, string? Reason);

public sealed record PdfBookmark(string Title, int Level, int? Page);

public sealed record PdfOutline(string Path, int PageCount, int TotalChars, IReadOnlyList<PdfPageOutline> Pages, IReadOnlyList<PdfBookmark> Bookmarks, bool OcrAvailable, IReadOnlyList<string> Notes);

/// <param name="Png">PNG 圖片內容。</param>
public sealed record RenderedPage(int Page, int Width, int Height, byte[] Png);
