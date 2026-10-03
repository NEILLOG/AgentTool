using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentReader.Core.Caching;
using DocumentReader.Core.Markdown;
using DocumentReader.Core.Models;
using OfficeTools.Common;
using OfficeTools.Common.Errors;
using OfficeTools.Common.Security;

namespace DocumentReader.Core.Ppt;

/// <summary>
/// 唯讀地把 .pptx 轉成 Markdown，以投影片為單位：先用 <see cref="GetOutline"/> 看每張投影片的標題與大小，
/// 再用 <see cref="ReadSlides"/> 讀指定範圍。解析結果依「路徑 + 最後修改時間 + 檔案大小」快取。
/// </summary>
public sealed class PptReader
{
    private const string Extension = ".pptx";

    private readonly PathGuard _guard;
    private readonly OfficeToolsOptions _options;
    private readonly DocumentReaderOptions _readerOptions;
    private readonly ParsedDocumentCache _cache;

    public PptReader(PathGuard guard, OfficeToolsOptions options, DocumentReaderOptions readerOptions, ParsedDocumentCache? cache = null)
    {
        ArgumentNullException.ThrowIfNull(guard);
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(readerOptions);
        _guard = guard;
        _options = options;
        _readerOptions = readerOptions;
        _cache = cache ?? new ParsedDocumentCache(readerOptions.CacheMaxEntries);
    }

    /// <summary>每張投影片的編號、標題、字元數，以及有沒有備註、圖片、表格、圖表。</summary>
    public PresentationOutline GetOutline(string path, PptReadOptions? pptOptions = null)
    {
        var (full, deck) = Load(path, pptOptions ?? new PptReadOptions());
        var slides = deck.Slides.Select(s => new SlideOutline(s.Number, s.Title, s.End - s.Start, s.Hidden, s.HasNotes, s.Pictures, s.Tables, s.Charts)).ToList();
        return new PresentationOutline(full, slides.Count, deck.Text.Length, slides, deck.Notes);
    }

    /// <summary>
    /// 讀第 <paramref name="from"/> 到 <paramref name="to"/> 張投影片（含）；<paramref name="to"/> 省略時只讀 <paramref name="from"/> 那一張。
    /// <paramref name="offset"/> 是在這個範圍內的字元位置，被截斷時用回傳的 NextOffset 續讀。
    /// </summary>
    public ReadResult ReadSlides(string path, int from, int? to = null, int offset = 0, int? maxChars = null, PptReadOptions? pptOptions = null)
    {
        var (_, deck) = Load(path, pptOptions ?? new PptReadOptions());
        var last = to ?? from;
        if (from < 1 || from > deck.Slides.Count || last < from || last > deck.Slides.Count)
        {
            throw new OfficeToolException(
                ErrorCodes.InvalidValue,
                $"投影片範圍 {from}–{last} 無效（簡報共 {deck.Slides.Count} 張）",
                deck.Slides.Count == 0 ? "簡報沒有投影片" : $"投影片編號從 1 到 {deck.Slides.Count}，且結束要大於等於開始");
        }

        return TextPaging.Slice(deck.Text, deck.Slides[from - 1].Start, deck.Slides[last - 1].End, offset, maxChars, _readerOptions.MaxCharsPerRead, sectionId: null, deck.Notes);
    }

    /// <summary>依字元數分頁讀整份簡報。</summary>
    public ReadResult Read(string path, int offset = 0, int? maxChars = null, PptReadOptions? pptOptions = null)
    {
        var (_, deck) = Load(path, pptOptions ?? new PptReadOptions());
        return TextPaging.Slice(deck.Text, 0, deck.Text.Length, offset, maxChars, _readerOptions.MaxCharsPerRead, sectionId: null, deck.Notes);
    }

    private (string FullPath, ParsedPresentation Deck) Load(string path, PptReadOptions pptOptions)
    {
        var full = _guard.ResolveExistingFile(path);
        if (!Path.GetExtension(full).Equals(Extension, StringComparison.OrdinalIgnoreCase))
        {
            throw new OfficeToolException(
                ErrorCodes.UnsupportedFormat,
                $"PPT 讀取工具只能讀 {Extension}：{Path.GetFileName(full)}",
                "Word / PDF / Excel 請用對應的工具；.ppt 請先用 PowerPoint 另存為 .pptx");
        }

        var info = new FileInfo(full);
        var key = new ParsedDocumentCache.Key(full, info.LastWriteTimeUtc.Ticks, info.Length, $"ppt:{pptOptions.IncludeNotes}");
        return (full, _cache.GetOrAdd(key, () => Parse(full, pptOptions)));
    }

    private ParsedPresentation Parse(string fullPath, PptReadOptions pptOptions)
    {
        var bytes = PackageChecks.ReadShared(fullPath);
        PackageChecks.ValidateZip(bytes, _options, "[Content_Types].xml", "ppt/presentation.xml");

        try
        {
            using var document = PresentationDocument.Open(
                new MemoryStream(bytes, writable: false),
                isEditable: false,
                new OpenSettings { MaxCharactersInPart = _options.MaxCharactersInPart });
            var main = document.PresentationPart ?? throw PackageChecks.Corrupt("簡報缺少主要內容部件", null);
            return new PptConverter(main, pptOptions).Convert();
        }
        catch (XmlException ex) when (ex.Message.Contains("MaxCharacters", StringComparison.Ordinal))
        {
            throw new OfficeToolException(ErrorCodes.FileTooLarge, "簡報內容超過單一部件的字元上限", "請先將簡報分割或縮小後再處理", ex);
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or XmlException or InvalidDataException or IOException or InvalidOperationException)
        {
            throw PackageChecks.Corrupt("簡報結構損壞，無法讀取", ex);
        }
    }
}
