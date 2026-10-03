using System.Collections.Concurrent;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentReader.Core.Caching;
using DocumentReader.Core.Markdown;
using DocumentReader.Core.Models;
using OfficeTools.Common;
using OfficeTools.Common.Errors;
using OfficeTools.Common.Security;
using PDFtoImage;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Exceptions;

namespace DocumentReader.Core.Pdf;

/// <summary>
/// 唯讀地把 .pdf 轉成 Markdown，以頁為單位：<see cref="GetOutline"/> 看每頁字數與是否需要 OCR，
/// <see cref="ReadPages"/> 讀指定頁，<see cref="RenderPages"/> 把頁面轉成圖片（給多模態模型看複雜版面）。
/// 文字層用 PdfPig 重建表格（依框線）與閱讀順序；沒有文字層的頁面交給 <see cref="IOcrEngine"/>，沒有引擎就只標記。
/// </summary>
public sealed partial class PdfReader
{
    private const string Extension = ".pdf";
    private const double MinTextChars = 15; // 圖片佔滿整頁卻少於這個字數 → 視為掃描頁
    private const double MaxGarbageRatio = 0.3;
    private const int MaxRenderPages = 10;
    private const int MaxRenderPixels = 4000;

    private sealed class Parsed
    {
        public required IReadOnlyList<PageResult> Pages { get; init; }

        public required IReadOnlyList<PdfBookmark> Bookmarks { get; init; }

        public required IReadOnlyList<string> Notes { get; init; }

        public ConcurrentDictionary<int, string> Ocr { get; } = new();
    }

    private sealed record PageResult(int Number, string Markdown, int Chars, int Images, int Tables, bool NeedsOcr, string? Reason);

    private readonly PathGuard _guard;
    private readonly OfficeToolsOptions _options;
    private readonly DocumentReaderOptions _readerOptions;
    private readonly IOcrEngine? _ocr;
    private readonly ParsedDocumentCache _cache;

    public PdfReader(PathGuard guard, OfficeToolsOptions options, DocumentReaderOptions readerOptions, IOcrEngine? ocr = null, ParsedDocumentCache? cache = null)
    {
        ArgumentNullException.ThrowIfNull(guard);
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(readerOptions);
        _guard = guard;
        _options = options;
        _readerOptions = readerOptions;
        _ocr = ocr;
        _cache = cache ?? new ParsedDocumentCache(readerOptions.CacheMaxEntries);
    }

    private bool OcrAvailable => _ocr is not null && _ocr.IsAvailable(out _);

    public PdfOutline GetOutline(string path)
    {
        var (full, parsed) = Load(path);
        var notes = parsed.Notes.ToList();
        var needs = parsed.Pages.Count(p => p.NeedsOcr);
        if (needs > 0)
        {
            notes.Add(OcrAvailable ? $"有 {needs} 頁沒有可用的文字層，讀取時會用 OCR 辨識" : $"有 {needs} 頁沒有可用的文字層（掃描檔或亂碼），目前沒有可用的 OCR 引擎；可用 RenderPages 轉成圖片交給多模態模型");
        }

        return new PdfOutline(
            full,
            parsed.Pages.Count,
            parsed.Pages.Sum(p => p.Chars),
            parsed.Pages.Select(p => new PdfPageOutline(p.Number, p.Chars, p.Images, p.Tables, p.NeedsOcr, p.Reason)).ToList(),
            parsed.Bookmarks,
            OcrAvailable,
            notes);
    }

    /// <summary>讀第 <paramref name="from"/> 到 <paramref name="to"/> 頁（含；省略 to 只讀一頁）。offset 是範圍內的字元位置，截斷時用 NextOffset 續讀。</summary>
    public ReadResult ReadPages(string path, int from, int? to = null, int offset = 0, int? maxChars = null, PdfReadOptions? pdfOptions = null)
    {
        var (_, parsed) = Load(path);
        var last = to ?? from;
        if (from < 1 || from > parsed.Pages.Count || last < from || last > parsed.Pages.Count)
        {
            throw new OfficeToolException(
                ErrorCodes.InvalidValue,
                $"頁碼範圍 {from}–{last} 無效（文件共 {parsed.Pages.Count} 頁）",
                parsed.Pages.Count == 0 ? "文件沒有頁面" : $"頁碼從 1 到 {parsed.Pages.Count}，且結束要大於等於開始");
        }

        return Compose(path, parsed, from, last, offset, maxChars, pdfOptions ?? new PdfReadOptions());
    }

    /// <summary>依字元數分頁讀整份文件。</summary>
    public ReadResult Read(string path, int offset = 0, int? maxChars = null, PdfReadOptions? pdfOptions = null)
    {
        var (_, parsed) = Load(path);
        return parsed.Pages.Count == 0
            ? TextPaging.Slice(string.Empty, 0, 0, offset, maxChars, _readerOptions.MaxCharsPerRead, null, parsed.Notes)
            : Compose(path, parsed, 1, parsed.Pages.Count, offset, maxChars, pdfOptions ?? new PdfReadOptions());
    }

    /// <summary>把頁面轉成 PNG（給多模態模型看）。一次最多 10 頁，圖片最長邊最多 4000 像素。</summary>
    public IReadOnlyList<RenderedPage> RenderPages(string path, IReadOnlyList<int> pages, int? dpi = null)
    {
        ArgumentNullException.ThrowIfNull(pages);
        var (full, parsed) = Load(path);
        if (pages.Count == 0 || pages.Count > MaxRenderPages)
        {
            throw new OfficeToolException(ErrorCodes.InvalidValue, $"一次要轉的頁數必須在 1 到 {MaxRenderPages} 之間（收到 {pages.Count}）", "分批轉換");
        }

        foreach (var page in pages)
        {
            if (page < 1 || page > parsed.Pages.Count)
            {
                throw new OfficeToolException(ErrorCodes.InvalidValue, $"頁碼 {page} 超出範圍（文件共 {parsed.Pages.Count} 頁）", "頁碼從 1 開始");
            }
        }

        var bytes = PackageChecks.ReadShared(full);
        var requested = dpi ?? _readerOptions.OcrDpi;
        if (requested is < 36 or > 600)
        {
            throw new OfficeToolException(ErrorCodes.InvalidValue, $"dpi {requested} 無效", "dpi 必須在 36 到 600 之間");
        }

        return pages.Select(p => Render(bytes, p, requested)).ToList();
    }

    // ---- 組合輸出 ----

    private ReadResult Compose(string path, Parsed parsed, int from, int last, int offset, int? maxChars, PdfReadOptions pdfOptions)
    {
        // 先確認 offset / maxChars 合法再做昂貴的 OCR：用空字串的長度檢查太粗，所以先組出文字
        var sb = new System.Text.StringBuilder();
        var notes = parsed.Notes.ToList();
        for (var n = from; n <= last; n++)
        {
            var page = parsed.Pages[n - 1];
            if (sb.Length > 0)
            {
                sb.Append("\n\n");
            }

            sb.Append("## 第 ").Append(n.ToString(CultureInfo.InvariantCulture)).Append(" 頁\n\n").Append(PageBody(path, parsed, page, pdfOptions, notes));
        }

        return TextPaging.Slice(sb.ToString().TrimEnd('\n'), 0, sb.ToString().TrimEnd('\n').Length, offset, maxChars, _readerOptions.MaxCharsPerRead, null, notes.Distinct().ToList());
    }

    private string PageBody(string path, Parsed parsed, PageResult page, PdfReadOptions pdfOptions, List<string> notes)
    {
        if (!page.NeedsOcr)
        {
            return page.Markdown.Length == 0 ? "（此頁沒有文字）" : page.Markdown;
        }

        if (pdfOptions.UseOcr && _ocr is not null && _ocr.IsAvailable(out _))
        {
            if (!parsed.Ocr.TryGetValue(page.Number, out var text))
            {
                text = RunOcr(path, page.Number, notes);
                if (text is not null)
                {
                    parsed.Ocr[page.Number] = text;
                }
            }

            if (text is not null)
            {
                notes.Add("標示「OCR」的頁面是用文字辨識的結果，可能有錯字");
                return "> （此頁以 OCR 辨識）\n\n" + (text.Length == 0 ? "（OCR 沒有辨識出文字）" : text);
            }
        }

        var reason = page.Reason ?? "沒有可用的文字層";
        notes.Add("有頁面沒有可用的文字層（掃描檔或亂碼），需要 OCR 或 RenderPages 轉圖片");
        return $"> （此頁{reason}，需要 OCR 或轉成圖片檢視）" + (page.Markdown.Length > 0 ? "\n\n" + page.Markdown : string.Empty);
    }

    /// <summary>OCR 失敗不算讀取失敗：回傳 null，頁面照「沒有 OCR」的方式標記，原因寫進 notes。</summary>
    private string? RunOcr(string path, int pageNumber, List<string> notes)
    {
        try
        {
            var full = _guard.ResolveExistingFile(path);
            var png = Render(PackageChecks.ReadShared(full), pageNumber, _readerOptions.OcrDpi).Png;
            var raw = _ocr!.Recognize(png);
            return string.Join("\n\n", raw.Replace("\r\n", "\n", StringComparison.Ordinal).Split("\n\n").Select(p => MarkdownWriter.EscapeLineStarts(MarkdownWriter.Escape(p.Trim()))).Where(p => p.Length > 0));
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            notes.Add($"第 {pageNumber.ToString(CultureInfo.InvariantCulture)} 頁 OCR 失敗：{ex.Message}");
            return null;
        }
    }

    // PDFtoImage 標了平台支援屬性（Windows / macOS / Linux 都支援，只是分析器看不出目標平台），這裡確認過就不警告
#pragma warning disable CA1416
    private static RenderedPage Render(byte[] pdf, int pageNumber, int dpi)
    {
        try
        {
            var size = Conversion.GetPageSize(new MemoryStream(pdf, writable: false), pageNumber - 1);
            var scale = dpi / 72.0;
            var longest = Math.Max(size.Width, size.Height) * scale;
            var effective = longest > MaxRenderPixels ? (int)(dpi * MaxRenderPixels / longest) : dpi;

            using var output = new MemoryStream();
            Conversion.SavePng(output, new MemoryStream(pdf, writable: false), pageNumber - 1, options: new RenderOptions(Dpi: Math.Max(effective, 36)));
            var png = output.ToArray();
            var pixelScale = Math.Max(effective, 36) / 72.0;
            return new RenderedPage(pageNumber, (int)Math.Round(size.Width * pixelScale), (int)Math.Round(size.Height * pixelScale), png);
        }
        catch (Exception ex) when (ex is not OfficeToolException and not OutOfMemoryException)
        {
            throw new OfficeToolException(ErrorCodes.CorruptFile, $"第 {pageNumber} 頁無法轉成圖片：{ex.Message}", "檔案可能損壞，或 PDF 轉圖片元件在這個平台不可用", ex);
        }
    }

#pragma warning restore CA1416

    // ---- 載入 ----

    private (string FullPath, Parsed Parsed) Load(string path)
    {
        var full = _guard.ResolveExistingFile(path);
        if (!Path.GetExtension(full).Equals(Extension, StringComparison.OrdinalIgnoreCase))
        {
            throw new OfficeToolException(
                ErrorCodes.UnsupportedFormat,
                $"PDF 讀取工具只能讀 {Extension}：{Path.GetFileName(full)}",
                "Word / PPT / Excel 請用對應的工具");
        }

        var info = new FileInfo(full);
        var key = new ParsedDocumentCache.Key(full, info.LastWriteTimeUtc.Ticks, info.Length, "pdf");
        return (full, _cache.GetOrAdd(key, () => Parse(full)));
    }

    private static Parsed Parse(string fullPath)
    {
        var bytes = PackageChecks.ReadShared(fullPath);
        if (bytes.AsSpan(0, Math.Min(bytes.Length, 1024)).IndexOf("%PDF-"u8) < 0)
        {
            throw PackageChecks.Corrupt("這不是有效的 PDF 檔案", null);
        }

        try
        {
            using var document = PdfDocument.Open(bytes);
            var contents = document.GetPages().Select(PdfPageExtractor.Extract).ToList();
            return Build(document, contents);
        }
        catch (PdfDocumentEncryptedException ex)
        {
            throw new OfficeToolException(ErrorCodes.PasswordProtected, "PDF 有密碼保護，無法讀取", "請先取得沒有密碼的版本", ex);
        }
        catch (Exception ex) when (ex is PdfDocumentFormatException or InvalidOperationException or ArgumentException or IOException or NotSupportedException or IndexOutOfRangeException or InvalidCastException)
        {
            throw PackageChecks.Corrupt("PDF 結構損壞，無法讀取", ex);
        }
    }

    private static Parsed Build(PdfDocument document, List<PageContent> contents)
    {
        var notes = new List<string>();
        var body = BodyFontSize(contents);
        var removed = RepeatedMarginWords(contents);
        if (removed.Count > 0)
        {
            notes.Add("已略過每頁重複出現的頁首 / 頁尾（頁碼、文件名稱等）");
        }

        var pages = new List<PageResult>();
        foreach (var content in contents)
        {
            var words = content.Words.Where(w => !removed.Contains(w)).ToList();
            var needsOcr = false;
            string? reason = null;
            if (content.Letters < MinTextChars && content.ImageCoverage >= 0.5)
            {
                needsOcr = true;
                reason = "是掃描影像，沒有文字層";
            }
            else if (content.Letters >= MinTextChars && content.GarbageRatio > MaxGarbageRatio)
            {
                needsOcr = true;
                reason = "文字層是亂碼（字型缺少 Unicode 對應）";
            }

            var tables = needsOcr ? [] : PdfTableFinder.Find(content with { Words = words });
            var markdown = needsOcr ? string.Empty : PdfLayout.RenderPage(content, words, body);
            if (content.Rotated)
            {
                notes.Add("有頁面被旋轉，文字順序可能不正確");
            }

            pages.Add(new PageResult(content.Number, markdown, content.Letters, content.Images, tables.Count, needsOcr, reason));
        }

        return new Parsed { Pages = pages, Bookmarks = Bookmarks(document), Notes = notes.Distinct().ToList() };
    }

    private static List<PdfBookmark> Bookmarks(PdfDocument document)
    {
        var result = new List<PdfBookmark>();
        try
        {
            if (document.TryGetBookmarks(out var bookmarks))
            {
                void Walk(IEnumerable<UglyToad.PdfPig.Outline.BookmarkNode> nodes)
                {
                    foreach (var node in nodes)
                    {
                        result.Add(new PdfBookmark(node.Title, node.Level, (node as UglyToad.PdfPig.Outline.DocumentBookmarkNode)?.PageNumber));
                        Walk(node.Children);
                    }
                }

                Walk(bookmarks.Roots);
            }
        }
        catch (Exception ex) when (ex is InvalidOperationException or ArgumentException or NotSupportedException or IndexOutOfRangeException)
        {
            // 書籤壞掉不影響內文
        }

        return result;
    }

    /// <summary>全文最常見的字級（以字元數加權）當作內文字級，標題判斷以它為基準。</summary>
    private static double BodyFontSize(List<PageContent> pages)
    {
        var groups = pages.SelectMany(p => p.Words).GroupBy(w => Math.Round(w.FontSize * 2) / 2).OrderByDescending(g => g.Sum(w => w.Text.Length)).FirstOrDefault();
        return groups?.Key ?? 10;
    }

    /// <summary>
    /// 每頁重複出現的頁首 / 頁尾：頁面最上或最下 8% 的行，數字換成 # 後在至少 3 頁、且超過一半的頁面出現。
    /// </summary>
    private static HashSet<PdfWord> RepeatedMarginWords(List<PageContent> pages)
    {
        var removed = new HashSet<PdfWord>();
        if (pages.Count < 3)
        {
            return removed;
        }

        var candidates = new List<(PageContent Page, string Key, List<PdfWord> Words)>();
        foreach (var page in pages)
        {
            var band = page.Height * 0.08;
            var margin = page.Words.Where(w => w.Box.Bottom >= page.Height - band || w.Box.Top <= band).ToList();
            foreach (var line in PdfLines.Cluster(margin))
            {
                var top = line.Box.Bottom >= page.Height / 2;
                var key = (top ? "T:" : "B:") + Digits().Replace(string.Join(" ", line.Words.Select(w => w.Text)), "#");
                candidates.Add((page, key, line.Words.ToList()));
            }
        }

        var needed = Math.Max(3, (pages.Count / 2) + 1);
        foreach (var group in candidates.GroupBy(c => c.Key))
        {
            if (group.Select(g => g.Page.Number).Distinct().Count() >= needed)
            {
                foreach (var word in group.SelectMany(g => g.Words))
                {
                    removed.Add(word);
                }
            }
        }

        return removed;
    }

    [GeneratedRegex(@"\d+", RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 1000)]
    private static partial Regex Digits();
}
