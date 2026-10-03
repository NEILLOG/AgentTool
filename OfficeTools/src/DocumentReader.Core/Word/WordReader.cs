using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentReader.Core.Caching;
using DocumentReader.Core.Markdown;
using DocumentReader.Core.Models;
using OfficeTools.Common;
using OfficeTools.Common.Errors;
using OfficeTools.Common.Security;

namespace DocumentReader.Core.Word;

/// <summary>
/// 唯讀地把 .docx 轉成 Markdown。docx 沒有頁碼資訊，所以用標題切段：先用 <see cref="GetOutline"/> 看標題樹，
/// 再用 <see cref="ReadSection"/> 讀某一節；沒有標題的文件改用 <see cref="Read"/> 依字數分頁。
/// 解析結果會依「路徑 + 最後修改時間 + 檔案大小」快取，同一份文件連續讀多次不會重複解析。
/// </summary>
public sealed class WordReader
{
    private const string Extension = ".docx";
    private const int MaxListedSections = 30;

    private readonly PathGuard _guard;
    private readonly OfficeToolsOptions _options;
    private readonly DocumentReaderOptions _readerOptions;
    private readonly ParsedDocumentCache _cache;

    public WordReader(PathGuard guard, OfficeToolsOptions options, DocumentReaderOptions readerOptions, ParsedDocumentCache? cache = null)
    {
        ArgumentNullException.ThrowIfNull(guard);
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(readerOptions);
        _guard = guard;
        _options = options;
        _readerOptions = readerOptions;
        _cache = cache ?? new ParsedDocumentCache(readerOptions.CacheMaxEntries);
    }

    /// <summary>標題樹：每一節的 sectionId、層級、標題與字元數。第一個標題之前的內容是前言（s0）。</summary>
    public DocumentOutline GetOutline(string path, WordReadOptions? wordOptions = null)
    {
        var (full, doc) = Load(path, wordOptions ?? new WordReadOptions());
        var byParent = doc.Sections.GroupBy(s => s.ParentId ?? string.Empty).ToDictionary(g => g.Key, g => g.ToList());

        OutlineNode Build(SectionInfo s) => new(
            s.Id,
            s.Level,
            s.Title,
            s.OwnEnd - s.Start,
            s.TotalEnd - s.Start,
            byParent.TryGetValue(s.Id, out var children) ? children.Select(Build).ToList() : []);

        var top = byParent.TryGetValue(string.Empty, out var roots) ? roots.Select(Build).ToList() : [];
        return new DocumentOutline(full, doc.Text.Length, doc.Sections.Any(s => s.Level > 0), top, doc.Notes);
    }

    /// <summary>讀某一節。<paramref name="offset"/> 是在這一節範圍內的字元位置，被截斷時用回傳的 NextOffset 續讀。</summary>
    public ReadResult ReadSection(string path, string sectionId, bool includeSubsections = true, int offset = 0, int? maxChars = null, WordReadOptions? wordOptions = null)
    {
        var (_, doc) = Load(path, wordOptions ?? new WordReadOptions());
        var section = doc.Sections.FirstOrDefault(s => string.Equals(s.Id, sectionId?.Trim(), StringComparison.OrdinalIgnoreCase));
        if (section is null)
        {
            var listed = string.Join("、", doc.Sections.Take(MaxListedSections).Select(s => $"{s.Id} {s.Title}"));
            var more = doc.Sections.Count > MaxListedSections ? $" …（共 {doc.Sections.Count} 節）" : string.Empty;
            throw new OfficeToolException(
                ErrorCodes.SectionNotFound,
                $"找不到章節「{sectionId}」",
                doc.Sections.Count == 0 ? "文件沒有內容可讀" : $"現有的章節：{listed}{more}；沒有標題的文件請用 Read 依字數分頁");
        }

        return Slice(doc, section.Start, includeSubsections ? section.TotalEnd : section.OwnEnd, offset, maxChars, section.Id);
    }

    /// <summary>
    /// 依字元數分頁讀整份文件（給沒有標題、或想從頭順著讀的文件用）。<paramref name="offset"/> 是整份文件的字元位置；
    /// 被截斷時會盡量停在段落邊界，並用 NextOffset 告訴你下一次從哪裡接著讀。
    /// </summary>
    public ReadResult Read(string path, int offset = 0, int? maxChars = null, WordReadOptions? wordOptions = null)
    {
        var (_, doc) = Load(path, wordOptions ?? new WordReadOptions());
        return Slice(doc, 0, doc.Text.Length, offset, maxChars, sectionId: null);
    }

    // ---- 載入與快取 ----

    private (string FullPath, ParsedDocument Document) Load(string path, WordReadOptions wordOptions)
    {
        var full = _guard.ResolveExistingFile(path);
        if (!Path.GetExtension(full).Equals(Extension, StringComparison.OrdinalIgnoreCase))
        {
            throw new OfficeToolException(
                ErrorCodes.UnsupportedFormat,
                $"Word 讀取工具只能讀 {Extension}：{Path.GetFileName(full)}",
                "Excel 請用 Excel 工具，PPT / PDF 請用對應的讀取工具；.doc 請先用 Word 另存為 .docx");
        }

        var info = new FileInfo(full);
        var key = new ParsedDocumentCache.Key(
            full,
            info.LastWriteTimeUtc.Ticks,
            info.Length,
            $"word:{wordOptions.IncludeHeadersFooters}:{wordOptions.IncludeFootnotes}:{wordOptions.IncludeComments}");
        return (full, _cache.GetOrAdd(key, () => Parse(full, wordOptions)));
    }

    private ParsedDocument Parse(string fullPath, WordReadOptions wordOptions)
    {
        var bytes = PackageChecks.ReadShared(fullPath);
        PackageChecks.ValidateZip(bytes, _options, "[Content_Types].xml", "word/document.xml");

        try
        {
            using var document = WordprocessingDocument.Open(
                new MemoryStream(bytes, writable: false),
                isEditable: false,
                new OpenSettings { MaxCharactersInPart = _options.MaxCharactersInPart });
            var main = document.MainDocumentPart ?? throw PackageChecks.Corrupt("文件缺少主要內容部件", null);
            return new WordConverter(main, wordOptions).Convert();
        }
        catch (XmlException ex) when (ex.Message.Contains("MaxCharacters", StringComparison.Ordinal)) // 訊息是「exceeded a limit set by MaxCharactersInDocument」，識別字不會被在地化
        {
            throw new OfficeToolException(ErrorCodes.FileTooLarge, "文件內容超過單一部件的字元上限", "請先將文件分割或縮小後再處理", ex);
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or XmlException or InvalidDataException or IOException or InvalidOperationException)
        {
            throw PackageChecks.Corrupt("文件結構損壞，無法讀取", ex);
        }
    }

    // ---- 分頁 ----

    private ReadResult Slice(ParsedDocument doc, int scopeStart, int scopeEnd, int offset, int? maxChars, string? sectionId) =>
        TextPaging.Slice(doc.Text, scopeStart, scopeEnd, offset, maxChars, _readerOptions.MaxCharsPerRead, sectionId, doc.Notes);
}
