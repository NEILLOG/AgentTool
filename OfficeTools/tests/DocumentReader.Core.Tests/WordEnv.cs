using DocumentReader.Core.Caching;
using DocumentReader.Core.Models;
using DocumentReader.Core.Word;
using OfficeTools.Common;
using OfficeTools.Common.Security;

namespace DocumentReader.Core.Tests;

internal sealed class WordEnv : IDisposable
{
    private readonly TempDirectory _tmp = new();

    public WordEnv(int maxChars = 1_000_000, int cacheEntries = 20, int maxUncompressedMb = 500, long maxCharactersInPart = 50_000_000)
    {
        Root = _tmp.Combine("root");
        Directory.CreateDirectory(Root);
        Options = new OfficeToolsOptions { AllowedRoots = [Root], MaxUncompressedMb = maxUncompressedMb, MaxCharactersInPart = maxCharactersInPart };
        ReaderOptions = new DocumentReaderOptions { MaxCharsPerRead = maxChars, CacheMaxEntries = cacheEntries };
        Cache = new ParsedDocumentCache(cacheEntries);
        Reader = new WordReader(new PathGuard(Options), Options, ReaderOptions, Cache);
    }

    public string Root { get; }

    public OfficeToolsOptions Options { get; }

    public DocumentReaderOptions ReaderOptions { get; }

    public ParsedDocumentCache Cache { get; }

    public WordReader Reader { get; }

    public string Path(string name) => System.IO.Path.Combine(Root, name);

    public string Outside(string name)
    {
        var full = _tmp.Combine("outside", name);
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(full)!);
        return full;
    }

    public string Save(DocxBuilder builder, string name = "doc.docx") => builder.Save(Path(name));

    /// <summary>把文件完整轉成 Markdown（不分頁）。</summary>
    public string Md(DocxBuilder builder, WordReadOptions? options = null, string name = "doc.docx")
    {
        var path = Save(builder, name);
        return Reader.Read(path, wordOptions: options).Markdown;
    }

    public void Dispose() => _tmp.Dispose();
}
