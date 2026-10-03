using DocumentReader.Core.Caching;
using DocumentReader.Core.Models;
using DocumentReader.Core.Ppt;
using OfficeTools.Common;
using OfficeTools.Common.Security;

namespace DocumentReader.Core.Tests;

internal sealed class PptEnv : IDisposable
{
    private readonly TempDirectory _tmp = new();

    public PptEnv(int maxChars = 1_000_000, int cacheEntries = 20)
    {
        Root = _tmp.Combine("root");
        Directory.CreateDirectory(Root);
        Options = new OfficeToolsOptions { AllowedRoots = [Root] };
        ReaderOptions = new DocumentReaderOptions { MaxCharsPerRead = maxChars, CacheMaxEntries = cacheEntries };
        Cache = new ParsedDocumentCache(cacheEntries);
        Reader = new PptReader(new PathGuard(Options), Options, ReaderOptions, Cache);
    }

    public string Root { get; }

    public OfficeToolsOptions Options { get; }

    public DocumentReaderOptions ReaderOptions { get; }

    public ParsedDocumentCache Cache { get; }

    public PptReader Reader { get; }

    public string Path(string name) => System.IO.Path.Combine(Root, name);

    public string Save(PptxBuilder builder, string name = "deck.pptx") => builder.Save(Path(name));

    /// <summary>把簡報完整轉成 Markdown（不分頁）。</summary>
    public string Md(PptxBuilder builder, PptReadOptions? options = null, string name = "deck.pptx") => Reader.Read(Save(builder, name), pptOptions: options).Markdown;

    public void Dispose() => _tmp.Dispose();
}
