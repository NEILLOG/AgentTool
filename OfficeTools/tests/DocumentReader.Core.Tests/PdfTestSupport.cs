using DocumentReader.Core.Caching;
using DocumentReader.Core.Pdf;
using OfficeTools.Common;
using OfficeTools.Common.Security;
using SkiaSharp;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Fonts.Standard14Fonts;
using UglyToad.PdfPig.Writer;

namespace DocumentReader.Core.Tests;

#pragma warning disable CA1001 // 測試用的短命物件：PdfDocumentBuilder 只持有記憶體，不需要釋放
/// <summary>用 PdfPig 的寫入器組出測試用 PDF（標準 Helvetica 字型，只能放拉丁字元；位置、框線、圖片、連結都可控制）。</summary>
internal sealed class PdfBuilder
{
    private readonly PdfDocumentBuilder _doc = new();
    private readonly PdfDocumentBuilder.AddedFont _font;
    private readonly PdfDocumentBuilder.AddedFont _bold;
    private PdfPageBuilder? _page;

    public PdfBuilder()
    {
        _font = _doc.AddStandard14Font(Standard14Font.Helvetica);
        _bold = _doc.AddStandard14Font(Standard14Font.HelveticaBold);
    }

    public PdfBuilder Page(double width = 595, double height = 842)
    {
        _page = _doc.AddPage(width, height);
        return this;
    }

    public PdfBuilder Text(string text, double x, double y, double size = 10, bool bold = false)
    {
        _page!.AddText(text, size, new PdfPoint(x, y), bold ? _bold : _font);
        return this;
    }

    /// <summary>從上到下的一組行，行距 = 字級 × 1.3。</summary>
    public PdfBuilder Lines(double x, double top, double size, params string[] lines)
    {
        for (var i = 0; i < lines.Length; i++)
        {
            Text(lines[i], x, top - (i * size * 1.3), size);
        }

        return this;
    }

    public PdfBuilder Line(double x1, double y1, double x2, double y2)
    {
        _page!.DrawLine(new PdfPoint(x1, y1), new PdfPoint(x2, y2), 0.5);
        return this;
    }

    /// <summary>用框線畫出格子：<paramref name="columnX"/> 是各欄邊界，<paramref name="rowY"/> 是各列邊界（由上而下）。</summary>
    public PdfBuilder Grid(double[] columnX, double[] rowY)
    {
        foreach (var y in rowY)
        {
            Line(columnX[0], y, columnX[^1], y);
        }

        foreach (var x in columnX)
        {
            Line(x, rowY[0], x, rowY[^1]);
        }

        return this;
    }

    public PdfBuilder Link(string uri, double left, double bottom, double right, double top)
    {
        _page!.AddLink(uri, new PdfRectangle(left, bottom, right, top));
        return this;
    }

    public PdfBuilder Image(double left, double bottom, double width, double height)
    {
        using var bitmap = new SKBitmap(40, 40);
        bitmap.Erase(SKColors.Gray);
        using var data = SKImage.FromBitmap(bitmap).Encode(SKEncodedImageFormat.Png, 100);
        _page!.AddPng(data.ToArray(), new PdfRectangle(left, bottom, left + width, bottom + height));
        return this;
    }

    public byte[] Build() => _doc.Build();

    public string Save(string path)
    {
        File.WriteAllBytes(path, Build());
        return path;
    }
}

internal sealed class FakeOcr(string text = "scanned text", bool available = true, bool throws = false) : IOcrEngine
{
    public int Calls { get; private set; }

    public byte[]? LastImage { get; private set; }

    public bool IsAvailable(out string? reason)
    {
        reason = available ? null : "no language pack";
        return available;
    }

    public string Recognize(byte[] png)
    {
        Calls++;
        LastImage = png;
        return throws ? throw new InvalidOperationException("engine crashed") : text;
    }
}

internal sealed class PdfEnv : IDisposable
{
    private readonly TempDirectory _tmp = new();

    public PdfEnv(IOcrEngine? ocr = null, int maxChars = 1_000_000, int cacheEntries = 20)
    {
        Root = _tmp.Combine("root");
        Directory.CreateDirectory(Root);
        Options = new OfficeToolsOptions { AllowedRoots = [Root] };
        ReaderOptions = new DocumentReaderOptions { MaxCharsPerRead = maxChars, CacheMaxEntries = cacheEntries };
        Cache = new ParsedDocumentCache(cacheEntries);
        Reader = new PdfReader(new PathGuard(Options), Options, ReaderOptions, ocr, Cache);
    }

    public string Root { get; }

    public OfficeToolsOptions Options { get; }

    public DocumentReaderOptions ReaderOptions { get; }

    public ParsedDocumentCache Cache { get; }

    public PdfReader Reader { get; }

    public string Path(string name) => System.IO.Path.Combine(Root, name);

    public string Save(PdfBuilder builder, string name = "doc.pdf") => builder.Save(Path(name));

    /// <summary>整份轉成 Markdown（不分頁）。</summary>
    public string Md(PdfBuilder builder, string name = "doc.pdf") => Reader.Read(Save(builder, name)).Markdown;

    public void Dispose() => _tmp.Dispose();
}
