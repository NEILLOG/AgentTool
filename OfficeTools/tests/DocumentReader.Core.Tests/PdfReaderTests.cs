using DocumentReader.Core.Models;
using DocumentReader.Core.Pdf;
using OfficeTools.Common.Errors;

namespace DocumentReader.Core.Tests;

public class PdfTextNormalizerTests
{
    [Theory]
    [InlineData("⾏⽇", "行日")] // 康熙部首
    [InlineData("⼦⽴⽽", "子立而")]
    [InlineData("ofﬁce ﬂow", "office flow")] // 連字
    [InlineData("a b", "a b")]
    [InlineData("⻄⺠", "西民")] // 部首補充（手動對應）
    [InlineData("豈", "豈")] // 相容漢字
    public void Compatibility_characters_are_mapped_back(string input, string expected) =>
        Assert.Equal(expected, PdfTextNormalizer.Normalize(input));

    [Theory]
    [InlineData("Ｍerriweather １２３ ①")]
    [InlineData("普通的中文與 English 123")]
    [InlineData("")]
    public void Other_text_is_left_alone(string input) =>
        Assert.Equal(input, PdfTextNormalizer.Normalize(input));

    [Fact]
    public void Cjk_and_garbage_detection()
    {
        Assert.True(PdfTextNormalizer.IsCjk('中'));
        Assert.True(PdfTextNormalizer.IsCjk('あ'));
        Assert.True(PdfTextNormalizer.IsCjk('，'));
        Assert.False(PdfTextNormalizer.IsCjk('a'));
        Assert.True(PdfTextNormalizer.IsGarbage(''));
        Assert.True(PdfTextNormalizer.IsGarbage('�'));
        Assert.True(PdfTextNormalizer.IsGarbage('\u0001'));
        Assert.False(PdfTextNormalizer.IsGarbage('中'));
        Assert.False(PdfTextNormalizer.IsGarbage('\n'));
        Assert.False(PdfTextNormalizer.IsGarbage('\0'));
    }
}

public sealed class PdfOutlineAndPagingTests : IDisposable
{
    private readonly PdfEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    private string ThreePages() => _env.Save(new PdfBuilder()
        .Page().Text("alpha page", 50, 700)
        .Page().Text("beta page", 50, 700).Image(50, 300, 100, 100)
        .Page().Text("gamma page", 50, 700));

    [Fact]
    public void The_outline_lists_pages_with_facts()
    {
        var outline = _env.Reader.GetOutline(ThreePages());

        Assert.Equal(3, outline.PageCount);
        Assert.Equal([1, 2, 3], outline.Pages.Select(p => p.Number));
        Assert.Equal([10, 9, 10], outline.Pages.Select(p => p.Chars));
        Assert.Equal([0, 1, 0], outline.Pages.Select(p => p.Images));
        Assert.All(outline.Pages, p => Assert.False(p.NeedsOcr));
        Assert.Equal(29, outline.TotalChars);
        Assert.False(outline.OcrAvailable);
        Assert.EndsWith("doc.pdf", outline.Path, StringComparison.Ordinal);
    }

    [Fact]
    public void Reading_a_page_returns_only_that_page()
    {
        var result = _env.Reader.ReadPages(ThreePages(), 2);

        Assert.Equal("## 第 2 頁\n\nbeta page", result.Markdown);
        Assert.False(result.Truncated);
    }

    [Fact]
    public void Reading_a_range_includes_both_ends()
    {
        var result = _env.Reader.ReadPages(ThreePages(), 2, 3);

        Assert.Equal("## 第 2 頁\n\nbeta page\n\n## 第 3 頁\n\ngamma page", result.Markdown);
    }

    [Fact]
    public void Reading_page_by_page_equals_reading_everything()
    {
        var path = ThreePages();

        Assert.Equal(
            _env.Reader.Read(path).Markdown,
            string.Join("\n\n", Enumerable.Range(1, 3).Select(i => _env.Reader.ReadPages(path, i).Markdown)));
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(4, null)]
    [InlineData(2, 1)]
    [InlineData(1, 4)]
    public void A_bad_page_range_is_rejected(int from, int? to)
    {
        var ex = Throws(() => _env.Reader.ReadPages(ThreePages(), from, to));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("3", ex.Message + ex.Hint, StringComparison.Ordinal);
    }

    [Fact]
    public void Long_content_is_paged_with_next_offset()
    {
        using var env = new PdfEnv(maxChars: 120);
        var b = new PdfBuilder().Page();
        for (var i = 0; i < 30; i++)
        {
            b.Text($"Paragraph number {i} with some text", 50, 800 - (i * 25));
        }

        var path = env.Save(b);
        var all = string.Empty;
        var offset = 0;
        var pages = 0;
        while (true)
        {
            var page = env.Reader.ReadPages(path, 1, offset: offset);
            Assert.True(page.Markdown.Length <= 120);
            all += page.Markdown + "\n";
            pages++;
            if (!page.Truncated)
            {
                break;
            }

            Assert.True(page.NextOffset > offset);
            offset = page.NextOffset!.Value;
        }

        Assert.True(pages > 3);
        for (var i = 0; i < 30; i++)
        {
            Assert.Contains($"Paragraph number {i} with", all, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void A_bad_offset_or_max_is_rejected()
    {
        var path = ThreePages();

        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Reader.ReadPages(path, 1, offset: 100000)).Code);
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Reader.Read(path, maxChars: 0)).Code);
    }

    [Fact]
    public void A_document_without_pages_reads_as_empty()
    {
        var path = _env.Path("none.pdf");
        File.WriteAllBytes(path, new PdfBuilder().Build());

        Assert.Equal(0, _env.Reader.GetOutline(path).PageCount);
        Assert.Equal(string.Empty, _env.Reader.Read(path).Markdown);
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Reader.ReadPages(path, 1)).Code);
    }
}

public sealed class PdfOcrAndRenderTests : IDisposable
{
    private readonly List<IDisposable> _envs = [];

    public void Dispose()
    {
        foreach (var env in _envs)
        {
            env.Dispose();
        }
    }

    private PdfEnv Env(IOcrEngine? ocr)
    {
        var env = new PdfEnv(ocr);
        _envs.Add(env);
        return env;
    }

    /// <summary>第 1 頁文字、第 2 頁整頁是圖片沒有文字（掃描檔）。</summary>
    private static PdfBuilder Scanned() => new PdfBuilder().Page().Text("normal text page", 50, 700).Page().Image(0, 0, 595, 842);

    [Fact]
    public void An_image_only_page_is_flagged_as_needing_ocr()
    {
        var env = Env(null);
        var outline = env.Reader.GetOutline(env.Save(Scanned()));

        Assert.False(outline.Pages[0].NeedsOcr);
        Assert.True(outline.Pages[1].NeedsOcr);
        Assert.Contains("掃描", outline.Pages[1].Reason, StringComparison.Ordinal);
        Assert.Contains(outline.Notes, n => n.Contains("OCR", StringComparison.Ordinal));
    }

    [Fact]
    public void Without_an_engine_the_scanned_page_is_marked_not_dropped()
    {
        var env = Env(null);
        var result = env.Reader.Read(env.Save(Scanned()));

        Assert.Contains("normal text page", result.Markdown, StringComparison.Ordinal);
        Assert.Contains("## 第 2 頁\n\n> （此頁是掃描影像，沒有文字層，需要 OCR 或轉成圖片檢視）", result.Markdown, StringComparison.Ordinal);
        Assert.Contains(result.Notes, n => n.Contains("OCR", StringComparison.Ordinal));
    }

    [Fact]
    public void An_available_engine_is_used_for_flagged_pages_only()
    {
        var ocr = new FakeOcr("recognized words\n\nsecond block");
        var env = Env(ocr);
        var result = env.Reader.Read(env.Save(Scanned()));

        Assert.Equal(1, ocr.Calls);
        Assert.Contains("> （此頁以 OCR 辨識）\n\nrecognized words\n\nsecond block", result.Markdown, StringComparison.Ordinal);
        Assert.Contains("normal text page", result.Markdown, StringComparison.Ordinal);
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, ocr.LastImage![..4]); // PNG
        Assert.True(env.Reader.GetOutline(env.Path("doc.pdf")).OcrAvailable);
    }

    [Fact]
    public void Ocr_results_are_cached_per_page()
    {
        var ocr = new FakeOcr();
        var env = Env(ocr);
        var path = env.Save(Scanned());

        _ = env.Reader.ReadPages(path, 2);
        _ = env.Reader.ReadPages(path, 1, 2);
        _ = env.Reader.Read(path);

        Assert.Equal(1, ocr.Calls);
    }

    [Fact]
    public void Ocr_can_be_turned_off_per_call()
    {
        var ocr = new FakeOcr();
        var env = Env(ocr);
        var result = env.Reader.ReadPages(env.Save(Scanned()), 2, pdfOptions: new PdfReadOptions(UseOcr: false));

        Assert.Equal(0, ocr.Calls);
        Assert.Contains("需要 OCR", result.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void An_unavailable_engine_behaves_like_no_engine()
    {
        var ocr = new FakeOcr(available: false);
        var env = Env(ocr);
        var path = env.Save(Scanned());
        var result = env.Reader.Read(path);

        Assert.Equal(0, ocr.Calls);
        Assert.Contains("需要 OCR", result.Markdown, StringComparison.Ordinal);
        Assert.False(env.Reader.GetOutline(path).OcrAvailable);
    }

    [Fact]
    public void A_crashing_engine_does_not_fail_the_read()
    {
        var env = Env(new FakeOcr(throws: true));
        var result = env.Reader.Read(env.Save(Scanned()));

        Assert.Contains("normal text page", result.Markdown, StringComparison.Ordinal);
        Assert.Contains(result.Notes, n => n.Contains("OCR 失敗", StringComparison.Ordinal) && n.Contains("engine crashed", StringComparison.Ordinal));
        Assert.Contains("需要 OCR", result.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void Ocr_text_is_escaped_for_markdown()
    {
        var env = Env(new FakeOcr("# fake heading\n\n*star*"));
        var result = env.Reader.ReadPages(env.Save(Scanned()), 2);

        Assert.Contains("\\# fake heading", result.Markdown, StringComparison.Ordinal);
        Assert.Contains("\\*star\\*", result.Markdown, StringComparison.Ordinal);
    }

    // ---- 轉圖片 ----

    [Fact]
    public void Pages_can_be_rendered_to_png()
    {
        var env = Env(null);
        var images = env.Reader.RenderPages(env.Save(new PdfBuilder().Page(300, 400).Text("hi", 20, 300).Page(300, 400)), [1, 2], dpi: 72);

        Assert.Equal([1, 2], images.Select(i => i.Page));
        Assert.All(images, i =>
        {
            Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, i.Png[..4]);
            Assert.Equal(300, i.Width);
            Assert.Equal(400, i.Height);
        });
    }

    [Fact]
    public void Higher_dpi_gives_a_larger_image_but_never_beyond_the_pixel_cap()
    {
        var env = Env(null);
        var path = env.Save(new PdfBuilder().Page(600, 800));

        Assert.Equal(1200, env.Reader.RenderPages(path, [1], dpi: 144)[0].Width);
        var capped = env.Reader.RenderPages(path, [1], dpi: 600)[0];
        Assert.True(Math.Max(capped.Width, capped.Height) <= 4000);
    }

    [Fact]
    public void Render_arguments_are_validated()
    {
        var env = Env(null);
        var path = env.Save(new PdfBuilder().Page());

        Assert.Equal(ErrorCodes.InvalidValue, Assert.Throws<OfficeToolException>(() => env.Reader.RenderPages(path, [])).Code);
        Assert.Equal(ErrorCodes.InvalidValue, Assert.Throws<OfficeToolException>(() => env.Reader.RenderPages(path, [2])).Code);
        Assert.Equal(ErrorCodes.InvalidValue, Assert.Throws<OfficeToolException>(() => env.Reader.RenderPages(path, [0])).Code);
        Assert.Equal(ErrorCodes.InvalidValue, Assert.Throws<OfficeToolException>(() => env.Reader.RenderPages(path, [1], dpi: 10)).Code);
        Assert.Equal(ErrorCodes.InvalidValue, Assert.Throws<OfficeToolException>(() => env.Reader.RenderPages(path, Enumerable.Repeat(1, 11).ToList())).Code);
    }
}

public sealed class PdfErrorsAndCacheTests : IDisposable
{
    private readonly PdfEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    private string Good(string name = "doc.pdf") => _env.Save(new PdfBuilder().Page().Text("hello", 50, 700), name);

    [Fact]
    public void A_missing_file_returns_FILE_NOT_FOUND() =>
        Assert.Equal(ErrorCodes.FileNotFound, Throws(() => _env.Reader.Read(_env.Path("missing.pdf"))).Code);

    [Fact]
    public void A_path_outside_the_roots_is_rejected()
    {
        var outside = Path.Combine(Path.GetTempPath(), "OfficeToolsTests", Guid.NewGuid().ToString("N"), "secret.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(outside)!);
        try
        {
            new PdfBuilder().Page().Save(outside);
            Assert.Equal(ErrorCodes.PathNotAllowed, Throws(() => _env.Reader.GetOutline(outside)).Code);
            Assert.Equal(ErrorCodes.PathNotAllowed, Throws(() => _env.Reader.RenderPages(outside, [1])).Code);
        }
        finally
        {
            Directory.Delete(Path.GetDirectoryName(outside)!, recursive: true);
        }
    }

    [Theory]
    [InlineData("book.xlsx")]
    [InlineData("doc.docx")]
    [InlineData("slides.pptx")]
    [InlineData("note.txt")]
    public void Other_formats_are_unsupported(string name)
    {
        File.WriteAllText(_env.Path(name), "x");

        Assert.Equal(ErrorCodes.UnsupportedFormat, Throws(() => _env.Reader.Read(_env.Path(name))).Code);
    }

    [Theory]
    [InlineData("這不是 PDF")]
    [InlineData("")]
    public void Content_that_is_not_a_pdf_is_corrupt(string content)
    {
        File.WriteAllText(_env.Path("bad.pdf"), content);

        Assert.Equal(ErrorCodes.CorruptFile, Throws(() => _env.Reader.Read(_env.Path("bad.pdf"))).Code);
    }

    [Fact]
    public void A_truncated_pdf_is_corrupt()
    {
        var bytes = new PdfBuilder().Page().Text("hello", 50, 700).Build();
        File.WriteAllBytes(_env.Path("cut.pdf"), bytes[..(bytes.Length / 3)]);

        Assert.Equal(ErrorCodes.CorruptFile, Throws(() => _env.Reader.Read(_env.Path("cut.pdf"))).Code);
    }

    [Fact]
    public void Parsing_is_cached_until_the_file_changes()
    {
        var path = Good();
        _ = _env.Reader.Read(path);
        _ = _env.Reader.GetOutline(path);
        _ = _env.Reader.ReadPages(path, 1);
        Assert.Equal(1, _env.Cache.Count);

        new PdfBuilder().Page().Text("changed and longer content", 50, 700).Save(path);
        Assert.Contains("changed and longer", _env.Reader.Read(path).Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void A_failed_parse_is_not_cached()
    {
        File.WriteAllText(_env.Path("x.pdf"), "bad");
        Throws(() => _env.Reader.Read(_env.Path("x.pdf")));

        Assert.Equal(0, _env.Cache.Count);
    }

    [Fact]
    public void The_file_is_not_locked_after_reading()
    {
        var path = Good();
        _ = _env.Reader.Read(path);
        _ = _env.Reader.RenderPages(path, [1], 72);

        File.Delete(path);
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void Concurrent_reads_agree()
    {
        var path = Good();
        var results = Enumerable.Range(0, 12).AsParallel().Select(_ => _env.Reader.Read(path).Markdown).ToList();

        Assert.All(results, r => Assert.Equal(results[0], r));
        Assert.Equal(1, _env.Cache.Count);
    }
}

/// <summary>用使用者提供的真實 PDF（repo 外的 sample/，不進版控）驗證；檔案不存在時略過。</summary>
public sealed class RealPdfSampleTests
{
    private static string? Find(string name)
    {
        for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir is not null; dir = dir.Parent)
        {
            var candidate = Path.Combine(dir.FullName, "sample", name);
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }

    private static string Copy(PdfEnv env, string name)
    {
        var target = env.Path(name.Replace("-", string.Empty, StringComparison.Ordinal));
        File.Copy(Find(name)!, target);
        return target;
    }

    [Fact]
    public void The_itinerary_pdf_has_tables_links_and_normalized_characters()
    {
        if (Find("sample-一般.pdf") is null)
        {
            return;
        }

        using var env = new PdfEnv();
        var path = Copy(env, "sample-一般.pdf");

        var outline = env.Reader.GetOutline(path);
        Assert.Equal(3, outline.PageCount);
        Assert.True(outline.Pages[0].Tables >= 1);

        var md = env.Reader.Read(path).Markdown;
        Assert.Contains("| 日期 | 時間 | 結束時間 | 位置 | 附註 |", md, StringComparison.Ordinal); // 康熙部首 ⽇ 已換成 日
        Assert.Contains("| 2026/5/13 | 上午10:00 |", md, StringComparison.Ordinal);
        Assert.Contains("[煙火資訊](https://okinawamarket.jp/", md, StringComparison.Ordinal);
        Assert.Contains("| 航空公司 | 航班號碼 |", md, StringComparison.Ordinal);
        Assert.DoesNotContain("⽇", md, StringComparison.Ordinal);
    }

    [Fact]
    public void The_quotation_pdf_keeps_its_table_and_the_text_around_it()
    {
        if (Find("sample-表格.pdf") is null)
        {
            return;
        }

        using var env = new PdfEnv();
        var md = env.Reader.Read(Copy(env, "sample-表格.pdf")).Markdown;

        Assert.Contains("虛擬主機報價單", md, StringComparison.Ordinal);
        Assert.Contains("| 設備 | 設備 | 數量 | 單價(年租金) | 單位 | 總價 |", md, StringComparison.Ordinal);
        Assert.Contains("14,900", md, StringComparison.Ordinal);
        Assert.Contains("聯絡電話：02-33665022/33665023", md, StringComparison.Ordinal);
    }

    [Fact]
    public void The_complex_layout_pdf_is_read_without_losing_content_and_can_be_rendered()
    {
        if (Find("sample-複雜排版.pdf") is null)
        {
            return;
        }

        using var env = new PdfEnv();
        var path = Copy(env, "sample-複雜排版.pdf");
        var md = env.Reader.Read(path).Markdown;

        Assert.Contains("中央研究院 (Academia Sinica)", md, StringComparison.Ordinal);
        Assert.Contains("- 世界銀行 (World Bank)", md, StringComparison.Ordinal);
        Assert.Contains("Relevant Institutions and Websites", md, StringComparison.Ordinal);
        Assert.DoesNotContain("⽽", md, StringComparison.Ordinal);
        Assert.Single(env.Reader.RenderPages(path, [1], 100));
    }
}
