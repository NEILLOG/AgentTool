using System.IO.Compression;
using System.Text;
using DocumentReader.Core.Models;
using OfficeTools.Common.Errors;
using static DocumentReader.Core.Tests.PptxBuilder;

namespace DocumentReader.Core.Tests;

public sealed class PptChartTests : IDisposable
{
    private readonly PptEnv _env = new();

    public void Dispose() => _env.Dispose();

    private string Md(string chartXml, string? extraShapes = null) =>
        _env.Md(new PptxBuilder().Slide(ChartFrame(2, "rId5") + extraShapes, charts: [("rId5", chartXml)]));

    [Fact]
    public void A_chart_shows_its_title_type_and_cached_data()
    {
        var md = Md(ChartXml("季營收", ["Q1", "Q2", "Q3"], [("2024", [10, 20, 30]), ("2025", [15, 25, 35.5])]));

        Assert.Equal("## 投影片 1\n\n[圖表: 季營收（直條圖）]\n\n|   | 2024 | 2025 |\n| --- | --- | --- |\n| Q1 | 10 | 15 |\n| Q2 | 20 | 25 |\n| Q3 | 30 | 35.5 |", md);
    }

    [Fact]
    public void Horizontal_bars_line_and_pie_charts_are_named()
    {
        var series = new (string, double?[])[] { ("s", [1]) };
        Assert.Contains("橫條圖", Md(ChartXml("t", ["a"], series, extraInGroup: "<c:barDir val=\"bar\"/>")), StringComparison.Ordinal);
        Assert.Contains("直條圖", Md(ChartXml("t", ["a"], series, extraInGroup: "<c:barDir val=\"col\"/>")), StringComparison.Ordinal);
        Assert.Contains("折線圖", Md(ChartXml("t", ["a"], series, kind: "lineChart")), StringComparison.Ordinal);
        Assert.Contains("圓餅圖", Md(ChartXml("t", ["a"], series, kind: "pieChart")), StringComparison.Ordinal);
    }

    [Fact]
    public void A_chart_without_a_title_still_lists_its_type_and_data()
    {
        var md = Md(ChartXml(null, ["a"], [("s", [1])], kind: "lineChart"));

        Assert.Contains("[圖表: 折線圖]", md, StringComparison.Ordinal);
        Assert.Contains("| a | 1 |", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Missing_points_leave_empty_cells()
    {
        var md = Md(ChartXml("t", ["a", "b", "c"], [("s", [1, null, 3])]));

        Assert.Contains("| b |   |", md, StringComparison.Ordinal);
        Assert.Contains("| c | 3 |", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Series_and_category_names_are_escaped_for_tables()
    {
        var md = Md(ChartXml("a|b", ["x|y"], [("s|1", [1])]));

        Assert.Contains("x\\|y", md, StringComparison.Ordinal);
        Assert.Contains("s\\|1", md, StringComparison.Ordinal);
    }

    [Fact]
    public void A_chart_without_cached_values_is_only_a_label()
    {
        var xml = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val=\"0\"/><c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>";

        Assert.Equal("## 投影片 1\n\n[圖表: 直條圖]", Md(xml));
    }

    [Fact]
    public void A_chart_whose_part_is_missing_is_a_plain_label()
    {
        var md = _env.Md(new PptxBuilder().Slide(ChartFrame(2, "rId77")));

        Assert.Equal("## 投影片 1\n\n[圖表]", md);
    }

    [Fact]
    public void Very_long_series_are_capped()
    {
        var categories = Enumerable.Range(1, 250).Select(i => $"c{i}").ToArray();
        var values = Enumerable.Range(1, 250).Select(i => (double?)i).ToArray();
        var md = Md(ChartXml("long", categories, [("s", values)]));

        Assert.Contains("| c100 | 100 |", md, StringComparison.Ordinal);
        Assert.DoesNotContain("c101", md, StringComparison.Ordinal);
        Assert.Contains("圖表共 250 筆資料，只列出前 100 筆", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Several_chart_groups_are_combined_into_one_table()
    {
        var chart = ChartXml("combo", ["a", "b"], [("柱", [1, 2])]).Replace("</c:plotArea>", ChartXml(null, ["a", "b"], [("線", [3, 4])], kind: "lineChart").Split("<c:plotArea>")[1].Split("</c:plotArea>")[0] + "</c:plotArea>", StringComparison.Ordinal);
        var md = Md(chart);

        Assert.Contains("直條圖＋折線圖", md, StringComparison.Ordinal);
        Assert.Contains("|   | 柱 | 線 |", md, StringComparison.Ordinal);
        Assert.Contains("| a | 1 | 3 |", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Scatter_charts_use_x_values_as_the_first_column()
    {
        var chart = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart><c:plotArea><c:scatterChart><c:ser><c:idx val=\"0\"/><c:tx><c:v>點</c:v></c:tx>"
            + "<c:xVal><c:numRef><c:numCache><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>1.5</c:v></c:pt><c:pt idx=\"1\"><c:v>2.5</c:v></c:pt></c:numCache></c:numRef></c:xVal>"
            + "<c:yVal><c:numRef><c:numCache><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>10</c:v></c:pt><c:pt idx=\"1\"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:yVal></c:ser></c:scatterChart></c:plotArea></c:chart></c:chartSpace>";
        var md = Md(chart);

        Assert.Contains("散佈圖", md, StringComparison.Ordinal);
        Assert.Contains("|   | 點 |\n| --- | --- |\n| 1.5 | 10 |\n| 2.5 | 20 |", md, StringComparison.Ordinal);
    }

    [Fact]
    public void A_chart_is_ordered_with_other_shapes_and_counted()
    {
        var builder = new PptxBuilder().Slide(
            Sp(2, Para("圖表之後"), y: 3000000) + ChartFrame(3, "rId5", y: 1000000) + Sp(4, Para("圖表之前"), y: 100000),
            charts: [("rId5", ChartXml("c", ["a"], [("s", [1])]))]);
        var path = _env.Save(builder);
        var md = _env.Reader.Read(path).Markdown;

        Assert.True(md.IndexOf("圖表之前", StringComparison.Ordinal) < md.IndexOf("[圖表", StringComparison.Ordinal));
        Assert.True(md.IndexOf("[圖表", StringComparison.Ordinal) < md.IndexOf("圖表之後", StringComparison.Ordinal));
        Assert.Equal(1, _env.Reader.GetOutline(path).Slides[0].Charts);
    }
}

public sealed class PptOutlineAndPagingTests : IDisposable
{
    private readonly PptEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    private string ThreeSlides() => _env.Save(new PptxBuilder()
        .Slide(Title("甲") + Sp(3, Para("甲的內文"), y: 900000) + Pic(4, null, y: 2000000), notes: "甲的備註")
        .Slide(Sp(2, Para("乙沒有標題")), hidden: true)
        .Slide(Title("丙") + Table(3, 0, 900000, [Cell("a"), Cell("b")])));

    [Fact]
    public void The_outline_lists_every_slide_with_its_facts()
    {
        var outline = _env.Reader.GetOutline(ThreeSlides());

        Assert.Equal(3, outline.SlideCount);
        Assert.Equal(["甲", string.Empty, "丙"], outline.Slides.Select(s => s.Title));
        Assert.Equal([1, 2, 3], outline.Slides.Select(s => s.Number));
        Assert.Equal([false, true, false], outline.Slides.Select(s => s.Hidden));
        Assert.Equal([true, false, false], outline.Slides.Select(s => s.HasNotes));
        Assert.Equal([1, 0, 0], outline.Slides.Select(s => s.Pictures));
        Assert.Equal([0, 0, 1], outline.Slides.Select(s => s.Tables));
        Assert.EndsWith("deck.pptx", outline.Path, StringComparison.Ordinal);
    }

    [Fact]
    public void Slide_char_counts_add_up_to_the_whole_text_minus_separators()
    {
        var path = ThreeSlides();
        var outline = _env.Reader.GetOutline(path);
        var whole = _env.Reader.Read(path).Markdown;

        Assert.Equal(whole.Length, outline.Slides.Sum(s => s.Chars) + (2 * (outline.SlideCount - 1)));
        Assert.True(outline.TotalChars >= whole.Length);
    }

    [Fact]
    public void Reading_one_slide_returns_only_that_slide()
    {
        var result = _env.Reader.ReadSlides(ThreeSlides(), 1);

        Assert.StartsWith("## 投影片 1：甲", result.Markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("投影片 2", result.Markdown, StringComparison.Ordinal);
        Assert.False(result.Truncated);
        Assert.Null(result.NextOffset);
    }

    [Fact]
    public void Reading_a_range_includes_both_ends()
    {
        var result = _env.Reader.ReadSlides(ThreeSlides(), 2, 3);

        Assert.StartsWith("## 投影片 2（隱藏）", result.Markdown, StringComparison.Ordinal);
        Assert.Contains("## 投影片 3：丙", result.Markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("投影片 1", result.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void Reading_every_slide_one_by_one_equals_reading_the_whole_deck()
    {
        var path = ThreeSlides();
        var joined = string.Join("\n\n", Enumerable.Range(1, 3).Select(i => _env.Reader.ReadSlides(path, i).Markdown));

        Assert.Equal(_env.Reader.Read(path).Markdown, joined);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(4, null)]
    [InlineData(2, 1)]
    [InlineData(1, 4)]
    [InlineData(-1, null)]
    public void A_bad_slide_range_is_rejected_with_the_slide_count(int from, int? to)
    {
        var ex = Throws(() => _env.Reader.ReadSlides(ThreeSlides(), from, to));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
        Assert.Contains("3", ex.Message + ex.Hint, StringComparison.Ordinal);
    }

    [Fact]
    public void Reading_slides_of_an_empty_presentation_is_rejected()
    {
        var ex = Throws(() => _env.Reader.ReadSlides(_env.Save(new PptxBuilder()), 1));

        Assert.Equal(ErrorCodes.InvalidValue, ex.Code);
    }

    [Fact]
    public void A_long_slide_is_paged_with_next_offset()
    {
        var env = new PptEnv(maxChars: 100);
        using (env)
        {
            var text = string.Concat(Enumerable.Range(1, 40).Select(i => Para($"第 {i} 段的內容文字")));
            var path = env.Save(new PptxBuilder().Slide(Title("長") + Sp(3, text, txBox: true, y: 900000)));

            var sb = new StringBuilder();
            var offset = 0;
            var pages = 0;
            while (true)
            {
                var page = env.Reader.ReadSlides(path, 1, offset: offset);
                Assert.True(page.Markdown.Length <= 100);
                sb.Append(page.Markdown).Append("\n\n");
                pages++;
                if (!page.Truncated)
                {
                    break;
                }

                Assert.True(page.NextOffset > offset);
                offset = page.NextOffset!.Value;
            }

            Assert.True(pages > 3);
            for (var i = 1; i <= 40; i++)
            {
                Assert.Contains($"第 {i} 段的內容文字", sb.ToString(), StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void An_offset_out_of_range_or_a_bad_max_is_rejected()
    {
        var path = ThreeSlides();

        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Reader.ReadSlides(path, 1, offset: 100000)).Code);
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Reader.ReadSlides(path, 1, offset: -1)).Code);
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Reader.Read(path, maxChars: 0)).Code);
    }

    [Fact]
    public void Whole_deck_paging_stops_on_paragraph_boundaries()
    {
        var path = ThreeSlides();
        var first = _env.Reader.Read(path, maxChars: 60);

        Assert.True(first.Truncated);
        Assert.True(first.Markdown.Length <= 60);
        var second = _env.Reader.Read(path, offset: first.NextOffset!.Value);
        Assert.Equal(_env.Reader.Read(path).Markdown, first.Markdown + "\n\n" + second.Markdown);
    }
}

public sealed class PptErrorsAndCacheTests : IDisposable
{
    private readonly PptEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    private string Good(string name = "deck.pptx") => _env.Save(new PptxBuilder().Slide(Title("T") + Sp(3, Para("x"), y: 900000)), name);

    private static byte[] Without(byte[] zipBytes, string entryName)
    {
        using var ms = new MemoryStream();
        ms.Write(zipBytes);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            zip.GetEntry(entryName)!.Delete();
        }

        return ms.ToArray();
    }

    [Fact]
    public void A_missing_file_returns_FILE_NOT_FOUND() =>
        Assert.Equal(ErrorCodes.FileNotFound, Throws(() => _env.Reader.Read(_env.Path("missing.pptx"))).Code);

    [Fact]
    public void A_path_outside_the_roots_is_rejected()
    {
        var outside = Path.Combine(Path.GetTempPath(), "OfficeToolsTests", Guid.NewGuid().ToString("N"), "secret.pptx");
        Directory.CreateDirectory(Path.GetDirectoryName(outside)!);
        try
        {
            new PptxBuilder().Slide(Title("x")).Save(outside);
            Assert.Equal(ErrorCodes.PathNotAllowed, Throws(() => _env.Reader.GetOutline(outside)).Code);
        }
        finally
        {
            Directory.Delete(Path.GetDirectoryName(outside)!, recursive: true);
        }
    }

    [Theory]
    [InlineData("book.xlsx")]
    [InlineData("doc.docx")]
    [InlineData("report.pdf")]
    [InlineData("old.ppt")]
    public void Other_formats_are_unsupported_with_advice(string name)
    {
        File.WriteAllText(_env.Path(name), "x");
        var ex = Throws(() => _env.Reader.Read(_env.Path(name)));

        Assert.Equal(ErrorCodes.UnsupportedFormat, ex.Code);
        Assert.False(string.IsNullOrEmpty(ex.Hint));
    }

    [Fact]
    public void Password_protected_files_are_reported()
    {
        var ole = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0 };
        File.WriteAllBytes(_env.Path("locked.pptx"), ole);

        Assert.Equal(ErrorCodes.PasswordProtected, Throws(() => _env.Reader.Read(_env.Path("locked.pptx"))).Code);
    }

    [Fact]
    public void Garbage_content_is_a_corrupt_file()
    {
        File.WriteAllText(_env.Path("bad.pptx"), "這不是 zip");

        Assert.Equal(ErrorCodes.CorruptFile, Throws(() => _env.Reader.Read(_env.Path("bad.pptx"))).Code);
    }

    [Fact]
    public void A_package_without_the_presentation_part_is_corrupt()
    {
        File.WriteAllBytes(_env.Path("nopres.pptx"), Without(new PptxBuilder().Slide(Title("T")).Build(), "ppt/presentation.xml"));

        Assert.Equal(ErrorCodes.CorruptFile, Throws(() => _env.Reader.Read(_env.Path("nopres.pptx"))).Code);
    }

    [Fact]
    public void A_slide_with_broken_xml_is_a_corrupt_file()
    {
        var builder = new PptxBuilder().Slide(Title("T"));
        var bytes = Without(builder.Build(), "ppt/slides/slide1.xml");
        using (var ms = new MemoryStream())
        {
            ms.Write(bytes);
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
            {
                using var writer = new StreamWriter(zip.CreateEntry("ppt/slides/slide1.xml").Open(), new UTF8Encoding(false));
                writer.Write("<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld>");
            }

            File.WriteAllBytes(_env.Path("broken.pptx"), ms.ToArray());
        }

        Assert.Equal(ErrorCodes.CorruptFile, Throws(() => _env.Reader.Read(_env.Path("broken.pptx"))).Code);
    }

    [Fact]
    public void A_slide_listed_in_the_presentation_but_missing_from_the_package_is_corrupt()
    {
        var bytes = Without(new PptxBuilder().Slide(Title("甲")).Slide(Title("乙")).Build(), "ppt/slides/slide1.xml");
        File.WriteAllBytes(_env.Path("gap.pptx"), bytes);

        Assert.Equal(ErrorCodes.CorruptFile, Throws(() => _env.Reader.Read(_env.Path("gap.pptx"))).Code);
    }

    [Fact]
    public void Too_large_uncompressed_content_is_refused()
    {
        var options = new OfficeTools.Common.OfficeToolsOptions { AllowedRoots = [_env.Root], MaxUncompressedMb = 1 };
        var reader = new DocumentReader.Core.Ppt.PptReader(new OfficeTools.Common.Security.PathGuard(options), options, new DocumentReaderOptions());
        var padding = new string('x', 3 * 1024 * 1024);
        File.WriteAllBytes(_env.Path("bomb.pptx"), new PptxBuilder().Slide(Title("T")).Entry("ppt/media/pad.xml", "<a>" + padding + "</a>").Build());

        Assert.Equal(ErrorCodes.FileTooLarge, Throws(() => reader.Read(_env.Path("bomb.pptx"))).Code);
    }

    [Fact]
    public void A_part_above_the_character_limit_is_refused()
    {
        var options = new OfficeTools.Common.OfficeToolsOptions { AllowedRoots = [_env.Root], MaxCharactersInPart = 2000 };
        var reader = new DocumentReader.Core.Ppt.PptReader(new OfficeTools.Common.Security.PathGuard(options), options, new DocumentReaderOptions());
        var big = string.Concat(Enumerable.Range(0, 200).Select(i => Para($"長文字 {i}")));
        File.WriteAllBytes(_env.Path("big.pptx"), new PptxBuilder().Slide(Sp(2, big)).Build());

        Assert.Equal(ErrorCodes.FileTooLarge, Throws(() => reader.Read(_env.Path("big.pptx"))).Code);
    }

    [Fact]
    public void Parsing_is_cached_until_the_file_changes()
    {
        var path = Good();
        _ = _env.Reader.Read(path);
        _ = _env.Reader.GetOutline(path);
        _ = _env.Reader.ReadSlides(path, 1);
        Assert.Equal(1, _env.Cache.Count);

        new PptxBuilder().Slide(Title("改過了，而且比較長一點點")).Save(path);
        Assert.Contains("改過了", _env.Reader.Read(path).Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void A_failed_parse_is_not_cached()
    {
        File.WriteAllText(_env.Path("x.pptx"), "bad");
        Throws(() => _env.Reader.Read(_env.Path("x.pptx")));

        Assert.Equal(0, _env.Cache.Count);
    }

    [Fact]
    public void The_file_is_not_locked_after_reading()
    {
        var path = Good();
        _ = _env.Reader.Read(path);

        File.Delete(path); // 還開著檔案的話在 Windows 上會失敗
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void Concurrent_reads_parse_once_and_agree()
    {
        var path = Good();
        var results = Enumerable.Range(0, 16).AsParallel().Select(_ => _env.Reader.Read(path).Markdown).ToList();

        Assert.All(results, r => Assert.Equal(results[0], r));
        Assert.Equal(1, _env.Cache.Count);
    }
}
