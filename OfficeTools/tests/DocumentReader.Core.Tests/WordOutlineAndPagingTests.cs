using DocumentReader.Core.Models;
using OfficeTools.Common.Errors;
using static DocumentReader.Core.Tests.DocxBuilder;

namespace DocumentReader.Core.Tests;

public sealed class WordOutlineAndPagingTests : IDisposable
{
    private readonly WordEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    private string Save(params string[] blocks) => _env.Save(new DocxBuilder().Add(blocks));

    private static IEnumerable<OutlineNode> Flatten(IEnumerable<OutlineNode> nodes) =>
        nodes.SelectMany(n => new[] { n }.Concat(Flatten(n.Children)));

    /// <summary>用另一個不限長度的 reader 讀同一份檔案，當作「完整內容」的對照（被測的 reader 有讀取上限，會截斷）。</summary>
    private static string Whole(string path, string? sectionId = null)
    {
        using var big = new WordEnv();
        File.Copy(path, big.Path("whole.docx"));
        return sectionId is null ? big.Reader.Read(big.Path("whole.docx")).Markdown : big.Reader.ReadSection(big.Path("whole.docx"), sectionId).Markdown;
    }

    private static string Shape(IEnumerable<OutlineNode> nodes) =>
        string.Join(",", nodes.Select(n => n.Children.Count == 0 ? n.SectionId : $"{n.SectionId}({Shape(n.Children)})"));

    /// <summary>前言 + 兩章（第一章有兩節，其中一節有小節）。</summary>
    private string Book() => Save(
        Para("前言文字"),
        H(1, "第一章"), Para("一章內文"),
        H(2, "第一節"), Para("一節內文"),
        H(3, "小節"), Para("小節內文"),
        H(2, "第二節"), Para("二節內文"),
        H(1, "第二章"), Para("二章內文"));

    // ---- 標題樹 ----

    [Fact]
    public void The_outline_is_a_tree_with_ids_in_document_order_and_a_preface_node()
    {
        var outline = _env.Reader.GetOutline(Book());

        Assert.True(outline.HasHeadings);
        Assert.Equal("s0,s1(s2(s3),s4),s5", Shape(outline.Sections));
        Assert.Equal(["(前言)", "第一章", "第一節", "小節", "第二節", "第二章"], Flatten(outline.Sections).Select(n => n.Title));
        Assert.Equal([0, 1, 2, 3, 2, 1], Flatten(outline.Sections).Select(n => n.Level));
        Assert.EndsWith("doc.docx", outline.Path, StringComparison.Ordinal);
    }

    [Fact]
    public void Each_node_reports_its_own_characters_and_the_total_including_subsections()
    {
        var path = Book();
        var outline = _env.Reader.GetOutline(path);
        var nodes = Flatten(outline.Sections).ToDictionary(n => n.SectionId);

        Assert.Equal("前言文字".Length, nodes["s0"].Chars);
        Assert.Equal("# 第一章\n\n一章內文".Length, nodes["s1"].Chars);
        Assert.Equal("## 第一節\n\n一節內文\n\n### 小節\n\n小節內文".Length, nodes["s2"].TotalChars);
        Assert.Equal(nodes["s1"].Chars + nodes["s2"].TotalChars + nodes["s4"].TotalChars, nodes["s1"].TotalChars - 4); // 子節之間各隔一個空行（2 字元）
        Assert.Equal(outline.TotalChars, _env.Reader.Read(path).Markdown.Length);
    }

    [Fact]
    public void A_document_without_headings_has_only_the_preface_and_says_so()
    {
        var outline = _env.Reader.GetOutline(Save(Para("只有"), Para("段落")));

        Assert.False(outline.HasHeadings);
        var only = Assert.Single(outline.Sections);
        Assert.Equal("s0", only.SectionId);
        Assert.Equal("只有\n\n段落".Length, only.Chars);
    }

    [Fact]
    public void A_document_that_starts_with_a_heading_has_no_preface()
    {
        var outline = _env.Reader.GetOutline(Save(H(1, "開頭"), Para("內文")));
        Assert.Equal("s1", Assert.Single(outline.Sections).SectionId);
    }

    [Fact]
    public void Skipped_heading_levels_nest_under_the_nearest_shallower_heading()
    {
        var outline = _env.Reader.GetOutline(Save(H(1, "A"), H(3, "跳級"), H(2, "B"), H(1, "C"), H(3, "又跳級")));
        Assert.Equal("s1(s2,s3),s4(s5)", Shape(outline.Sections));
        Assert.Equal(3, Flatten(outline.Sections).First(n => n.Title == "跳級").Level); // 保留原本的層級
    }

    [Fact]
    public void A_deeper_first_heading_is_still_a_top_level_node()
    {
        var outline = _env.Reader.GetOutline(Save(H(2, "先出現二級"), H(1, "再出現一級")));
        Assert.Equal("s1,s2", Shape(outline.Sections));
    }

    [Fact]
    public void Heading_titles_are_single_line_and_whitespace_normalized()
    {
        var outline = _env.Reader.GetOutline(Save(P(Run("多行") + Br() + Run("  標題   文字  "), style: "1")));
        Assert.Equal("多行 標題 文字", Assert.Single(outline.Sections).Title);
    }

    [Fact]
    public void Outline_carries_the_conversion_notes()
    {
        var outline = _env.Reader.GetOutline(Save(H(1, "章"), P(Image("圖"))));
        Assert.Contains(outline.Notes, n => n.Contains("1 張圖片", StringComparison.Ordinal));
    }

    // ---- 讀一節 ----

    [Fact]
    public void A_section_includes_its_heading_and_subsections_by_default()
    {
        var result = _env.Reader.ReadSection(Book(), "s2");

        Assert.Equal("## 第一節\n\n一節內文\n\n### 小節\n\n小節內文", result.Markdown);
        Assert.Equal("s2", result.SectionId);
        Assert.False(result.Truncated);
        Assert.Null(result.NextOffset);
        Assert.Equal(result.Markdown.Length, result.TotalChars);
    }

    [Fact]
    public void A_section_can_be_read_without_its_subsections()
    {
        var result = _env.Reader.ReadSection(Book(), "s2", includeSubsections: false);
        Assert.Equal("## 第一節\n\n一節內文", result.Markdown);
    }

    [Fact]
    public void A_top_level_section_covers_everything_up_to_the_next_top_level_heading()
    {
        var path = Book();
        Assert.Equal("# 第一章\n\n一章內文\n\n## 第一節\n\n一節內文\n\n### 小節\n\n小節內文\n\n## 第二節\n\n二節內文", _env.Reader.ReadSection(path, "s1").Markdown);
        Assert.Equal("# 第二章\n\n二章內文", _env.Reader.ReadSection(path, "s5").Markdown);
        Assert.Equal("前言文字", _env.Reader.ReadSection(path, "s0").Markdown);
    }

    [Fact]
    public void Reading_every_top_level_section_in_order_reproduces_the_whole_document()
    {
        var path = Book();
        var outline = _env.Reader.GetOutline(path);

        var joined = string.Join("\n\n", outline.Sections.Select(n => _env.Reader.ReadSection(path, n.SectionId).Markdown));

        Assert.Equal(_env.Reader.Read(path).Markdown, joined);
    }

    [Theory]
    [InlineData("S2")]
    [InlineData(" s2 ")]
    public void Section_ids_are_case_and_space_insensitive(string id) =>
        Assert.StartsWith("## 第一節", _env.Reader.ReadSection(Book(), id).Markdown, StringComparison.Ordinal);

    [Theory]
    [InlineData("s99")]
    [InlineData("")]
    [InlineData("第一章")]
    public void An_unknown_section_lists_the_existing_ones(string id)
    {
        var ex = Throws(() => _env.Reader.ReadSection(Book(), id));

        Assert.Equal(ErrorCodes.SectionNotFound, ex.Code);
        Assert.Contains("s1 第一章", ex.Hint, StringComparison.Ordinal);
        Assert.Contains("s4 第二節", ex.Hint, StringComparison.Ordinal);
    }

    [Fact]
    public void A_long_section_list_in_the_hint_is_capped()
    {
        var blocks = Enumerable.Range(1, 40).Select(i => H(1, $"章{i}")).ToArray();
        var ex = Throws(() => _env.Reader.ReadSection(Save(blocks), "nope"));

        Assert.Contains("共 40 節", ex.Hint, StringComparison.Ordinal);
        Assert.DoesNotContain("s40", ex.Hint, StringComparison.Ordinal);
    }

    [Fact]
    public void Reading_a_section_of_an_empty_document_says_there_is_nothing_to_read()
    {
        var ex = Throws(() => _env.Reader.ReadSection(Save(), "s1"));
        Assert.Equal(ErrorCodes.SectionNotFound, ex.Code);
        Assert.Contains("沒有內容", ex.Hint, StringComparison.Ordinal);
    }

    // ---- 分頁 ----

    private string Paragraphs(int count, int length, out string[] texts)
    {
        texts = Enumerable.Range(0, count).Select(i => $"第{i:00}段" + new string((char)('a' + (i % 26)), length - 4)).ToArray();
        return Save(texts.Select(t => Para(t)).ToArray());
    }

    [Fact]
    public void Paging_stops_at_paragraph_boundaries_and_the_chunks_rebuild_the_document()
    {
        using var env = new WordEnv(maxChars: 200);
        var texts = Enumerable.Range(0, 20).Select(i => $"第{i:00}段" + new string((char)('a' + (i % 26)), 46)).ToArray();
        var path = env.Save(new DocxBuilder().Add(texts.Select(t => Para(t)).ToArray()));

        var chunks = new List<ReadResult>();
        var offset = 0;
        do
        {
            var chunk = env.Reader.Read(path, offset);
            chunks.Add(chunk);
            offset = chunk.NextOffset ?? 0;
        }
        while (chunks[^1].Truncated);

        Assert.True(chunks.Count > 4);
        Assert.All(chunks, c => Assert.True(c.Markdown.Length <= 200));
        // 每個區塊都只含完整的段落（沒有在段落中間被切開）
        foreach (var chunk in chunks)
        {
            foreach (var paragraph in chunk.Markdown.Split("\n\n"))
            {
                Assert.Contains(paragraph, texts);
            }
        }

        Assert.Equal(string.Join("\n\n", texts), string.Join("\n\n", chunks.Select(c => c.Markdown)));
        Assert.Null(chunks[^1].NextOffset);
        Assert.Equal(chunks.Count - 1, chunks.Count(c => c.Truncated));
    }

    [Fact]
    public void Offsets_in_the_result_describe_where_the_chunk_sits_in_the_document()
    {
        using var env = new WordEnv(maxChars: 120);
        var path = env.Save(new DocxBuilder().Add(Enumerable.Range(0, 10).Select(i => Para(new string('x', 50))).ToArray()));
        var full = Whole(path);

        var first = env.Reader.Read(path);
        var second = env.Reader.Read(path, first.NextOffset!.Value);

        Assert.Equal(0, first.StartOffset);
        Assert.Equal(first.EndOffset, first.NextOffset);
        Assert.Equal(first.NextOffset, second.StartOffset);
        Assert.Equal(full.Length, first.TotalChars);
        Assert.Equal(full[first.StartOffset..first.EndOffset].TrimEnd('\n'), first.Markdown);
    }

    [Fact]
    public void A_single_paragraph_longer_than_the_limit_is_cut_hard_and_continues_exactly()
    {
        using var env = new WordEnv(maxChars: 100);
        var text = string.Concat(Enumerable.Range(0, 60).Select(i => $"{i:00}字")); // 一個 180 字元的段落，沒有任何邊界
        var path = env.Save(new DocxBuilder().Add(Para(text)));

        var parts = new List<string>();
        var offset = 0;
        for (var guard = 0; guard < 10; guard++)
        {
            var chunk = env.Reader.Read(path, offset);
            parts.Add(chunk.Markdown);
            if (!chunk.Truncated)
            {
                break;
            }

            offset = chunk.NextOffset!.Value;
        }

        Assert.Equal(2, parts.Count);
        Assert.Equal(100, parts[0].Length);
        Assert.Equal(text, string.Concat(parts));
    }

    [Fact]
    public void A_hard_cut_never_splits_a_surrogate_pair()
    {
        using var env = new WordEnv(maxChars: 10);
        var path = env.Save(new DocxBuilder().Add(Para(new string('a', 9) + "😀" + new string('b', 5))));

        var first = env.Reader.Read(path);

        Assert.Equal(9, first.Markdown.Length); // 第 10 個字元是代理對的前半，所以停在前面
        Assert.All(first.Markdown, c => Assert.False(char.IsSurrogate(c)));
        var second = env.Reader.Read(path, first.NextOffset!.Value);
        Assert.StartsWith("😀", second.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void Paging_prefers_a_line_break_over_a_hard_cut_when_there_is_no_paragraph_boundary()
    {
        using var env = new WordEnv(maxChars: 60);
        var lines = Enumerable.Range(0, 6).Select(i => new string((char)('a' + i), 20)).ToArray();
        var path = env.Save(new DocxBuilder().Add(P(string.Join(string.Empty, lines.Select((l, i) => (i > 0 ? Br() : string.Empty) + Run(l))))));

        var first = env.Reader.Read(path);

        Assert.StartsWith(lines[0], first.Markdown, StringComparison.Ordinal);
        Assert.EndsWith(lines[1] + "  ", first.Markdown.Replace("\n", string.Empty), StringComparison.Ordinal);
    }

    [Fact]
    public void A_section_can_be_paged_with_offsets_relative_to_the_section()
    {
        using var env = new WordEnv(maxChars: 100);
        var texts = Enumerable.Range(0, 8).Select(i => $"內{i}" + new string('字', 40)).ToArray();
        var path = env.Save(new DocxBuilder().Add(H(1, "前"), Para("前文"), H(1, "目標"), P(string.Concat(texts.Select(t => Run(t)))), H(1, "後"), Para("後文")));
        var whole = Whole(path, "s2");

        var pieces = new List<string>();
        var offset = 0;
        ReadResult chunk;
        do
        {
            chunk = env.Reader.ReadSection(path, "s2", offset: offset);
            Assert.Equal("s2", chunk.SectionId);
            pieces.Add(chunk.Markdown);
            offset = chunk.NextOffset ?? offset;
        }
        while (chunk.Truncated);

        Assert.True(pieces.Count > 2);
        Assert.StartsWith("# 目標", pieces[0], StringComparison.Ordinal);
        Assert.DoesNotContain("後文", string.Concat(pieces), StringComparison.Ordinal);
        Assert.Equal(whole, string.Concat(pieces));
    }

    [Fact]
    public void A_smaller_maxChars_is_honored_and_a_larger_one_is_capped_to_the_configured_limit()
    {
        using var env = new WordEnv(maxChars: 100);
        var path = env.Save(new DocxBuilder().Add(Para(new string('x', 500))));

        Assert.Equal(30, env.Reader.Read(path, maxChars: 30).Markdown.Length);
        Assert.Equal(100, env.Reader.Read(path, maxChars: 10_000).Markdown.Length);
    }

    [Fact]
    public void Offset_equal_to_the_length_returns_an_empty_final_chunk_and_beyond_is_an_error()
    {
        var path = Save(Para("短文"));

        var end = _env.Reader.Read(path, offset: 2);
        Assert.Equal(string.Empty, end.Markdown);
        Assert.False(end.Truncated);

        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Reader.Read(path, offset: 3)).Code);
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Reader.Read(path, offset: -1)).Code);
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Reader.Read(path, maxChars: 0)).Code);
        Assert.Equal(ErrorCodes.InvalidValue, Throws(() => _env.Reader.ReadSection(Book(), "s1", offset: 99_999)).Code);
    }

    [Fact]
    public void Every_chunk_makes_progress_even_with_a_tiny_limit()
    {
        using var env = new WordEnv(maxChars: 1);
        var path = env.Save(new DocxBuilder().Add(Para("甲乙丙")));

        var text = string.Empty;
        var offset = 0;
        for (var guard = 0; guard < 10; guard++)
        {
            var chunk = env.Reader.Read(path, offset);
            text += chunk.Markdown;
            if (!chunk.Truncated)
            {
                break;
            }

            Assert.True(chunk.NextOffset > offset);
            offset = chunk.NextOffset!.Value;
        }

        Assert.Equal("甲乙丙", text);
    }
}
