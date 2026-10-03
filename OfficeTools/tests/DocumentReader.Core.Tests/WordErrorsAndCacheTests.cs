using System.IO.Compression;
using System.Text;
using DocumentReader.Core.Caching;
using DocumentReader.Core.Markdown;
using DocumentReader.Core.Models;
using OfficeTools.Common.Errors;
using static DocumentReader.Core.Tests.DocxBuilder;

namespace DocumentReader.Core.Tests;

public sealed class WordErrorsTests : IDisposable
{
    private readonly WordEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static OfficeToolException Throws(Action act) => Assert.Throws<OfficeToolException>(act);

    private string Good(string name = "doc.docx") => _env.Save(new DocxBuilder().Add(H(1, "標題"), Para("內文")), name);

    private static byte[] Edit(byte[] zipBytes, string entryName, Func<string, string>? edit)
    {
        using var ms = new MemoryStream();
        ms.Write(zipBytes);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry(entryName)!;
            string xml;
            using (var reader = new StreamReader(entry.Open()))
            {
                xml = reader.ReadToEnd();
            }

            entry.Delete();
            if (edit is not null)
            {
                using var writer = new StreamWriter(zip.CreateEntry(entryName).Open(), new UTF8Encoding(false));
                writer.Write(edit(xml));
            }
        }

        return ms.ToArray();
    }

    [Fact]
    public void A_missing_file_returns_FILE_NOT_FOUND() =>
        Assert.Equal(ErrorCodes.FileNotFound, Throws(() => _env.Reader.Read(_env.Path("missing.docx"))).Code);

    [Fact]
    public void A_path_outside_the_roots_is_rejected()
    {
        var outside = _env.Outside("secret.docx");
        new DocxBuilder().Add(Para("x")).Save(outside);

        Assert.Equal(ErrorCodes.PathNotAllowed, Throws(() => _env.Reader.GetOutline(outside)).Code);
        Assert.Equal(ErrorCodes.PathNotAllowed, Throws(() => _env.Reader.ReadSection(outside, "s1")).Code);
    }

    [Theory]
    [InlineData("book.xlsx")]
    [InlineData("slides.pptx")]
    [InlineData("report.pdf")]
    [InlineData("old.doc")]
    [InlineData("notes.txt")]
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
        File.WriteAllBytes(_env.Path("locked.docx"), [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0]);
        Assert.Equal(ErrorCodes.PasswordProtected, Throws(() => _env.Reader.Read(_env.Path("locked.docx"))).Code);
    }

    [Theory]
    [InlineData("not a zip at all")]
    [InlineData("")]
    public void Garbage_is_corrupt(string content)
    {
        File.WriteAllText(_env.Path("bad.docx"), content);
        Assert.Equal(ErrorCodes.CorruptFile, Throws(() => _env.Reader.Read(_env.Path("bad.docx"))).Code);
    }

    [Fact]
    public void A_truncated_zip_is_corrupt()
    {
        var path = Good();
        var bytes = File.ReadAllBytes(path);
        File.WriteAllBytes(path, bytes[..(bytes.Length / 2)]);
        Assert.Equal(ErrorCodes.CorruptFile, Throws(() => _env.Reader.Read(path)).Code);
    }

    [Fact]
    public void A_zip_without_the_document_part_is_corrupt()
    {
        var path = Good();
        File.WriteAllBytes(path, Edit(File.ReadAllBytes(path), "word/document.xml", edit: null));
        Assert.Equal(ErrorCodes.CorruptFile, Throws(() => _env.Reader.Read(path)).Code);
    }

    [Fact]
    public void Malformed_xml_in_the_document_is_corrupt_not_a_crash()
    {
        var path = Good();
        File.WriteAllBytes(path, Edit(File.ReadAllBytes(path), "word/document.xml", xml => xml.Replace("</w:body>", "<w:p>", StringComparison.Ordinal)));
        Assert.Equal(ErrorCodes.CorruptFile, Throws(() => _env.Reader.Read(path)).Code);
    }

    [Fact]
    public void A_zip_bomb_is_rejected_by_uncompressed_size()
    {
        using var env = new WordEnv(maxUncompressedMb: 1);
        var path = env.Save(new DocxBuilder().Add(Para("x")));
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            using var s = zip.CreateEntry("word/filler.bin").Open();
            s.Write(new byte[3 * 1024 * 1024]);
        }

        Assert.True(new FileInfo(path).Length < 100_000);
        Assert.Equal(ErrorCodes.FileTooLarge, Throws(() => env.Reader.Read(path)).Code);
    }

    [Fact]
    public void A_part_with_too_many_characters_is_rejected()
    {
        using var env = new WordEnv(maxCharactersInPart: 2_000);
        var path = env.Save(new DocxBuilder().Add(Enumerable.Range(0, 200).Select(i => Para($"第 {i} 段的內容文字")).ToArray()));

        var ex = Throws(() => env.Reader.Read(path));

        Assert.Equal(ErrorCodes.FileTooLarge, ex.Code);
    }

    [Fact]
    public void Reading_does_not_keep_the_file_locked_or_change_it()
    {
        var path = Good();
        var before = File.ReadAllBytes(path);
        var modified = File.GetLastWriteTimeUtc(path);

        _env.Reader.Read(path);
        _env.Reader.GetOutline(path);

        using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        Assert.True(exclusive.CanWrite);
        exclusive.Close();
        Assert.Equal(before, File.ReadAllBytes(path));
        Assert.Equal(modified, File.GetLastWriteTimeUtc(path));
    }

    [Fact]
    public void An_exclusively_locked_file_returns_FILE_LOCKED()
    {
        var path = Good();
        using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        Assert.Equal(ErrorCodes.FileLocked, Throws(() => _env.Reader.Read(path)).Code);
    }

    [Fact]
    public void A_file_without_the_main_part_relationship_is_corrupt_not_an_unhandled_exception()
    {
        var path = Good();
        File.WriteAllBytes(path, Edit(File.ReadAllBytes(path), "_rels/.rels", xml => xml.Replace("officeDocument", "somethingElse", StringComparison.Ordinal)));
        Assert.Equal(ErrorCodes.CorruptFile, Throws(() => _env.Reader.Read(path)).Code);
    }
}

public sealed class WordCacheTests : IDisposable
{
    private readonly WordEnv _env = new(cacheEntries: 2);

    public void Dispose() => _env.Dispose();

    private string Doc(string name, string text) => _env.Save(new DocxBuilder().Add(H(1, name), Para(text)), name);

    [Fact]
    public void Repeated_reads_of_the_same_document_use_one_cache_entry()
    {
        var path = Doc("a.docx", "內容");

        _env.Reader.Read(path);
        _env.Reader.GetOutline(path);
        _env.Reader.ReadSection(path, "s1");

        Assert.Equal(1, _env.Cache.Count);
    }

    [Fact]
    public void Different_read_options_are_cached_separately()
    {
        var path = Doc("a.docx", "內容");

        _env.Reader.Read(path);
        _env.Reader.Read(path, wordOptions: new WordReadOptions(IncludeFootnotes: true));

        Assert.Equal(2, _env.Cache.Count);
    }

    [Fact]
    public void A_modified_file_is_read_again_and_new_content_shows_up()
    {
        var path = Doc("a.docx", "舊內容");
        Assert.Contains("舊內容", _env.Reader.Read(path).Markdown, StringComparison.Ordinal);

        new DocxBuilder().Add(H(1, "a.docx"), Para("新內容比較長一點點")).Save(path);
        File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddMinutes(5));

        Assert.Contains("新內容", _env.Reader.Read(path).Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void A_changed_size_alone_also_invalidates_the_entry()
    {
        var path = Doc("a.docx", "舊");
        var stamp = File.GetLastWriteTimeUtc(path);
        _env.Reader.Read(path);

        new DocxBuilder().Add(H(1, "a.docx"), Para("新的而且更長")).Save(path);
        File.SetLastWriteTimeUtc(path, stamp); // 修改時間相同，只有大小不同

        Assert.Contains("新的而且更長", _env.Reader.Read(path).Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void The_cache_is_bounded_and_evicts_the_least_recently_used_document()
    {
        var a = Doc("a.docx", "A");
        var b = Doc("b.docx", "B");
        var c = Doc("c.docx", "C");

        _env.Reader.Read(a);
        _env.Reader.Read(b);
        _env.Reader.Read(a); // a 比 b 新
        _env.Reader.Read(c); // 超過 2 筆，b 被淘汰

        Assert.Equal(2, _env.Cache.Count);
        File.Delete(b); // 如果 b 還在快取裡，刪檔後仍讀得到；不在的話會 FILE_NOT_FOUND
        Assert.Equal(ErrorCodes.FileNotFound, Assert.Throws<OfficeToolException>(() => _env.Reader.Read(b)).Code);
    }

    [Fact]
    public void A_failed_parse_is_not_cached_and_succeeds_once_the_file_is_fixed()
    {
        var path = _env.Path("later.docx");
        File.WriteAllText(path, "壞掉的檔案");
        Assert.Throws<OfficeToolException>(() => _env.Reader.Read(path));
        Assert.Equal(0, _env.Cache.Count);

        new DocxBuilder().Add(Para("修好了")).Save(path);
        Assert.Equal("修好了", _env.Reader.Read(path).Markdown);
    }
}

public class ParsedDocumentCacheTests
{
    private static ParsedDocumentCache.Key K(string name, long ticks = 1) => new(name, ticks, 10, "v");

    [Fact]
    public void The_factory_runs_once_per_key_even_with_many_concurrent_callers()
    {
        var cache = new ParsedDocumentCache(10);
        var calls = 0;

        var results = Enumerable.Range(0, 32).AsParallel().Select(_ => cache.GetOrAdd(K("a"), () =>
        {
            Interlocked.Increment(ref calls);
            Thread.Sleep(20);
            return new object();
        })).ToArray();

        Assert.Equal(1, calls);
        Assert.All(results, r => Assert.Same(results[0], r));
    }

    [Fact]
    public void Keys_that_differ_in_any_part_are_separate_entries()
    {
        var cache = new ParsedDocumentCache(10);
        var a = cache.GetOrAdd(K("a"), () => new object());

        Assert.NotSame(a, cache.GetOrAdd(K("b"), () => new object()));
        Assert.NotSame(a, cache.GetOrAdd(K("a", ticks: 2), () => new object()));
        Assert.NotSame(a, cache.GetOrAdd(new ParsedDocumentCache.Key("a", 1, 11, "v"), () => new object()));
        Assert.NotSame(a, cache.GetOrAdd(new ParsedDocumentCache.Key("a", 1, 10, "other"), () => new object()));
        Assert.Same(a, cache.GetOrAdd(K("a"), () => new object()));
    }

    [Fact]
    public void Eviction_follows_least_recently_used_order()
    {
        var cache = new ParsedDocumentCache(2);
        var a = cache.GetOrAdd(K("a"), () => new object());
        cache.GetOrAdd(K("b"), () => new object());
        cache.GetOrAdd(K("a"), () => new object()); // a 變成最近使用
        cache.GetOrAdd(K("c"), () => new object()); // 淘汰 b

        Assert.Equal(2, cache.Count);
        Assert.Same(a, cache.GetOrAdd(K("a"), () => new object()));
        var rebuilt = false;
        cache.GetOrAdd(K("b"), () =>
        {
            rebuilt = true;
            return new object();
        });
        Assert.True(rebuilt);
    }

    [Fact]
    public void A_throwing_factory_leaves_nothing_behind_and_is_retried()
    {
        var cache = new ParsedDocumentCache(2);
        Assert.Throws<InvalidOperationException>(() => cache.GetOrAdd<object>(K("a"), () => throw new InvalidOperationException("boom")));
        Assert.Equal(0, cache.Count);
        Assert.NotNull(cache.GetOrAdd(K("a"), () => new object()));
    }

    [Fact]
    public void Clear_empties_the_cache_and_a_capacity_below_one_still_keeps_one_entry()
    {
        var cache = new ParsedDocumentCache(0);
        cache.GetOrAdd(K("a"), () => new object());
        cache.GetOrAdd(K("b"), () => new object());
        Assert.Equal(1, cache.Count);

        cache.Clear();
        Assert.Equal(0, cache.Count);
    }
}

public class MarkdownWriterTests
{
    private static string Inline(bool emphasis, params InlinePiece[] pieces) =>
        MarkdownWriter.RenderInline(pieces, emphasis, "  \n", inTable: false);

    [Theory]
    [InlineData("plain", "plain")]
    [InlineData("a*b", "a\\*b")]
    [InlineData("a`b", "a\\`b")]
    [InlineData("a\\b", "a\\\\b")]
    [InlineData("a_b [c] #d", "a_b [c] #d")]
    public void Escape_only_touches_characters_that_would_be_misread(string text, string expected) =>
        Assert.Equal(expected, MarkdownWriter.Escape(text));

    [Fact]
    public void Pipes_are_escaped_only_inside_tables()
    {
        Assert.Equal("a|b", MarkdownWriter.Escape("a|b"));
        Assert.Equal("a\\|b", MarkdownWriter.Escape("a|b", inTable: true));
    }

    [Fact]
    public void Inline_rendering_merges_pieces_with_identical_formatting()
    {
        Assert.Equal("**ab**", Inline(true, new InlinePiece("a", Bold: true), new InlinePiece("b", Bold: true)));
        Assert.Equal("**a***b*", Inline(true, new InlinePiece("a", Bold: true), new InlinePiece("b", Italic: true)));
    }

    [Fact]
    public void Emphasis_can_be_turned_off_and_whitespace_only_pieces_are_never_wrapped()
    {
        Assert.Equal("a b", Inline(false, new InlinePiece("a ", Bold: true), new InlinePiece("b", Italic: true)));
        Assert.Equal("a **b** c", Inline(true, new InlinePiece("a "), new InlinePiece("b", Bold: true), new InlinePiece(" c")));
        Assert.Equal("a  c", Inline(true, new InlinePiece("a"), new InlinePiece("  ", Bold: true), new InlinePiece("c")));
    }

    [Fact]
    public void Strikethrough_wraps_outside_the_other_markers() =>
        Assert.Equal("~~**x**~~", Inline(true, new InlinePiece("x", Bold: true, Strike: true)));

    [Fact]
    public void Newlines_in_pieces_use_the_requested_line_break()
    {
        Assert.Equal("a  \nb", Inline(true, new InlinePiece("a\nb")));
        Assert.Equal("a<br>b", MarkdownWriter.RenderInline([new InlinePiece("a\nb")], true, "<br>", inTable: true));
    }

    [Fact]
    public void A_link_around_whitespace_only_text_is_dropped_instead_of_producing_an_empty_link() =>
        Assert.Equal("  ", Inline(true, new InlinePiece("  ", Link: "https://x.com")));

    [Fact]
    public void Adjacent_pieces_with_different_links_become_separate_links() =>
        Assert.Equal("[a](https://1.com)[b](https://2.com)", Inline(true, new InlinePiece("a", Link: "https://1.com"), new InlinePiece("b", Link: "https://2.com")));

    [Fact]
    public void Tables_pad_rows_and_fill_empty_cells_with_a_space()
    {
        var table = MarkdownWriter.RenderTable([["a", "b", "c"], ["1"], ["", "2", "3"]]);
        Assert.Equal("| a | b | c |\n| --- | --- | --- |\n| 1 |   |   |\n|   | 2 | 3 |", table);
    }

    [Fact]
    public void The_builder_separates_blocks_with_blank_lines_and_list_items_with_single_newlines()
    {
        var builder = new MarkdownDocumentBuilder();
        builder.AddParagraph("p1");
        builder.AddListItem("- a", 1);
        builder.AddListItem("- b", 1);
        builder.AddListItem("- c", 2);
        builder.AddParagraph("p2");

        Assert.Equal("p1\n\n- a\n- b\n\n- c\n\np2", builder.Build([]).Text);
    }

    [Fact]
    public void The_builder_ignores_blank_blocks_and_computes_section_ranges()
    {
        var builder = new MarkdownDocumentBuilder();
        builder.AddParagraph("   ");
        builder.AddHeading(1, "A");
        builder.AddParagraph("body");
        builder.AddHeading(2, "B");
        builder.AddHeading(1, "C");
        var doc = builder.Build([]);

        Assert.Equal("# A\n\nbody\n\n## B\n\n# C", doc.Text);
        var a = doc.Sections[0];
        Assert.Equal("# A\n\nbody\n\n## B", doc.Text[a.Start..a.TotalEnd]);
        Assert.Equal("# A\n\nbody", doc.Text[a.Start..a.OwnEnd]);
        Assert.Equal("## B", doc.Text[doc.Sections[1].Start..doc.Sections[1].TotalEnd]);
        Assert.Equal("# C", doc.Text[doc.Sections[2].Start..doc.Sections[2].TotalEnd]);
        Assert.Equal("s1", doc.Sections[1].ParentId);
        Assert.Null(doc.Sections[2].ParentId);
    }
}
