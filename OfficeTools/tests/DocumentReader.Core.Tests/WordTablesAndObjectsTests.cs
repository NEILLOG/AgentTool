using DocumentReader.Core.Models;
using static DocumentReader.Core.Tests.DocxBuilder;

namespace DocumentReader.Core.Tests;

public sealed class WordTablesAndObjectsTests : IDisposable
{
    private readonly WordEnv _env = new();

    public void Dispose() => _env.Dispose();

    private string Md(params string[] blocks) => _env.Md(new DocxBuilder().Add(blocks));

    private (string Markdown, IReadOnlyList<string> Notes) Read(params string[] blocks)
    {
        var result = _env.Reader.Read(_env.Save(new DocxBuilder().Add(blocks)));
        return (result.Markdown, result.Notes);
    }

    // ---- 表格 ----

    [Fact]
    public void A_simple_table_becomes_a_markdown_table_with_the_first_row_as_header() =>
        Assert.Equal("| 名稱 | 數量 |\n| --- | --- |\n| 蘋果 | 3 |\n| 香蕉 | 5 |", Md(SimpleTable(["名稱", "數量"], ["蘋果", "3"], ["香蕉", "5"])));

    [Fact]
    public void A_single_row_table_still_gets_a_header_separator() =>
        Assert.Equal("| 甲 | 乙 |\n| --- | --- |", Md(SimpleTable(["甲", "乙"])));

    [Fact]
    public void Horizontally_merged_cells_repeat_the_value_in_every_spanned_column()
    {
        var md = Md(Table(
            Row(Cell("標題", span: 3)),
            Row(Cell("A"), Cell("B", span: 2)),
            Row(Cell("1"), Cell("2"), Cell("3"))));

        Assert.Equal("| 標題 | 標題 | 標題 |\n| --- | --- | --- |\n| A | B | B |\n| 1 | 2 | 3 |", md);
    }

    [Fact]
    public void Vertically_merged_cells_repeat_the_value_down_the_merged_rows()
    {
        var md = Md(Table(
            Row(Cell("組別"), Cell("成員")),
            Row(Cell("甲", vMerge: "restart"), Cell("小明")),
            Row(Cell("", vMerge: ""), Cell("小華")),
            Row(Cell("乙", vMerge: "restart"), Cell("小美"))));

        Assert.Equal("| 組別 | 成員 |\n| --- | --- |\n| 甲 | 小明 |\n| 甲 | 小華 |\n| 乙 | 小美 |", md);
    }

    [Fact]
    public void Rows_with_different_widths_are_padded_to_the_widest_row()
    {
        var md = Md(Table(Row(Cell("A"), Cell("B"), Cell("C")), Row(Cell("1"))));
        Assert.Equal("| A | B | C |\n| --- | --- | --- |\n| 1 |   |   |", md);
    }

    [Fact]
    public void Grid_before_and_after_pad_the_row_with_empty_cells()
    {
        var shifted = "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/><w:gridAfter w:val=\"1\"/></w:trPr>" + Cell("中") + "</w:tr>";
        var md = Md(Table(Row(Cell("A"), Cell("B"), Cell("C")), shifted));

        Assert.Equal("| A | B | C |\n| --- | --- | --- |\n|   | 中 |   |", md);
    }

    [Fact]
    public void Pipes_are_escaped_and_multiple_paragraphs_or_breaks_use_br()
    {
        var multi = Cell(Para("第一段") + Para("第二段 a|b"));
        var broken = Cell(P(Run("行一") + Br() + Run("行二")));
        var md = Md(Table(Row(Cell("H1"), Cell("H2")), Row(multi, broken)));

        Assert.Equal("| H1 | H2 |\n| --- | --- |\n| 第一段<br>第二段 a\\|b | 行一<br>行二 |", md);
    }

    [Fact]
    public void Formatting_links_and_images_inside_cells_are_kept()
    {
        var builder = new DocxBuilder();
        var link = builder.Link("https://example.com");
        builder.Add(Table(Row(Cell("H")), Row(Cell(P(Run("粗", b: true) + Run(" ") + Hyperlink(link, Run("連結")) + Image("圖示"))))));

        Assert.Equal("| H |\n| --- |\n| **粗** [連結](https://example.com)[圖片: 圖示] |", _env.Md(builder));
    }

    [Fact]
    public void A_nested_table_is_flattened_to_text()
    {
        var inner = Table(Row(Cell("a"), Cell("b")), Row(Cell("c"), Cell("d")));
        var md = Md(Table(Row(Cell("外")), Row(Cell(Para("前") + inner))));

        Assert.Equal("| 外 |\n| --- |\n| 前<br>（巢狀表格：a / b; c / d） |", md);
    }

    [Fact]
    public void List_items_inside_cells_get_markers()
    {
        var md = Md(Table(Row(Cell("H")), Row(Cell(Para("甲", numId: 1) + Para("乙", numId: 1)))));
        Assert.Equal("| H |\n| --- |\n| • 甲<br>• 乙 |", md);
    }

    [Fact]
    public void Empty_tables_are_skipped_and_text_around_tables_stays_in_order()
    {
        Assert.Equal("前\n\n| 格 |\n| --- |\n\n後", Md(Para("前"), Table(Row(Cell("格"))), Para("後")));
        Assert.Equal("前\n\n後", Md(Para("前"), Table(Row(Cell(""), Cell(" "))), Para("後")));
        Assert.Equal("前", Md(Para("前"), "<w:tbl><w:tblPr/></w:tbl>"));
    }

    [Fact]
    public void Content_controls_wrapping_tables_are_unwrapped() =>
        Assert.Equal("| 甲 |\n| --- |", Md(BlockSdt(SimpleTable(["甲"]))));

    // ---- 圖片與物件 ----

    [Fact]
    public void Images_show_their_alt_text_or_a_plain_marker_and_are_counted_in_the_notes()
    {
        var (md, notes) = Read(P(Run("前 ") + Image("公司標誌") + Run(" 中 ") + Image(null) + Run(" 後")));

        Assert.Equal("前 [圖片: 公司標誌] 中 [圖片] 後", md);
        Assert.Contains(notes, n => n.Contains("2 張圖片", StringComparison.Ordinal));
    }

    [Fact]
    public void Alt_text_is_cleaned_up_to_a_single_line() =>
        Assert.Equal("[圖片: 一 二 三]", Md(P(Image("  一\n二\t三  "))));

    [Fact]
    public void Charts_and_smartart_are_marked_and_counted()
    {
        var (md, notes) = Read(P(Chart("營收圖")), P(Chart()), P(SmartArt()));

        Assert.Equal("[圖表: 營收圖]\n\n[圖表]\n\n[SmartArt]", md);
        Assert.Contains(notes, n => n.Contains("2 個圖表", StringComparison.Ordinal));
        Assert.Contains(notes, n => n.Contains("SmartArt", StringComparison.Ordinal));
    }

    [Fact]
    public void Legacy_vml_pictures_are_images_too() =>
        Assert.Equal("[圖片: 舊版圖片]", Md(P(VmlImage("舊版圖片"))));

    [Fact]
    public void A_text_box_is_output_once_not_twice_from_choice_and_fallback()
    {
        var md = Md(P(Run("段落文字") + TextBox("方塊內的字")));
        Assert.Equal("段落文字\n\n[文字方塊] 方塊內的字", md);
    }

    [Fact]
    public void A_shape_without_text_or_picture_is_ignored()
    {
        var decorative = "<w:r><w:drawing><wp:inline><wp:docPr id=\"9\" name=\"Line\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:wsp/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
        Assert.Equal("文字", Md(P(Run("文字") + decorative)));
    }

    [Fact]
    public void Embedded_objects_are_marked() =>
        Assert.Equal("[嵌入物件]", Md(P("<w:r><w:object/></w:r>")));

    [Fact]
    public void An_image_inside_a_hyperlink_becomes_a_link_to_the_marker()
    {
        var builder = new DocxBuilder();
        builder.Add(P(Hyperlink(builder.Link("https://example.com"), Image("可點的圖"))));
        Assert.Equal("[\\[圖片: 可點的圖\\]](https://example.com)", _env.Md(builder)); // 連結文字裡的中括號要跳脫
    }

    [Fact]
    public void Tracked_change_note_is_reported_only_when_revisions_exist()
    {
        Assert.Contains(Read(P(Del("x") + Run("y"))).Notes, n => n.Contains("追蹤修訂", StringComparison.Ordinal));
        Assert.DoesNotContain(Read(Para("y")).Notes, n => n.Contains("追蹤修訂", StringComparison.Ordinal));
    }

    [Fact]
    public void Equation_note_is_reported() =>
        Assert.Contains(Read(P(Math("x"))).Notes, n => n.Contains("1 個數學公式", StringComparison.Ordinal));

    // ---- 頁首頁尾、註腳、註解 ----

    private static DocxBuilder Full()
    {
        var builder = new DocxBuilder();
        builder.Add(H(1, "正文"), P(Run("有註腳") + FootnoteRef(2) + Run("的句子")), Para("結尾"));
        builder.Header("機密文件").Footer("第 1 頁").Footnotes((2, "這是註腳內容")).Comments((0, "王小明", "請再確認"));
        return builder;
    }

    [Fact]
    public void Headers_footnotes_and_comments_are_not_output_by_default_but_the_notes_say_so()
    {
        var path = _env.Save(Full());
        var result = _env.Reader.Read(path);

        Assert.Equal("# 正文\n\n有註腳的句子\n\n結尾", result.Markdown);
        Assert.Contains(result.Notes, n => n.Contains("頁首", StringComparison.Ordinal) && n.Contains("IncludeHeadersFooters", StringComparison.Ordinal));
        Assert.Contains(result.Notes, n => n.Contains("1 則註腳", StringComparison.Ordinal));
        Assert.Contains(result.Notes, n => n.Contains("1 則註解", StringComparison.Ordinal));
    }

    [Fact]
    public void Headers_and_footers_are_appended_as_an_appendix_when_asked()
    {
        var md = _env.Md(Full(), new WordReadOptions(IncludeHeadersFooters: true));
        Assert.EndsWith("\n\n# 附錄：頁首頁尾\n\n- 頁首：機密文件\n- 頁尾：第 1 頁", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Footnotes_get_reference_markers_and_definitions_only_when_asked()
    {
        var md = _env.Md(Full(), new WordReadOptions(IncludeFootnotes: true));

        Assert.Contains("有註腳[^f2]的句子", md, StringComparison.Ordinal);
        Assert.EndsWith("# 附錄：註腳\n\n[^f2]: 這是註腳內容", md, StringComparison.Ordinal);
        Assert.DoesNotContain("[^f", _env.Md(Full()), StringComparison.Ordinal);
    }

    [Fact]
    public void Comments_are_listed_with_their_author_when_asked()
    {
        var md = _env.Md(Full(), new WordReadOptions(IncludeComments: true));
        Assert.EndsWith("# 附錄：註解\n\n- 王小明：請再確認", md, StringComparison.Ordinal);
    }

    [Fact]
    public void All_options_together_keep_a_stable_order()
    {
        var md = _env.Md(Full(), new WordReadOptions(true, true, true));

        var headers = md.IndexOf("附錄：頁首頁尾", StringComparison.Ordinal);
        var notes = md.IndexOf("附錄：註腳", StringComparison.Ordinal);
        var comments = md.IndexOf("附錄：註解", StringComparison.Ordinal);
        Assert.True(headers > 0 && headers < notes && notes < comments);
    }

    [Fact]
    public void Identical_headers_are_listed_once()
    {
        var builder = new DocxBuilder().Add(Para("內文"));
        builder.Header("同一個頁首", "header1").Header("同一個頁首", "header2").Header("不同的頁首", "header3");

        var md = _env.Md(builder, new WordReadOptions(IncludeHeadersFooters: true));

        Assert.Equal(1, md.Split("頁首：同一個頁首").Length - 1);
        Assert.Contains("頁首：不同的頁首", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Separator_footnotes_are_not_listed_and_empty_documents_have_no_appendix()
    {
        var builder = new DocxBuilder().Add(Para("內文")).Footnotes();
        var md = _env.Md(builder, new WordReadOptions(true, true, true));
        Assert.Equal("內文", md);
    }

    [Fact]
    public void Appendix_headings_do_not_leak_into_the_previous_section()
    {
        var path = _env.Save(Full());
        var outline = _env.Reader.GetOutline(path, new WordReadOptions(true, true, true));

        Assert.Equal(["正文", "附錄：頁首頁尾", "附錄：註腳", "附錄：註解"], outline.Sections.Select(s => s.Title));
    }
}
