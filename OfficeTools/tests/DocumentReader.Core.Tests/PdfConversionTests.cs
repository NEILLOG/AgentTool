namespace DocumentReader.Core.Tests;

public sealed class PdfConversionTests : IDisposable
{
    private readonly PdfEnv _env = new();

    public void Dispose() => _env.Dispose();

    private string Md(PdfBuilder b) => _env.Md(b);

    // ---- 段落與順序 ----

    [Fact]
    public void Lines_that_fill_the_width_are_joined_into_one_paragraph()
    {
        var md = Md(new PdfBuilder().Page().Lines(50, 780, 10, "The quick brown fox jumps over the lazy dog and keeps", "running through the forest until the end of the day."));

        Assert.Equal("## 第 1 頁\n\nThe quick brown fox jumps over the lazy dog and keeps running through the forest until the end of the day.", md);
    }

    [Fact]
    public void A_line_that_stops_short_of_the_edge_keeps_its_line_break()
    {
        var md = Md(new PdfBuilder().Page().Lines(50, 780, 10, "Short one", "A much longer second line that reaches far to the right edge", "tail"));

        Assert.Contains("Short one  \nA much longer", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Paragraphs_separated_by_extra_space_become_separate_paragraphs()
    {
        var md = Md(new PdfBuilder().Page().Lines(50, 780, 10, "First paragraph text.").Lines(50, 740, 10, "Second paragraph text."));

        Assert.Equal("## 第 1 頁\n\nFirst paragraph text.\n\nSecond paragraph text.", md);
    }

    [Fact]
    public void Two_columns_are_read_one_column_at_a_time()
    {
        var left = new[] { "Left one", "Left two", "Left three" };
        var right = new[] { "Right one", "Right two", "Right three" };
        var md = Md(new PdfBuilder().Page().Lines(50, 780, 10, left).Lines(330, 780, 10, right));

        Assert.True(md.IndexOf("Left three", StringComparison.Ordinal) < md.IndexOf("Right one", StringComparison.Ordinal), md);
    }

    [Fact]
    public void A_title_spanning_both_columns_comes_first()
    {
        var md = Md(new PdfBuilder().Page().Text("A Wide Title Across The Whole Page Width And It Keeps Going Past The Gutter Between", 50, 800, 10).Lines(50, 760, 10, "Left one", "Left two").Lines(330, 760, 10, "Right one", "Right two"));

        var order = new[] { "A Wide Title", "Left one", "Left two", "Right one", "Right two" }.Select(s => md.IndexOf(s, StringComparison.Ordinal)).ToList();
        Assert.Equal(order.OrderBy(i => i), order);
    }

    [Fact]
    public void The_wider_gap_decides_between_splitting_columns_and_rows()
    {
        // 兩欄各有兩段、段落間距也貫穿整頁：欄間距（約 245pt）比段落間距（約 70pt）寬，應該一欄讀完再讀下一欄
        var md = Md(new PdfBuilder().Page().Text("Left A", 50, 780).Text("Left B", 50, 700).Text("Right A", 330, 780).Text("Right B", 330, 700));

        var order = new[] { "Left A", "Left B", "Right A", "Right B" }.Select(s => md.IndexOf(s, StringComparison.Ordinal)).ToList();
        Assert.Equal(order.OrderBy(i => i), order);
    }

    [Fact]
    public void A_line_ending_with_a_colon_is_a_label_and_keeps_its_break()
    {
        var md = Md(new PdfBuilder().Page().Lines(50, 780, 10, "Please fill in all of the following required fields below:", "name", "address"));

        Assert.Contains("below:  \nname", md, StringComparison.Ordinal);
    }

    [Fact]
    public void A_heading_directly_above_body_text_is_split_from_it()
    {
        var md = Md(new PdfBuilder().Page().Text("Big Title", 50, 800, 24).Text("body text one", 50, 786).Text("body text two here", 50, 774).Text("body three", 50, 762));

        Assert.Contains("### Big Title\n\nbody text one", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Text_drawn_out_of_order_is_still_read_top_to_bottom()
    {
        var md = Md(new PdfBuilder().Page().Text("bottom line here", 50, 300).Text("top line here", 50, 700));

        Assert.True(md.IndexOf("top line", StringComparison.Ordinal) < md.IndexOf("bottom line", StringComparison.Ordinal));
    }

    [Fact]
    public void Words_on_one_line_are_read_left_to_right()
    {
        var md = Md(new PdfBuilder().Page().Text("world", 120, 700).Text("Hello", 50, 700));

        Assert.Contains("Hello world", md, StringComparison.Ordinal);
    }

    // ---- 標題與清單 ----

    [Fact]
    public void Much_larger_text_becomes_a_heading()
    {
        var body = string.Concat(Enumerable.Repeat("body text body text body text ", 1));
        var md = Md(new PdfBuilder().Page().Text("Big Title", 50, 800, 24).Lines(50, 760, 10, body, body + "more", body + "again"));

        Assert.Contains("### Big Title", md, StringComparison.Ordinal);
        Assert.DoesNotContain("### body", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Text_of_body_size_is_never_a_heading()
    {
        var md = Md(new PdfBuilder().Page().Lines(50, 780, 10, "Just text", "More text", "Even more text"));

        Assert.DoesNotContain("#", md.Replace("## 第 1 頁", string.Empty, StringComparison.Ordinal), StringComparison.Ordinal);
    }

    [Fact]
    public void Bullet_characters_become_list_items()
    {
        var md = Md(new PdfBuilder().Page().Lines(50, 780, 10, "• first item", "• second item", "• third item"));

        Assert.Contains("- first item\n- second item\n- third item", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Numbered_lines_become_ordered_items_but_decimals_do_not()
    {
        var md = Md(new PdfBuilder().Page().Lines(50, 780, 10, "1. alpha", "2. beta").Lines(50, 700, 10, "3.5 million users"));

        Assert.Contains("1. alpha\n2. beta", md, StringComparison.Ordinal);
        Assert.Contains("3.5 million users", md, StringComparison.Ordinal);
        Assert.DoesNotContain("- 3.5", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_syntax_in_the_text_is_escaped()
    {
        var md = Md(new PdfBuilder().Page().Text("# not a heading", 50, 780).Text("a*b*c", 50, 700));

        Assert.Contains("\\# not a heading", md, StringComparison.Ordinal);
        Assert.Contains("a\\*b\\*c", md, StringComparison.Ordinal);
    }

    // ---- 連結 ----

    [Fact]
    public void Link_annotations_become_markdown_links()
    {
        var md = Md(new PdfBuilder().Page().Text("see the docs now", 50, 700).Link("https://example.com/docs", 50, 695, 100, 712));

        Assert.Contains("[see the docs](https://example.com/docs)", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Unsafe_link_schemes_are_dropped()
    {
        var md = Md(new PdfBuilder().Page().Text("click here", 50, 700).Link("javascript:alert(1)", 50, 695, 130, 712));

        Assert.DoesNotContain("javascript", md, StringComparison.Ordinal);
        Assert.Contains("click here", md, StringComparison.Ordinal);
    }

    // ---- 表格 ----

    private static PdfBuilder TableDoc() =>
        new PdfBuilder().Page().Grid([50, 150, 250, 350], [700, 680, 660, 640])
            .Text("Name", 55, 685).Text("Qty", 155, 685).Text("Price", 255, 685)
            .Text("Apple", 55, 665).Text("3", 155, 665).Text("1.50", 255, 665)
            .Text("Pear", 55, 645).Text("12", 155, 645).Text("2.00", 255, 645);

    [Fact]
    public void Ruled_grids_become_markdown_tables()
    {
        var md = Md(TableDoc());

        Assert.Contains("| Name | Qty | Price |\n| --- | --- | --- |\n| Apple | 3 | 1.50 |\n| Pear | 12 | 2.00 |", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Text_inside_a_table_is_not_repeated_outside_it()
    {
        var md = Md(TableDoc());

        Assert.Single(System.Text.RegularExpressions.Regex.Matches(md, "Apple"));
    }

    [Fact]
    public void Text_around_a_table_keeps_its_place()
    {
        var md = Md(TableDoc().Text("Before the table", 50, 760).Text("After the table", 50, 560));

        Assert.True(md.IndexOf("Before the table", StringComparison.Ordinal) < md.IndexOf("| Name", StringComparison.Ordinal));
        Assert.True(md.IndexOf("| Pear", StringComparison.Ordinal) < md.IndexOf("After the table", StringComparison.Ordinal));
    }

    [Fact]
    public void A_row_without_inner_vertical_lines_spans_all_columns()
    {
        var b = new PdfBuilder().Page();
        // 外框與水平線，欄線只畫在第 2、3 列（第 1 列是橫跨整個寬度的分組標題）
        b.Line(50, 700, 350, 700).Line(50, 680, 350, 680).Line(50, 660, 350, 660).Line(50, 640, 350, 640);
        b.Line(50, 700, 50, 640).Line(350, 700, 350, 640);
        b.Line(150, 680, 150, 640).Line(250, 680, 250, 640);
        b.Text("Group A", 55, 685).Text("x", 55, 665).Text("y", 155, 665).Text("z", 255, 665).Text("p", 55, 645).Text("q", 155, 645).Text("r", 255, 645);
        var md = Md(b);

        Assert.Contains("Group A", md, StringComparison.Ordinal);
        Assert.Contains("| x | y | z |", md, StringComparison.Ordinal);
        Assert.DoesNotContain("| Group A | Group A", md, StringComparison.Ordinal);
    }

    [Fact]
    public void A_single_box_is_not_a_table()
    {
        var b = new PdfBuilder().Page().Grid([50, 350], [700, 640]).Text("Just a framed note", 60, 665);
        var md = Md(b);

        Assert.DoesNotContain("|", md, StringComparison.Ordinal);
        Assert.Contains("Just a framed note", md, StringComparison.Ordinal);
    }

    [Fact]
    public void A_frame_with_one_column_is_not_a_table_even_with_several_rows()
    {
        var md = Md(new PdfBuilder().Page().Grid([50, 350], [700, 680, 660]).Text("first framed row", 55, 685).Text("second framed row", 55, 665));

        Assert.DoesNotContain("|", md, StringComparison.Ordinal);
        Assert.Contains("first framed row", md, StringComparison.Ordinal);
        Assert.Contains("second framed row", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Pipes_in_cells_are_escaped()
    {
        var b = new PdfBuilder().Page().Grid([50, 150, 250], [700, 680]).Text("a|b", 55, 685).Text("c", 155, 685);
        var md = Md(b);

        Assert.Contains("a\\|b", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Empty_tables_rows_are_dropped()
    {
        var b = new PdfBuilder().Page().Grid([50, 150, 250], [700, 680, 660, 640]).Text("h1", 55, 685).Text("h2", 155, 685).Text("v1", 55, 625 + 0);
        var md = Md(b);

        Assert.Contains("| h1 | h2 |", md, StringComparison.Ordinal);
        Assert.DoesNotContain("|   |   |\n", md.Replace("| --- | --- |", string.Empty, StringComparison.Ordinal), StringComparison.Ordinal);
    }

    // ---- 頁首頁尾 ----

    [Fact]
    public void Headers_and_footers_repeated_on_every_page_are_removed()
    {
        var b = new PdfBuilder();
        for (var i = 1; i <= 4; i++)
        {
            b.Page().Text("ACME Confidential", 50, 815).Text($"Body of page {i}", 50, 600).Text($"Page {i}", 280, 25);
        }

        var md = _env.Md(b);

        Assert.DoesNotContain("ACME", md, StringComparison.Ordinal);
        Assert.DoesNotContain("Page 3", md, StringComparison.Ordinal);
        Assert.Contains("Body of page 3", md, StringComparison.Ordinal);
        Assert.Contains(_env.Reader.GetOutline(_env.Path("doc.pdf")).Notes, n => n.Contains("頁首", StringComparison.Ordinal));
    }

    [Fact]
    public void Repeated_text_in_the_middle_of_pages_is_kept()
    {
        var b = new PdfBuilder();
        for (var i = 1; i <= 4; i++)
        {
            b.Page().Text("Total", 50, 400).Text($"Body {i}", 50, 600);
        }

        Assert.Contains("Total", _env.Md(b), StringComparison.Ordinal);
    }

    [Fact]
    public void Margin_text_that_appears_on_only_some_pages_is_kept()
    {
        var b = new PdfBuilder();
        for (var i = 1; i <= 4; i++)
        {
            b.Page().Text($"Body {i}", 50, 600);
            if (i == 1)
            {
                b.Text("Only On Page One", 50, 815);
            }
        }

        Assert.Contains("Only On Page One", _env.Md(b), StringComparison.Ordinal);
    }

    [Fact]
    public void Two_page_documents_never_lose_margin_text()
    {
        var b = new PdfBuilder().Page().Text("Header", 50, 815).Page().Text("Header", 50, 815);

        Assert.Contains("Header", _env.Md(b), StringComparison.Ordinal);
    }

    // ---- 其他 ----

    [Fact]
    public void An_empty_page_says_so()
    {
        var md = _env.Md(new PdfBuilder().Page());

        Assert.Equal("## 第 1 頁\n\n（此頁沒有文字）", md);
    }

    [Fact]
    public void Pages_are_separated_by_page_headings()
    {
        var md = _env.Md(new PdfBuilder().Page().Text("one", 50, 700).Page().Text("two", 50, 700));

        Assert.Equal("## 第 1 頁\n\none\n\n## 第 2 頁\n\ntwo", md);
    }
}
