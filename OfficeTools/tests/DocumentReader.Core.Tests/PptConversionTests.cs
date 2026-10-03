using DocumentReader.Core.Models;
using static DocumentReader.Core.Tests.PptxBuilder;

namespace DocumentReader.Core.Tests;

public sealed class PptConversionTests : IDisposable
{
    private readonly PptEnv _env = new();

    public void Dispose() => _env.Dispose();

    private string Md(PptxBuilder builder, PptReadOptions? options = null) => _env.Md(builder, options);

    // ---- 標題與基本結構 ----

    [Fact]
    public void Each_slide_is_a_heading_with_its_title()
    {
        var md = Md(new PptxBuilder()
            .Slide(Title("第一張") + Sp(3, Para("內文一"), y: 1000000))
            .Slide(Title("第二張") + Sp(3, Para("內文二"), y: 1000000)));

        Assert.Equal("## 投影片 1：第一張\n\n內文一\n\n## 投影片 2：第二張\n\n內文二", md);
    }

    [Fact]
    public void A_slide_without_a_title_is_numbered_only()
    {
        var md = Md(new PptxBuilder().Slide(Sp(2, Para("只有內文"), x: 0)));

        Assert.StartsWith("## 投影片 1\n\n只有內文", md, StringComparison.Ordinal);
    }

    [Fact]
    public void The_centered_title_placeholder_is_also_a_title()
    {
        var md = Md(new PptxBuilder().Slide(Sp(2, Para("封面標題"), ph: "type=\"ctrTitle\"") + Sp(3, Para("副標題"), ph: "type=\"subTitle\" idx=\"1\"", y: 900000)));

        Assert.Equal("## 投影片 1：封面標題\n\n副標題", md);
    }

    [Fact]
    public void A_multi_line_title_is_joined_into_one_heading()
    {
        var md = Md(new PptxBuilder().Slide(Sp(2, Para("上") + Para("下"), ph: "type=\"title\"")));

        Assert.StartsWith("## 投影片 1：上 下", md, StringComparison.Ordinal);
    }

    [Fact]
    public void A_second_title_placeholder_is_kept_as_text()
    {
        var md = Md(new PptxBuilder().Slide(Title("主標題") + Sp(3, Para("另一個標題"), ph: "type=\"title\"", y: 900000)));

        Assert.Equal("## 投影片 1：主標題\n\n另一個標題", md);
    }

    [Fact]
    public void Footer_date_and_slide_number_placeholders_are_skipped()
    {
        var md = Md(new PptxBuilder().Slide(
            Title("T")
            + Sp(3, Para("頁尾文字"), ph: "type=\"ftr\" idx=\"11\"")
            + Sp(4, Para("2026/10/3"), ph: "type=\"dt\" idx=\"10\"")
            + Sp(5, Para("7"), ph: "type=\"sldNum\" idx=\"12\"")
            + Sp(6, Para("內文"), y: 900000)));

        Assert.Equal("## 投影片 1：T\n\n內文", md);
    }

    [Fact]
    public void Shapes_without_text_are_ignored()
    {
        var md = Md(new PptxBuilder().Slide(Title("T") + Sp(3, "<a:p/>", y: 900000) + Sp(4, string.Empty, y: 1500000)));

        Assert.Equal("## 投影片 1：T", md);
    }

    [Fact]
    public void An_empty_presentation_has_no_text()
    {
        Assert.Equal(string.Empty, Md(new PptxBuilder()));
    }

    [Fact]
    public void Slide_order_follows_the_presentation_not_the_file_names()
    {
        var md = Md(new PptxBuilder { SlideOrder = [2, 1] }.Slide(Title("檔案一")).Slide(Title("檔案二")));

        Assert.Equal("## 投影片 1：檔案二\n\n## 投影片 2：檔案一", md);
    }

    [Fact]
    public void A_hidden_slide_is_marked()
    {
        var md = Md(new PptxBuilder().Slide(Title("藏起來"), hidden: true).Slide(Title("看得見")));

        Assert.Contains("## 投影片 1：藏起來（隱藏）", md, StringComparison.Ordinal);
        Assert.Contains("## 投影片 2：看得見\n", md + "\n", StringComparison.Ordinal);
        Assert.DoesNotContain("看得見（隱藏）", md, StringComparison.Ordinal);
        Assert.Contains(_env.Reader.GetOutline(_env.Path("deck.pptx")).Notes, n => n.Contains("1 張隱藏", StringComparison.Ordinal));
    }

    // ---- 順序 ----

    [Fact]
    public void Shapes_are_ordered_top_to_bottom()
    {
        var md = Md(new PptxBuilder().Slide(
            Sp(2, Para("下面"), y: 3000000)
            + Sp(3, Para("上面"), y: 1000000)
            + Sp(4, Para("中間"), y: 2000000)));

        Assert.Equal("## 投影片 1\n\n上面\n\n中間\n\n下面", md);
    }

    [Fact]
    public void Shapes_on_the_same_row_are_ordered_left_to_right()
    {
        // 右邊那個比左邊高一點點（在同一列的容許範圍內），仍應先讀左邊
        var md = Md(new PptxBuilder().Slide(
            Sp(2, Para("右"), x: 5000000, y: 1000000)
            + Sp(3, Para("左"), x: 500000, y: 1050000)));

        Assert.Equal("## 投影片 1\n\n左\n\n右", md);
    }

    [Fact]
    public void Shapes_far_apart_vertically_are_not_the_same_row()
    {
        var md = Md(new PptxBuilder().Slide(
            Sp(2, Para("右上"), x: 5000000, y: 500000)
            + Sp(3, Para("左下"), x: 0, y: 2500000)));

        Assert.Equal("## 投影片 1\n\n右上\n\n左下", md);
    }

    [Fact]
    public void Placeholders_come_before_other_shapes_in_document_order()
    {
        var md = Md(new PptxBuilder().Slide(
            Sp(2, Para("文字方塊在上"), y: 100000, txBox: true)
            + Sp(3, Para("本文一"), ph: "idx=\"1\" type=\"body\"", y: 3000000)
            + Sp(4, Para("本文二"), ph: "idx=\"2\" type=\"body\"", y: 1000000)));

        Assert.Equal("## 投影片 1\n\n- 本文一\n\n- 本文二\n\n文字方塊在上", md);
    }

    [Fact]
    public void The_title_is_first_even_when_it_is_lowest_in_the_document()
    {
        var md = Md(new PptxBuilder().Slide(Sp(2, Para("內文"), y: 100000) + Title("標題", id: 3)));

        Assert.Equal("## 投影片 1：標題\n\n內文", md);
    }

    [Fact]
    public void A_group_is_expanded_and_placed_by_the_group_position()
    {
        var md = Md(new PptxBuilder().Slide(
            Sp(2, Para("群組之後"), y: 2000000)
            + Grp(3, Sp(4, Para("群組內下"), y: 500) + Sp(5, Para("群組內上"), y: 100), y: 1000000)
            + Sp(6, Para("群組之前"), y: 100000)));

        Assert.Equal("## 投影片 1\n\n群組之前\n\n群組內上\n\n群組內下\n\n群組之後", md);
    }

    [Fact]
    public void Nested_groups_are_expanded_recursively()
    {
        var md = Md(new PptxBuilder().Slide(Grp(2, Grp(3, Sp(4, Para("最深一層"))))));

        Assert.Equal("## 投影片 1\n\n最深一層", md);
    }

    [Fact]
    public void Hidden_shapes_and_hidden_groups_are_skipped()
    {
        var hiddenGroup = Grp(10, Sp(11, Para("群組裡的字"))).Replace("name=\"Group 10\"", "name=\"Group 10\" hidden=\"1\"", StringComparison.Ordinal);
        var md = Md(new PptxBuilder().Slide(
            Sp(2, Para("隱藏的形狀"), extraCNvPr: "hidden=\"1\"")
            + hiddenGroup
            + Sp(3, Para("看得見"), y: 1000000)));

        Assert.Equal("## 投影片 1\n\n看得見", md);
    }

    [Fact]
    public void Alternate_content_uses_the_choice_only()
    {
        var alternate = "<mc:AlternateContent><mc:Choice Requires=\"x\">" + Sp(2, Para("新版本")) + "</mc:Choice><mc:Fallback>" + Sp(3, Para("舊版本")) + "</mc:Fallback></mc:AlternateContent>";
        var md = Md(new PptxBuilder().Slide(alternate));

        Assert.Equal("## 投影片 1\n\n新版本", md);
    }

    // ---- 清單 ----

    [Fact]
    public void Explicit_bullets_become_list_items_with_nesting()
    {
        var md = Md(new PptxBuilder().Slide(Sp(2, Bullet("一") + Bullet("一之一", 1) + Bullet("一之一之一", 2) + Bullet("二"))));

        Assert.Equal("## 投影片 1\n\n- 一\n  - 一之一\n    - 一之一之一\n- 二", md);
    }

    [Fact]
    public void Auto_numbered_paragraphs_are_counted_per_level()
    {
        var md = Md(new PptxBuilder().Slide(Sp(2, Numbered("甲") + Numbered("乙") + Numbered("乙之一", 1) + Numbered("乙之二", 1) + Numbered("丙"))));

        Assert.Equal("## 投影片 1\n\n1. 甲\n2. 乙\n  1. 乙之一\n  2. 乙之二\n3. 丙", md);
    }

    [Fact]
    public void A_plain_paragraph_restarts_the_numbering()
    {
        var md = Md(new PptxBuilder().Slide(Sp(2, Numbered("甲") + Numbered("乙") + Para("插入") + Numbered("丙"))));

        Assert.Equal("## 投影片 1\n\n1. 甲\n2. 乙\n\n插入\n\n1. 丙", md);
    }

    [Fact]
    public void Body_placeholders_have_bullets_unless_bullets_are_turned_off()
    {
        var body = Para("預設有符號") + ParaPr("<a:pPr><a:buNone/></a:pPr>", "明確關掉");
        var md = Md(new PptxBuilder().Slide(Sp(2, body, ph: "idx=\"1\"")));

        Assert.Equal("## 投影片 1\n\n- 預設有符號\n\n明確關掉", md);
    }

    [Fact]
    public void Text_boxes_have_no_bullets_by_default()
    {
        var md = Md(new PptxBuilder().Slide(Sp(2, Para("一") + Para("二"), txBox: true)));

        Assert.Equal("## 投影片 1\n\n一\n\n二", md);
    }

    [Fact]
    public void Lists_in_different_shapes_are_separated()
    {
        var md = Md(new PptxBuilder().Slide(Sp(2, Bullet("甲")) + Sp(3, Bullet("乙"), y: 2000000)));

        Assert.Equal("## 投影片 1\n\n- 甲\n\n- 乙", md);
    }

    [Fact]
    public void A_line_break_inside_a_list_item_stays_in_the_item()
    {
        var item = "<a:p><a:pPr><a:buChar char=\"•\"/></a:pPr>" + Run("上") + "<a:br/>" + Run("下") + "</a:p>";
        var md = Md(new PptxBuilder().Slide(Sp(2, item)));

        Assert.Equal("## 投影片 1\n\n- 上\n  下", md);
    }

    // ---- 行內 ----

    [Fact]
    public void Bold_italic_and_strike_are_kept()
    {
        var md = Md(new PptxBuilder().Slide(Sp(2, Paragraph(Run("粗", "b=\"1\""), Run("斜", "i=\"1\""), Run("刪", "strike=\"sngStrike\""), Run("平")))));

        Assert.Contains("**粗***斜*~~刪~~平", md, StringComparison.Ordinal);
    }

    [Fact]
    public void External_hyperlinks_become_markdown_links()
    {
        var link = Run("點我", rPrChildren: "<a:hlinkClick r:id=\"rId9\"/>");
        var md = Md(new PptxBuilder().Slide(Sp(2, Paragraph(Run("看 "), link)), links: [("rId9", "https://example.com/a b")]));

        Assert.Contains("看 [點我](https://example.com/a%20b)", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Links_that_are_not_web_addresses_are_dropped_but_the_text_stays()
    {
        var link = Run("跳頁", rPrChildren: "<a:hlinkClick r:id=\"rId9\" action=\"ppaction://hlinksldjump\"/>");
        var md = Md(new PptxBuilder().Slide(Sp(2, Paragraph(link)), links: [("rId9", "javascript:alert(1)")]));

        Assert.Equal("## 投影片 1\n\n跳頁", md);
    }

    [Fact]
    public void Line_breaks_become_hard_breaks()
    {
        var md = Md(new PptxBuilder().Slide(Sp(2, "<a:p>" + Run("上") + "<a:br/>" + Run("下") + "</a:p>")));

        Assert.Equal("## 投影片 1\n\n上  \n下", md);
    }

    [Fact]
    public void Fields_keep_their_cached_text()
    {
        var md = Md(new PptxBuilder().Slide(Sp(2, "<a:p>" + Run("第 ") + "<a:fld id=\"{1}\" type=\"slidenum\"><a:rPr lang=\"en\"/><a:t>3</a:t></a:fld>" + Run(" 頁") + "</a:p>")));

        Assert.Equal("## 投影片 1\n\n第 3 頁", md);
    }

    [Fact]
    public void Vertical_tabs_become_line_breaks_and_control_characters_are_removed()
    {
        var md = Md(new PptxBuilder().Slide(Sp(2, "<a:p><a:r><a:t>上_x000B_下_x0001_\t尾</a:t></a:r></a:p>")));

        Assert.Equal("## 投影片 1\n\n上  \n下 尾", md);
    }

    [Fact]
    public void Markdown_syntax_in_the_text_is_escaped()
    {
        var md = Md(new PptxBuilder().Slide(Sp(2, Para("# 不是標題") + Para("1. 不是清單") + Para("a*b*c") + Para("- 不是項目"), txBox: true)));

        Assert.Equal("## 投影片 1\n\n\\# 不是標題\n\n1\\. 不是清單\n\na\\*b\\*c\n\n\\- 不是項目", md);
    }

    [Fact]
    public void A_heading_line_in_the_title_cannot_break_the_structure()
    {
        var md = Md(new PptxBuilder().Slide(Title("# 假標題")));

        Assert.StartsWith("## 投影片 1：# 假標題", md, StringComparison.Ordinal);
        Assert.Equal(1, md.Split('\n').Count(l => l.StartsWith('#')));
    }

    // ---- 表格 ----

    [Fact]
    public void Tables_become_markdown_tables()
    {
        var table = Table(2, 0, 1000000, [Cell("姓名"), Cell("分數")], [Cell("小明"), Cell("90")]);
        var md = Md(new PptxBuilder().Slide(table));

        Assert.Equal("## 投影片 1\n\n| 姓名 | 分數 |\n| --- | --- |\n| 小明 | 90 |", md);
    }

    [Fact]
    public void Merged_cells_repeat_their_value()
    {
        var table = Table(
            2,
            0,
            0,
            [Cell("合併標題", "gridSpan=\"2\""), Cell(string.Empty, "hMerge=\"1\""), Cell("C")],
            [Cell("縱向", "rowSpan=\"2\""), Cell("a"), Cell("b")],
            [Cell(string.Empty, "vMerge=\"1\""), Cell("c"), Cell("d")]);
        var md = Md(new PptxBuilder().Slide(table));

        Assert.Equal("## 投影片 1\n\n| 合併標題 | 合併標題 | C |\n| --- | --- | --- |\n| 縱向 | a | b |\n| 縱向 | c | d |", md);
    }

    [Fact]
    public void Pipes_and_line_breaks_in_cells_are_escaped()
    {
        var cell = "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>" + Para("a|b") + Para("第二行") + "</a:txBody><a:tcPr/></a:tc>";
        var md = Md(new PptxBuilder().Slide(Table(2, 0, 0, [cell])));

        Assert.Contains("| a\\|b<br>第二行 |", md, StringComparison.Ordinal);
    }

    [Fact]
    public void A_table_with_only_empty_cells_is_dropped()
    {
        var md = Md(new PptxBuilder().Slide(Title("T") + Table(3, 0, 1000000, [Cell(string.Empty), Cell(string.Empty)])));

        Assert.Equal("## 投影片 1：T", md);
    }

    [Fact]
    public void Tables_are_ordered_with_other_shapes_by_position()
    {
        var md = Md(new PptxBuilder().Slide(
            Sp(2, Para("表格之後"), y: 3000000)
            + Table(3, 0, 1000000, [Cell("x")])
            + Sp(4, Para("表格之前"), y: 100000)));

        Assert.Equal("## 投影片 1\n\n表格之前\n\n| x |\n| --- |\n\n表格之後", md);
    }

    // ---- 圖片與其他物件 ----

    [Fact]
    public void Pictures_show_their_alt_text()
    {
        var md = Md(new PptxBuilder().Slide(Pic(2, "營收成長圖 [2025]")));

        Assert.Equal("## 投影片 1\n\n[圖片: 營收成長圖 [2025]]", md);
    }

    [Fact]
    public void Pictures_without_alt_text_are_collapsed_when_consecutive()
    {
        var md = Md(new PptxBuilder().Slide(Pic(2, null, x: 0) + Pic(3, string.Empty, x: 100) + Pic(4, null, x: 200)));

        Assert.Equal("## 投影片 1\n\n[圖片 ×3]", md);
    }

    [Fact]
    public void A_single_picture_without_alt_text_is_not_multiplied()
    {
        Assert.Equal("## 投影片 1\n\n[圖片]", Md(new PptxBuilder().Slide(Pic(2, null))));
    }

    [Fact]
    public void Pictures_with_alt_text_are_not_collapsed()
    {
        var md = Md(new PptxBuilder().Slide(Pic(2, "甲", x: 0) + Pic(3, "甲", x: 100)));

        Assert.Equal("## 投影片 1\n\n[圖片: 甲]\n\n[圖片: 甲]", md);
    }

    [Fact]
    public void Hidden_pictures_are_skipped()
    {
        Assert.Equal("## 投影片 1", Md(new PptxBuilder().Slide(Pic(2, "不顯示", hidden: true))));
    }

    [Fact]
    public void Smartart_and_embedded_objects_are_placeholders_with_notes()
    {
        var smart = Frame(2, "<dgm:relIds xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>", "http://schemas.openxmlformats.org/drawingml/2006/diagram", y: 100000);
        var ole = Frame(3, "<p:oleObj xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>", "http://schemas.openxmlformats.org/presentationml/2006/ole", y: 2000000);
        var md = Md(new PptxBuilder().Slide(smart + ole));

        Assert.Equal("## 投影片 1\n\n[SmartArt]\n\n[嵌入物件]", md);
        var notes = _env.Reader.GetOutline(_env.Path("deck.pptx")).Notes;
        Assert.Contains(notes, n => n.Contains("SmartArt", StringComparison.Ordinal));
        Assert.Contains(notes, n => n.Contains("嵌入物件", StringComparison.Ordinal));
    }

    // ---- 備註 ----

    [Fact]
    public void Speaker_notes_are_appended_as_a_quote()
    {
        var md = Md(new PptxBuilder().Slide(Title("T") + Sp(3, Para("內文"), y: 900000), notes: "講者要說的話\n第二段"));

        Assert.Equal("## 投影片 1：T\n\n內文\n\n> **備註：** 講者要說的話\n> 第二段", md);
    }

    [Fact]
    public void Notes_can_be_turned_off()
    {
        var builder = new PptxBuilder().Slide(Title("T"), notes: "不要看");
        var md = Md(builder, new PptReadOptions(IncludeNotes: false));

        Assert.Equal("## 投影片 1：T", md);
    }

    [Fact]
    public void The_notes_option_is_part_of_the_cache_key()
    {
        var path = _env.Save(new PptxBuilder().Slide(Title("T"), notes: "備註內容"));

        Assert.Contains("備註內容", _env.Reader.Read(path).Markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("備註內容", _env.Reader.Read(path, pptOptions: new PptReadOptions(false)).Markdown, StringComparison.Ordinal);
        Assert.Equal(2, _env.Cache.Count);
    }

    [Fact]
    public void The_slide_number_placeholder_in_notes_is_not_included()
    {
        var md = Md(new PptxBuilder().Slide(Title("T"), notes: "只有這個"));

        Assert.DoesNotContain("7", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Notes_escape_markdown_syntax()
    {
        var md = Md(new PptxBuilder().Slide(Title("T"), notes: "# 不是標題"));

        Assert.Contains("> **備註：** \\# 不是標題", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Empty_notes_are_not_shown_and_not_flagged()
    {
        var path = _env.Save(new PptxBuilder().Slide(Title("T"), notes: string.Empty));

        Assert.Equal("## 投影片 1：T", _env.Reader.Read(path).Markdown);
        Assert.False(_env.Reader.GetOutline(path).Slides[0].HasNotes);
    }
}
