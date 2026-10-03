using DocumentReader.Core.Models;
using OfficeTools.Common.Errors;
using static DocumentReader.Core.Tests.DocxBuilder;

namespace DocumentReader.Core.Tests;

public sealed class WordConversionTests : IDisposable
{
    private readonly WordEnv _env = new();

    public void Dispose() => _env.Dispose();

    private string Md(params string[] blocks) => _env.Md(new DocxBuilder().Add(blocks));

    // ---- 標題 ----

    [Fact]
    public void Headings_are_found_by_style_name_even_when_the_style_id_is_a_digit_like_in_chinese_word()
    {
        var md = Md(H(1, "第一章"), Para("內文"), H(2, "第一節"), H(3, "小節"));

        Assert.Equal("# 第一章\n\n內文\n\n## 第一節\n\n### 小節", md);
    }

    [Fact]
    public void Headings_also_work_with_english_style_ids()
    {
        var builder = new DocxBuilder { ChineseStyleIds = false }.Add(H(1, "One", chineseIds: false), H(2, "Two", chineseIds: false));
        Assert.Equal("# One\n\n## Two", _env.Md(builder));
    }

    [Fact]
    public void A_style_based_on_a_heading_style_is_a_heading()
    {
        Assert.Equal("## 自訂", Md(Para("自訂", "MyHeading")));
    }

    [Fact]
    public void A_style_named_in_chinese_is_a_heading()
    {
        Assert.Equal("# 標題樣式", Md(Para("標題樣式", "CustomChinese")));
    }

    [Fact]
    public void Direct_outline_level_on_a_normal_paragraph_makes_it_a_heading()
    {
        Assert.Equal("# 直接設定\n\n### 三級", Md(Para("直接設定", outline: 0), Para("三級", outline: 2)));
    }

    [Fact]
    public void Outline_level_nine_means_body_text_even_on_a_heading_style()
    {
        Assert.Equal("看起來像標題的內文", Md(Para("看起來像標題的內文", "BodyOutline9")));
        Assert.Equal("直接設定成內文", Md(Para("直接設定成內文", "1", outline: 9)));
    }

    [Fact]
    public void The_title_style_is_a_level_one_heading()
    {
        Assert.Equal("# 文件標題", Md(Para("文件標題", "Title")));
    }

    [Fact]
    public void Deep_heading_levels_use_at_most_six_hashes_in_markdown()
    {
        var md = Md(H(7, "七"), H(9, "九"));
        Assert.Equal("###### 七\n\n###### 九", md);
    }

    [Fact]
    public void Normal_paragraphs_are_not_headings() => Assert.Equal("普通段落", Md(Para("普通段落", "Normal")));

    [Fact]
    public void Empty_headings_and_empty_paragraphs_are_skipped() =>
        Assert.Equal("# 有字\n\n內文", Md(H(1, ""), Para("   "), H(1, "有字"), Para(""), Para("內文")));

    [Fact]
    public void Heading_text_drops_formatting_and_links()
    {
        var builder = new DocxBuilder();
        var link = builder.Link("https://example.com");
        builder.Add(P(Hyperlink(link, Run("連結標題", b: true)), style: "1"));

        Assert.Equal("# 連結標題", _env.Md(builder));
    }

    [Fact]
    public void Automatic_heading_numbers_are_added_as_a_prefix()
    {
        var md = Md(
            Para("總則", "1", numId: 4, ilvl: 0),
            Para("目的", "2", numId: 4, ilvl: 1),
            Para("範圍", "2", numId: 4, ilvl: 1),
            Para("細項", "3", numId: 4, ilvl: 2),
            Para("組織", "1", numId: 4, ilvl: 0),
            Para("架構", "2", numId: 4, ilvl: 1));

        Assert.Equal("# 1 總則\n\n## 1.1 目的\n\n## 1.2 範圍\n\n### 1.2.1 細項\n\n# 2 組織\n\n## 2.1 架構", md);
    }

    [Fact]
    public void Legal_numbering_uses_decimal_for_the_upper_levels()
    {
        var md = Md(Para("甲", "1", numId: 8, ilvl: 0), Para("乙", "2", numId: 8, ilvl: 1));
        Assert.Equal("# A 甲\n\n## 1.1 乙", md); // 第 0 層是字母 A；第 1 層標了 isLgl，所以 %1.%2 都用阿拉伯數字
    }

    // ---- 行內格式 ----

    [Fact]
    public void Bold_italic_and_strikethrough_runs()
    {
        var md = Md(P(Run("普通 ") + Run("粗體", b: true) + Run(" ") + Run("斜體", i: true) + Run(" ") + Run("粗斜", b: true, i: true) + Run(" ") + Run("刪除", strike: true)));
        Assert.Equal("普通 **粗體** *斜體* ***粗斜*** ~~刪除~~", md);
    }

    [Fact]
    public void Adjacent_runs_with_the_same_formatting_are_merged_into_one_span() =>
        Assert.Equal("**一二三** 四", Md(P(Run("一", b: true) + Run("二", b: true) + Run("三", b: true) + Run(" 四"))));

    [Fact]
    public void Whitespace_is_moved_outside_the_emphasis_markers_so_the_markdown_stays_valid() =>
        Assert.Equal("前 **粗體** 後", Md(P(Run("前") + Run(" 粗體 ", b: true) + Run("後"))));

    [Fact]
    public void Character_styles_such_as_strong_and_emphasis_apply()
    {
        Assert.Equal("**強** *調*", Md(P(Run("強", style: "Strong") + Run(" ") + Run("調", style: "Emphasis"))));
    }

    [Fact]
    public void Bold_switched_off_with_val_zero_is_not_bold() =>
        Assert.Equal("一般", Md(P(Run("一般", boldOff: true))));

    [Fact]
    public void Characters_that_markdown_would_misread_are_escaped()
    {
        Assert.Equal(@"a \* b \`c\` d\\e", Md(Para(@"a * b `c` d\e")));
    }

    [Theory]
    [InlineData("# 不是標題", @"\# 不是標題")]
    [InlineData("- 不是清單", @"\- 不是清單")]
    [InlineData("+ 也不是", @"\+ 也不是")]
    [InlineData("> 不是引用", @"\> 不是引用")]
    [InlineData("1. 不是編號", @"1\. 不是編號")]
    [InlineData("12) 也不是", @"12\) 也不是")]
    [InlineData("---", @"\---")]
    public void Paragraphs_that_start_like_block_syntax_are_escaped(string text, string expected) =>
        Assert.Equal(expected, Md(Para(text)));

    [Fact]
    public void Text_that_merely_contains_block_syntax_in_the_middle_is_left_alone() =>
        Assert.Equal("價格 # 3 - 5 > 2", Md(Para("價格 # 3 - 5 > 2")));

    [Fact]
    public void Underscores_and_brackets_are_not_escaped_to_keep_identifiers_and_citations_readable() =>
        Assert.Equal("my_var_name 與 [1] 與 a_b_c", Md(Para("my_var_name 與 [1] 與 a_b_c")));

    [Fact]
    public void Line_breaks_become_markdown_hard_breaks_and_page_breaks_and_tabs_are_handled() =>
        Assert.Equal("第一行  \n第二行 欄位 結尾", Md(P(Run("第一行") + Br() + Run("第二行") + Tab() + Run("欄位") + PageBreak() + Run(" 結尾"))));

    [Fact]
    public void A_line_after_a_hard_break_that_looks_like_a_list_is_escaped() =>
        Assert.Equal("說明  \n\\- 不是清單", Md(P(Run("說明") + Br() + Run("- 不是清單"))));

    [Fact]
    public void Hidden_text_is_skipped() =>
        Assert.Equal("看得見", Md(P(Run("看得見") + Run("藏起來", vanish: true))));

    [Fact]
    public void Chinese_english_emoji_and_symbols_survive() =>
        Assert.Equal("中文 English 😀 ©®™ € 日本語 한국어", Md(Para("中文 English 😀 ©®™ € 日本語 한국어")));

    [Fact]
    public void Leading_and_trailing_spaces_inside_a_paragraph_are_preserved_between_runs() =>
        Assert.Equal("a b  c", Md(P(Run("a ") + Run("b ") + Run(" c"))));

    // ---- 超連結 ----

    [Fact]
    public void External_hyperlinks_become_markdown_links()
    {
        var builder = new DocxBuilder();
        var link = builder.Link("https://example.com/a b(c)");
        builder.Add(P(Run("見 ") + Hyperlink(link, Run("官網")) + Run(" 了解")));

        Assert.Equal("見 [官網](https://example.com/a%20b%28c%29) 了解", _env.Md(builder));
    }

    [Fact]
    public void Links_keep_their_inner_formatting_and_internal_anchors_are_plain_text()
    {
        var builder = new DocxBuilder();
        var link = builder.Link("https://example.com");
        builder.Add(P(Hyperlink(link, Run("粗", b: true) + Run("連結", b: true)) + Run(" ") + Anchor("_Toc1", Run("內部"))));

        Assert.Equal("[**粗連結**](https://example.com) 內部", _env.Md(builder));
    }

    [Fact]
    public void A_mailto_link_works()
    {
        var builder = new DocxBuilder();
        builder.Add(P(Hyperlink(builder.Link("mailto:a@b.com"), Run("寫信"))));
        Assert.Equal("[寫信](mailto:a@b.com)", _env.Md(builder));
    }

    // ---- 清單 ----

    [Fact]
    public void Bullet_and_numbered_lists_with_nesting()
    {
        var md = Md(
            Para("蘋果", numId: 1),
            Para("子項", numId: 1, ilvl: 1),
            Para("香蕉", numId: 1),
            Para("第一", numId: 2),
            Para("第二", numId: 2),
            Para("第二之一", numId: 2, ilvl: 1),
            Para("第二之二", numId: 2, ilvl: 1),
            Para("第三", numId: 2));

        Assert.Equal("- 蘋果\n  - 子項\n- 香蕉\n\n1. 第一\n2. 第二\n  1. 第二之一\n  2. 第二之二\n3. 第三", md);
    }

    [Fact]
    public void A_new_list_instance_restarts_the_numbering_and_a_start_override_is_honored()
    {
        var md = Md(Para("甲", numId: 2), Para("乙", numId: 2), Para("說明"), Para("丙", numId: 3), Para("丁", numId: 5), Para("戊", numId: 5));
        Assert.Equal("1. 甲\n2. 乙\n\n說明\n\n1. 丙\n\n5. 丁\n6. 戊", md);
    }

    [Fact]
    public void A_higher_level_item_resets_the_deeper_counters()
    {
        var md = Md(Para("a", numId: 2), Para("a1", numId: 2, ilvl: 1), Para("a2", numId: 2, ilvl: 1), Para("b", numId: 2), Para("b1", numId: 2, ilvl: 1));
        Assert.Equal("1. a\n  1. a1\n  2. a2\n2. b\n  1. b1", md);
    }

    [Fact]
    public void Letter_and_roman_numbered_lists_use_plain_numbers_as_markdown_markers() =>
        Assert.Equal("1. 一\n2. 二\n\n1. 甲\n2. 乙", Md(Para("一", numId: 6), Para("二", numId: 6), Para("中間"), Para("甲", numId: 7), Para("乙", numId: 7)).Replace("\n\n中間\n\n", "\n\n"));

    [Fact]
    public void A_list_style_that_carries_numbering_makes_list_items() =>
        Assert.Equal("- 一\n- 二", Md(Para("一", "ListBullet"), Para("二", "ListBullet")));

    [Fact]
    public void A_paragraph_can_cancel_the_numbering_of_its_style() =>
        Assert.Equal("- 有符號\n\n沒有符號", Md(Para("有符號", "ListBullet"), Para("沒有符號", "ListBullet", numId: 0)));

    [Fact]
    public void A_list_followed_by_a_paragraph_is_separated_by_a_blank_line_and_items_by_a_single_newline() =>
        Assert.Equal("- 一\n- 二\n\n段落", Md(Para("一", numId: 1), Para("二", numId: 1), Para("段落")));

    [Fact]
    public void A_line_break_inside_a_list_item_stays_in_the_item() =>
        Assert.Equal("- 第一行\n  第二行\n- 下一項", Md(P(Run("第一行") + Br() + Run("第二行"), numId: 1), Para("下一項", numId: 1)));

    [Fact]
    public void Empty_numbered_paragraphs_do_not_consume_numbers() =>
        Assert.Equal("1. 甲\n2. 乙", Md(Para("甲", numId: 2), Para("", numId: 2), Para("乙", numId: 2)));

    [Fact]
    public void Headings_and_list_items_can_share_a_numbering_instance_without_breaking_either()
    {
        var md = Md(Para("章", "1", numId: 4, ilvl: 0), Para("項目", numId: 1), Para("章二", "1", numId: 4, ilvl: 0));
        Assert.Equal("# 1 章\n\n- 項目\n\n# 2 章二", md);
    }

    // ---- 其他行內內容 ----

    [Fact]
    public void Tracked_changes_output_the_accepted_version()
    {
        var md = Md(P(Run("保留 ") + Del("被刪除") + Ins(Run("新增的")) + Run(" 結尾")));
        Assert.Equal("保留 新增的 結尾", md);
    }

    [Fact]
    public void A_paragraph_that_was_entirely_deleted_disappears()
    {
        Assert.Equal("留下", Md(P(Del("整段都被刪")), Para("留下")));
    }

    [Fact]
    public void Moved_text_is_kept_at_its_new_location_only()
    {
        var moved = "<w:moveFrom w:id=\"3\" w:author=\"a\"><w:r><w:t>搬走</w:t></w:r></w:moveFrom><w:moveTo w:id=\"4\" w:author=\"a\"><w:r><w:t>搬來</w:t></w:r></w:moveTo>";
        Assert.Equal("搬來", Md(P(moved)));
    }

    [Fact]
    public void Content_controls_are_unwrapped_inline_and_at_block_level()
    {
        Assert.Equal("前 控制項 後\n\n區塊控制項", Md(P(Run("前 ") + InlineSdt(Run("控制項")) + Run(" 後")), BlockSdt(Para("區塊控制項"))));
    }

    [Fact]
    public void Field_results_are_kept_and_field_codes_are_not_shown()
    {
        Assert.Equal("第 3 頁，共 9 頁", Md(P(Run("第 ") + SimpleField("3") + Run(" 頁，共 ") + ComplexField("NUMPAGES", "9") + Run(" 頁"))));
    }

    [Fact]
    public void Equations_become_bracketed_plain_text() =>
        Assert.Equal("面積 [公式: πr2] 公尺", Md(P(Run("面積 ") + Math("πr2") + Run(" 公尺"))));

    [Fact]
    public void Table_of_contents_paragraphs_and_controls_are_skipped_with_a_note()
    {
        var builder = new DocxBuilder().Add(
            Para("目錄", "TOCHeading"),
            Para("第一章 ........ 1", "TOC1"),
            BlockSdt(Para("第二章 ........ 2"), gallery: "Table of Contents"),
            H(1, "第一章"),
            Para("內文"));

        var path = _env.Save(builder);
        var result = _env.Reader.Read(path);

        Assert.Equal("# 第一章\n\n內文", result.Markdown);
        Assert.Contains(result.Notes, n => n.Contains("目錄", StringComparison.Ordinal));
    }

    [Fact]
    public void An_empty_document_gives_empty_markdown()
    {
        var path = _env.Save(new DocxBuilder());
        var result = _env.Reader.Read(path);

        Assert.Equal(string.Empty, result.Markdown);
        Assert.Equal(0, result.TotalChars);
        Assert.False(result.Truncated);
        Assert.Empty(_env.Reader.GetOutline(path).Sections);
    }

    [Fact]
    public void Real_world_mixture_reads_in_document_order()
    {
        var builder = new DocxBuilder();
        var link = builder.Link("https://example.com");
        builder.Add(
            Para("季度報告", "Title"),
            H(1, "摘要"),
            P(Run("本季營收 ") + Run("成長 12%", b: true) + Run("，詳見 ") + Hyperlink(link, Run("附件")) + Run("。")),
            Para("重點一", numId: 1),
            Para("重點二", numId: 1),
            H(1, "數據"),
            SimpleTable(["月份", "營收"], ["1 月", "100"], ["2 月", "120"]));

        var expected = "# 季度報告\n\n# 摘要\n\n本季營收 **成長 12%**，詳見 [附件](https://example.com)。\n\n- 重點一\n- 重點二\n\n# 數據\n\n| 月份 | 營收 |\n| --- | --- |\n| 1 月 | 100 |\n| 2 月 | 120 |";
        Assert.Equal(expected, _env.Md(builder));
    }
}
