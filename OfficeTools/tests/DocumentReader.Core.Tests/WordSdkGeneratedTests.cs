using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentReader.Core.Tests;

/// <summary>
/// 用 Open XML SDK 自己的物件模型產生 docx（序列化方式與手寫 XML 不同），交叉驗證讀取器對 SDK 產出的檔案也正常。
/// 這台機器上沒有真正的 Word 檔可測；真實文件請另外驗證（見 plan/05）。
/// </summary>
public sealed class WordSdkGeneratedTests : IDisposable
{
    private readonly WordEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static Paragraph Heading(string styleId, string text) =>
        new(new ParagraphProperties(new ParagraphStyleId { Val = styleId }), new Run(new Text(text)));

    private string Create(Action<MainDocumentPart, Body> build)
    {
        var path = _env.Path("sdk.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(
                new Style(new StyleName { Val = "heading 1" }, new StyleParagraphProperties(new OutlineLevel { Val = 0 })) { Type = StyleValues.Paragraph, StyleId = "Heading1" },
                new Style(new StyleName { Val = "heading 2" }, new BasedOn { Val = "Heading1" }, new StyleParagraphProperties(new OutlineLevel { Val = 1 })) { Type = StyleValues.Paragraph, StyleId = "Heading2" });
            build(main, main.Document.Body!);
            main.Document.Save();
        }

        return path;
    }

    [Fact]
    public void Headings_runs_and_hyperlinks_built_with_the_sdk_are_read_correctly()
    {
        var path = Create((main, body) =>
        {
            var relationship = main.AddHyperlinkRelationship(new Uri("https://example.com/path?q=1"), isExternal: true);
            body.Append(
                Heading("Heading1", "概述"),
                new Paragraph(
                    new Run(new Text("普通 ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new RunProperties(new Bold()), new Text("粗體")),
                    new Run(new Text(" 與 ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Hyperlink(new Run(new Text("連結"))) { Id = relationship.Id }),
                Heading("Heading2", "細節"),
                new Paragraph(new Run(new Text("內文"))));
        });

        var md = _env.Reader.Read(path).Markdown;

        Assert.Equal("# 概述\n\n普通 **粗體** 與 [連結](https://example.com/path?q=1)\n\n## 細節\n\n內文", md);
    }

    [Fact]
    public void Tables_built_with_the_sdk_including_merged_cells_are_read_correctly()
    {
        var path = Create((_, body) =>
        {
            TableCell Cell(string text, int? span = null, MergedCellValues? vMerge = null)
            {
                var properties = new TableCellProperties();
                if (span is not null)
                {
                    properties.Append(new GridSpan { Val = span });
                }

                if (vMerge is not null)
                {
                    properties.Append(vMerge == MergedCellValues.Continue ? new VerticalMerge() : new VerticalMerge { Val = vMerge });
                }

                return new TableCell(properties, new Paragraph(new Run(new Text(text))));
            }

            body.Append(new Table(
                new TableRow(Cell("標題", span: 2)),
                new TableRow(Cell("甲", vMerge: MergedCellValues.Restart), Cell("x")),
                new TableRow(Cell("", vMerge: MergedCellValues.Continue), Cell("y"))));
        });

        Assert.Equal("| 標題 | 標題 |\n| --- | --- |\n| 甲 | x |\n| 甲 | y |", _env.Reader.Read(path).Markdown);
    }

    [Fact]
    public void Footnotes_and_tracked_changes_built_with_the_sdk_are_read_correctly()
    {
        var path = Create((main, body) =>
        {
            var footnotes = main.AddNewPart<FootnotesPart>();
            footnotes.Footnotes = new Footnotes(
                new Footnote(new Paragraph(new Run(new SeparatorMark()))) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
                new Footnote(new Paragraph(new Run(new FootnoteReferenceMark()), new Run(new Text(" 註腳文字") { Space = SpaceProcessingModeValues.Preserve }))) { Id = 1 });

            body.Append(new Paragraph(
                new Run(new Text("內文") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FootnoteReference { Id = 1 }),
                new DeletedRun(new Run(new DeletedText("刪掉的"))) { Author = "a", Date = DateTime.UtcNow, Id = "1" },
                new InsertedRun(new Run(new Text("新增的"))) { Author = "a", Date = DateTime.UtcNow, Id = "2" }));
        });

        var plain = _env.Reader.Read(path);
        var withNotes = _env.Reader.Read(path, wordOptions: new Models.WordReadOptions(IncludeFootnotes: true));

        Assert.Equal("內文新增的", plain.Markdown);
        Assert.Contains("內文[^f1]新增的", withNotes.Markdown, StringComparison.Ordinal);
        Assert.EndsWith("[^f1]: 註腳文字", withNotes.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void A_document_with_no_styles_part_still_reads_but_headings_fall_back_to_direct_outline_levels()
    {
        var path = _env.Path("nostyles.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new ParagraphProperties(new OutlineLevel { Val = 0 }), new Run(new Text("直接大綱層級"))),
                new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }), new Run(new Text("樣式不存在")))));
        }

        Assert.Equal("# 直接大綱層級\n\n樣式不存在", _env.Reader.Read(path).Markdown);
    }

    [Fact]
    public void A_cyclic_style_chain_does_not_hang()
    {
        var path = Create((main, body) =>
        {
            main.StyleDefinitionsPart!.Styles!.Append(
                new Style(new StyleName { Val = "A" }, new BasedOn { Val = "B" }) { Type = StyleValues.Paragraph, StyleId = "A" },
                new Style(new StyleName { Val = "B" }, new BasedOn { Val = "A" }) { Type = StyleValues.Paragraph, StyleId = "B" });
            body.Append(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "A" }), new Run(new Text("循環樣式"))));
        });

        Assert.Equal("循環樣式", _env.Reader.Read(path).Markdown);
    }
}
