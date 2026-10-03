using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocumentReader.Core.Tests;

/// <summary>
/// 另一種產生方式：用 Open XML SDK 的物件模型建立簡報（與手寫 XML 的 <see cref="PptxBuilder"/> 互相對照），
/// 避免兩邊用同樣的錯誤假設而互相掩護。
/// </summary>
public sealed class PptSdkGeneratedTests : IDisposable
{
    private readonly PptEnv _env = new();

    public void Dispose() => _env.Dispose();

    private static P.Shape TextShape(uint id, string text, long x, long y, P.PlaceholderValues? placeholder = null)
    {
        var nvPr = new P.ApplicationNonVisualDrawingProperties();
        if (placeholder is { } ph)
        {
            nvPr.AppendChild(new P.PlaceholderShape { Type = ph });
        }

        return new P.Shape(
            new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties { Id = id, Name = $"Shape {id}" }, new P.NonVisualShapeDrawingProperties(), nvPr),
            new P.ShapeProperties(new A.Transform2D(new A.Offset { X = x, Y = y }, new A.Extents { Cx = 1000, Cy = 1000 })),
            new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text(text)))));
    }

    private string Create(params (string Title, string Body, string? Notes)[] slides)
    {
        var path = _env.Path("sdk.pptx");
        using (var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var main = doc.AddPresentationPart();
            main.Presentation = new P.Presentation(new P.SlideIdList(), new P.SlideSize { Cx = 9144000, Cy = 5143500 });
            uint nextId = 256;
            foreach (var (title, body, notes) in slides)
            {
                var slidePart = main.AddNewPart<SlidePart>();
                slidePart.Slide = new P.Slide(new P.CommonSlideData(new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty }, new P.NonVisualGroupShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
                    new P.GroupShapeProperties(),
                    TextShape(2, title, 0, 0, P.PlaceholderValues.Title),
                    TextShape(3, body, 0, 1_000_000))));

                if (notes is not null)
                {
                    var notesPart = slidePart.AddNewPart<NotesSlidePart>();
                    notesPart.NotesSlide = new P.NotesSlide(new P.CommonSlideData(new P.ShapeTree(
                        new P.NonVisualGroupShapeProperties(new P.NonVisualDrawingProperties { Id = 1, Name = string.Empty }, new P.NonVisualGroupShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
                        new P.GroupShapeProperties(),
                        TextShape(2, notes, 0, 0, P.PlaceholderValues.Body))));
                }

                main.Presentation.SlideIdList!.AppendChild(new P.SlideId { Id = nextId++, RelationshipId = main.GetIdOfPart(slidePart) });
            }

            main.Presentation.Save();
        }

        return path;
    }

    [Fact]
    public void An_sdk_generated_deck_converts_titles_bodies_and_notes()
    {
        var path = Create(("開場", "歡迎大家", "先自我介紹"), ("結尾", "謝謝", null));
        var md = _env.Reader.Read(path).Markdown;

        Assert.Equal("## 投影片 1：開場\n\n歡迎大家\n\n> **備註：** 先自我介紹\n\n## 投影片 2：結尾\n\n謝謝", md);
    }

    [Fact]
    public void An_sdk_generated_deck_has_the_expected_outline()
    {
        var outline = _env.Reader.GetOutline(Create(("甲", "x", "n"), ("乙", "y", null)));

        Assert.Equal(["甲", "乙"], outline.Slides.Select(s => s.Title));
        Assert.Equal([true, false], outline.Slides.Select(s => s.HasNotes));
    }
}

/// <summary>
/// 用使用者提供的真實檔案（repo 外的 sample/ 資料夾，不進版控）驗證；檔案不存在時直接略過。
/// 斷言只放「換電腦、換版本都不會變」的事實：張數、標題、某些內容有出現。
/// </summary>
public sealed class RealSampleFileTests
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

    [Fact]
    public void The_real_pptx_converts_every_slide()
    {
        if (Find("sample.pptx") is not { } file)
        {
            return;
        }

        using var env = new PptEnv();
        var copy = env.Path("sample.pptx");
        File.Copy(file, copy);

        var outline = env.Reader.GetOutline(copy);
        Assert.Equal(40, outline.SlideCount);
        Assert.Equal("Kindergarten Day", outline.Slides[0].Title);
        Assert.Equal("This is a table", outline.Slides[15].Title);
        Assert.Equal(1, outline.Slides[15].Tables);

        var table = env.Reader.ReadSlides(copy, 16).Markdown;
        Assert.Contains("| Year | Venus | Mars |", table, StringComparison.Ordinal);
        Assert.Contains("$3,368.20", table, StringComparison.Ordinal);

        var all = env.Reader.Read(copy).Markdown;
        Assert.Contains("## 投影片 40", all, StringComparison.Ordinal);
        Assert.Contains("https://bit.ly/3A1uf1Q", all, StringComparison.Ordinal);
    }

    [Fact]
    public void The_real_docx_has_a_heading_tree_and_content()
    {
        if (Find("sample.docx") is not { } file)
        {
            return;
        }

        using var env = new WordEnv();
        var copy = env.Path("sample.docx");
        File.Copy(file, copy);

        var outline = env.Reader.GetOutline(copy);
        Assert.True(outline.HasHeadings);
        Assert.Contains(outline.Sections, s => s.Title.StartsWith("壹、", StringComparison.Ordinal));

        var section = env.Reader.ReadSection(copy, outline.Sections.First(s => s.Title.StartsWith("貳、", StringComparison.Ordinal)).SectionId);
        Assert.Contains("理、監事", section.Markdown, StringComparison.Ordinal);
    }
}
