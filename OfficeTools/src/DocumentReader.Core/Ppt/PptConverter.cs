using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentReader.Core.Markdown;
using DocumentReader.Core.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocumentReader.Core.Ppt;

internal enum OutKind
{
    Paragraph,
    ListItem,
    Table,
}

internal sealed record Out(OutKind Kind, string Text, int ListKey = 0);

internal sealed record SlideInfo(int Number, string Title, bool Hidden, bool HasNotes, int Pictures, int Tables, int Charts, int Start, int End);

internal sealed record ParsedPresentation(string Text, IReadOnlyList<SlideInfo> Slides, IReadOnlyList<string> Notes);

/// <summary>
/// 把簡報轉成 Markdown：每張投影片一個 "## 投影片 N：標題" 區塊。
/// 文字順序是標題、其他預留位置（依文件順序），再來是其他形狀（由上而下、由左而右；群組當成一個單位）。
/// </summary>
internal sealed class PptConverter(PresentationPart presentation, PptReadOptions options)
{
    private const int HeadingLevel = 2;
    private const int MaxBulletLevel = 8;

    private sealed record Item(long X, long Y, bool Placeholder, List<Out> Outs);

    private readonly MarkdownDocumentBuilder _builder = new();
    private readonly List<(int Number, string Title, bool Hidden, bool HasNotes, int Pictures, int Tables, int Charts)> _slides = [];
    private readonly long _rowTolerance = Math.Max(1, (presentation.Presentation?.SlideSize?.Cy?.Value ?? 5_143_500) / 30);

    // 目前這張投影片的狀態
    private string _title = string.Empty;
    private bool _hasTitle;
    private int _pictures;
    private int _tables;
    private int _charts;
    private int _listKey;
    private readonly int[] _counters = new int[MaxBulletLevel + 1];

    // 整份簡報
    private int _totalPictures;
    private int _totalCharts;
    private int _diagrams;
    private int _embedded;
    private int _hiddenSlides;
    private int _missingSlides;

    public ParsedPresentation Convert()
    {
        var number = 0;
        var ids = presentation.Presentation?.SlideIdList?.Elements<P.SlideId>() ?? Enumerable.Empty<P.SlideId>();
        foreach (var id in ids)
        {
            if (id.RelationshipId?.Value is not { } rid || PartOrNull(presentation, rid) is not SlidePart slide)
            {
                _missingSlides++;
                continue;
            }

            number++;
            ConvertSlide(slide, number);
        }

        var notes = new List<string>();
        AddConversionNotes(notes);
        var doc = _builder.Build(notes);

        var slides = new List<SlideInfo>();
        for (var i = 0; i < _slides.Count; i++)
        {
            var s = _slides[i];
            var section = doc.Sections[i];
            slides.Add(new SlideInfo(s.Number, s.Title, s.Hidden, s.HasNotes, s.Pictures, s.Tables, s.Charts, section.Start, section.OwnEnd));
        }

        return new ParsedPresentation(doc.Text, slides, notes);
    }

    private void ConvertSlide(SlidePart part, int number)
    {
        _title = string.Empty;
        _hasTitle = false;
        _pictures = _tables = _charts = 0;

        var tree = part.Slide?.CommonSlideData?.ShapeTree;
        var outs = tree is null ? [] : Flatten(Order(Collect(tree, part)));

        var hidden = part.Slide?.Show?.Value == false;
        var noteText = options.IncludeNotes ? NotesText(part) : string.Empty;
        var heading = $"投影片 {number.ToString(CultureInfo.InvariantCulture)}" + (_title.Length > 0 ? "：" + _title : string.Empty) + (hidden ? "（隱藏）" : string.Empty);
        _builder.AddHeading(HeadingLevel, heading);
        Emit(outs);

        if (noteText.Length > 0)
        {
            _builder.AddParagraph("> **備註：** " + noteText.Replace("\n", "\n> ", StringComparison.Ordinal));
        }

        _hiddenSlides += hidden ? 1 : 0;
        _totalPictures += _pictures;
        _totalCharts += _charts;
        _slides.Add((number, _title, hidden, noteText.Length > 0, _pictures, _tables, _charts));
    }

    // ---- 收集形狀 ----

    private List<Item> Collect(OpenXmlElement container, OpenXmlPartContainer part)
    {
        var items = new List<Item>();
        foreach (var child in container.ChildElements)
        {
            var element = child.LocalName == "AlternateContent" ? Chosen(child) : child;
            if (ReferenceEquals(element, child))
            {
                AddItem(items, element, part);
            }
            else
            {
                items.AddRange(Collect(element, part));
            }
        }

        return items;
    }

    private void AddItem(List<Item> items, OpenXmlElement element, OpenXmlPartContainer part)
    {
        switch (element)
        {
            case P.Shape shape when !IsHidden(shape.NonVisualShapeProperties?.NonVisualDrawingProperties):
                AddShape(items, shape, part);
                break;
            case P.Picture picture when !IsHidden(picture.NonVisualPictureProperties?.NonVisualDrawingProperties):
                var properties = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
                var alt = Clean(properties?.Description?.Value is { Length: > 0 } d ? d : properties?.Title?.Value ?? string.Empty);
                _pictures++;
                var offset = picture.ShapeProperties?.Transform2D?.Offset;
                items.Add(new Item(offset?.X?.Value ?? 0, offset?.Y?.Value ?? 0, false, [new Out(OutKind.Paragraph, alt.Length > 0 ? $"[圖片: {MarkdownWriter.Escape(alt)}]" : "[圖片]")]));
                break;
            case P.GraphicFrame frame when !IsHidden(frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties):
                AddGraphicFrame(items, frame, part);
                break;
            case P.GroupShape group when !IsHidden(group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties):
                var outs = Flatten(Order(Collect(group, part)));
                if (outs.Count > 0)
                {
                    var groupOffset = group.GroupShapeProperties?.TransformGroup?.Offset;
                    items.Add(new Item(groupOffset?.X?.Value ?? 0, groupOffset?.Y?.Value ?? 0, false, outs));
                }

                break;
        }
    }

    private void AddShape(List<Item> items, P.Shape shape, OpenXmlPartContainer part)
    {
        var placeholder = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
        var type = placeholder?.Type?.Value;
        if (type is { } t && (t == P.PlaceholderValues.DateAndTime || t == P.PlaceholderValues.Footer || t == P.PlaceholderValues.SlideNumber || t == P.PlaceholderValues.Header))
        {
            return;
        }

        var paragraphs = shape.TextBody?.Elements<A.Paragraph>().ToList() ?? [];
        var isTitle = type is { } tt && (tt == P.PlaceholderValues.Title || tt == P.PlaceholderValues.CenteredTitle);
        if (isTitle && !_hasTitle)
        {
            _title = Clean(string.Join(" ", paragraphs.Select(p => PlainText(p, part)).Where(x => x.Length > 0)));
            _hasTitle = true;
            return;
        }

        // 內容 / 本文預留位置預設有項目符號（取決於母片，這裡不解析母片，只在段落沒有明確寫 buNone 時採用）
        var bulletByDefault = placeholder is not null && (type is null || type.Value == P.PlaceholderValues.Body || type.Value == P.PlaceholderValues.Object);
        var outs = Paragraphs(paragraphs, part, bulletByDefault);
        if (outs.Count == 0)
        {
            return;
        }

        var offset = shape.ShapeProperties?.Transform2D?.Offset;
        items.Add(new Item(offset?.X?.Value ?? 0, offset?.Y?.Value ?? 0, placeholder is not null, outs));
    }

    private void AddGraphicFrame(List<Item> items, P.GraphicFrame frame, OpenXmlPartContainer part)
    {
        var offset = frame.Transform?.Offset;
        long x = offset?.X?.Value ?? 0, y = offset?.Y?.Value ?? 0;
        var data = frame.Graphic?.GraphicData;
        if (data is null)
        {
            return;
        }

        foreach (var child in data.ChildElements)
        {
            switch (child.LocalName)
            {
                case "tbl" when child is A.Table table:
                    var markdown = TableMarkdown(table, part);
                    if (markdown is not null)
                    {
                        _tables++;
                        items.Add(new Item(x, y, false, [new Out(OutKind.Table, markdown)]));
                    }

                    break;
                case "chart":
                    _charts++;
                    var id = child.GetAttributes().Where(a => a.LocalName == "id").Select(a => a.Value).FirstOrDefault();
                    var outs = new List<Out>();
                    if (id is not null && PartOrNull(part, id) is ChartPart chart)
                    {
                        outs.AddRange(PptChart.Render(chart).Select(b => new Out(b.IsTable ? OutKind.Table : OutKind.Paragraph, b.Text)));
                    }
                    else
                    {
                        outs.Add(new Out(OutKind.Paragraph, "[圖表]"));
                    }

                    items.Add(new Item(x, y, false, outs));
                    break;
                case "relIds":
                    _diagrams++;
                    items.Add(new Item(x, y, false, [new Out(OutKind.Paragraph, "[SmartArt]")]));
                    break;
                case "oleObj":
                    _embedded++;
                    items.Add(new Item(x, y, false, [new Out(OutKind.Paragraph, "[嵌入物件]")]));
                    break;
            }
        }
    }

    // ---- 排序與輸出 ----

    /// <summary>預留位置先（依文件順序），其餘依位置：垂直距離在容許範圍內的視為同一列，列內由左而右。</summary>
    private List<Item> Order(List<Item> items)
    {
        var result = items.Where(i => i.Placeholder).ToList();
        var rest = items.Where(i => !i.Placeholder).OrderBy(i => i.Y).ThenBy(i => i.X).ToList();
        for (var i = 0; i < rest.Count;)
        {
            var anchor = rest[i].Y;
            var j = i;
            while (j < rest.Count && rest[j].Y - anchor <= _rowTolerance)
            {
                j++;
            }

            result.AddRange(rest.GetRange(i, j - i).OrderBy(r => r.X).ThenBy(r => r.Y));
            i = j;
        }

        return result;
    }

    private static List<Out> Flatten(List<Item> items) => items.SelectMany(i => i.Outs).ToList();

    private void Emit(List<Out> outs)
    {
        for (var i = 0; i < outs.Count; i++)
        {
            var o = outs[i];
            if (o is { Kind: OutKind.Paragraph, Text: "[圖片]" })
            {
                var run = 1;
                while (i + run < outs.Count && outs[i + run] is { Kind: OutKind.Paragraph, Text: "[圖片]" })
                {
                    run++;
                }

                _builder.AddParagraph(run > 1 ? $"[圖片 ×{run.ToString(CultureInfo.InvariantCulture)}]" : "[圖片]");
                i += run - 1;
                continue;
            }

            switch (o.Kind)
            {
                case OutKind.ListItem:
                    _builder.AddListItem(o.Text, o.ListKey);
                    break;
                case OutKind.Table:
                    _builder.AddTable(o.Text);
                    break;
                default:
                    _builder.AddParagraph(o.Text);
                    break;
            }
        }
    }

    // ---- 文字 ----

    private List<Out> Paragraphs(List<A.Paragraph> paragraphs, OpenXmlPartContainer part, bool bulletByDefault)
    {
        var outs = new List<Out>();
        var key = ++_listKey;
        Array.Clear(_counters);

        foreach (var paragraph in paragraphs)
        {
            var pieces = Pieces(paragraph, part);
            if (!pieces.Any(p => p.Text.Trim().Length > 0))
            {
                continue;
            }

            var properties = paragraph.ParagraphProperties;
            var level = Math.Clamp(properties?.Level?.Value ?? 0, 0, MaxBulletLevel);
            var numbered = properties?.GetFirstChild<A.AutoNumberedBullet>() is not null;
            var bullet = numbered
                || properties?.GetFirstChild<A.CharacterBullet>() is not null
                || properties?.GetFirstChild<A.PictureBullet>() is not null
                || (bulletByDefault && properties?.GetFirstChild<A.NoBullet>() is null);

            if (!bullet)
            {
                Array.Clear(_counters);
                outs.Add(new Out(OutKind.Paragraph, MarkdownWriter.EscapeLineStarts(MarkdownWriter.RenderInline(pieces, emphasis: true, lineBreak: "  \n", inTable: false))));
                continue;
            }

            var marker = "- ";
            if (numbered)
            {
                Array.Clear(_counters, level + 1, MaxBulletLevel - level);
                marker = $"{(++_counters[level]).ToString(CultureInfo.InvariantCulture)}. ";
            }
            else
            {
                Array.Clear(_counters, level, MaxBulletLevel + 1 - level);
            }

            var indent = new string(' ', 2 * level);
            var text = MarkdownWriter.EscapeLineStarts(MarkdownWriter.RenderInline(pieces, emphasis: true, lineBreak: "\n", inTable: false));
            var continuation = "\n" + indent + new string(' ', marker.Length);
            outs.Add(new Out(OutKind.ListItem, indent + marker + text.Replace("\n", continuation, StringComparison.Ordinal), key));
        }

        return outs;
    }

    private static List<InlinePiece> Pieces(A.Paragraph paragraph, OpenXmlPartContainer part)
    {
        var pieces = new List<InlinePiece>();
        foreach (var child in paragraph.ChildElements)
        {
            switch (child)
            {
                case A.Run run:
                    pieces.Add(Piece(run.Text?.Text, run.RunProperties, part));
                    break;
                case A.Field field:
                    pieces.Add(Piece(field.Text?.Text, field.RunProperties, part));
                    break;
                case A.Break:
                    pieces.Add(new InlinePiece("\n"));
                    break;
            }
        }

        return pieces;
    }

    private static InlinePiece Piece(string? text, A.RunProperties? properties, OpenXmlPartContainer part)
    {
        string? link = null;
        if (properties?.GetFirstChild<A.HyperlinkOnClick>()?.Id?.Value is { Length: > 0 } id)
        {
            var uri = part.HyperlinkRelationships.FirstOrDefault(h => h.Id == id)?.Uri;
            if (uri is { IsAbsoluteUri: true } && uri.Scheme is "http" or "https" or "mailto")
            {
                link = uri.OriginalString;
            }
        }

        var strike = properties?.Strike?.Value is { } s && s != A.TextStrikeValues.NoStrike;
        return new InlinePiece(Sanitize(text ?? string.Empty), properties?.Bold?.Value == true, properties?.Italic?.Value == true, strike, link);
    }

    private static string PlainText(A.Paragraph paragraph, OpenXmlPartContainer part) =>
        MarkdownWriter.RenderInline(Pieces(paragraph, part).Select(p => p with { Link = null }).ToList(), emphasis: false, lineBreak: " ", inTable: false);

    /// <summary>OOXML 用 _xHHHH_ 表示 XML 不能直接放的字元（例如從 Excel 貼上的垂直定位字元 _x000B_），SDK 不會幫忙還原。</summary>
    private static string Sanitize(string text)
    {
        var decoded = text.Contains("_x", StringComparison.Ordinal)
            ? System.Text.RegularExpressions.Regex.Replace(
                text,
                "_x([0-9A-Fa-f]{4})_",
                m => ((char)int.Parse(m.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToString(),
                System.Text.RegularExpressions.RegexOptions.CultureInvariant,
                TimeSpan.FromSeconds(1))
            : text;

        var sb = new System.Text.StringBuilder(decoded.Length);
        foreach (var c in decoded)
        {
            sb.Append(c switch { '\v' => '\n', '\t' => ' ', < ' ' and not '\n' => '\0', _ => c });
        }

        return sb.ToString().Replace("\0", string.Empty, StringComparison.Ordinal);
    }

    private static string Clean(string text) =>
        System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ", System.Text.RegularExpressions.RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1)).Trim();

    private static bool IsHidden(P.NonVisualDrawingProperties? properties) => properties?.Hidden?.Value == true;

    // ---- 表格 ----

    /// <summary>被合併的儲存格（hMerge / vMerge）在 PPT 裡是獨立的空儲存格，這裡填入合併來源的值，讓每一列都有完整欄位。</summary>
    private static string? TableMarkdown(A.Table table, OpenXmlPartContainer part)
    {
        var rows = new List<List<string>>();
        foreach (var tr in table.Elements<A.TableRow>())
        {
            var row = new List<string>();
            foreach (var cell in tr.Elements<A.TableCell>())
            {
                var text = string.Join("<br>", (cell.TextBody?.Elements<A.Paragraph>() ?? [])
                    .Select(p => MarkdownWriter.RenderInline(Pieces(p, part), emphasis: true, lineBreak: "<br>", inTable: true).Trim())
                    .Where(x => x.Length > 0));

                if (cell.HorizontalMerge?.Value == true && row.Count > 0 && text.Length == 0)
                {
                    text = row[^1];
                }
                else if (cell.VerticalMerge?.Value == true && text.Length == 0 && rows.Count > 0 && row.Count < rows[^1].Count)
                {
                    text = rows[^1][row.Count];
                }

                row.Add(text);
            }

            rows.Add(row);
        }

        return rows.Count == 0 || rows.All(r => r.All(c => c.Trim().Length == 0)) ? null : MarkdownWriter.RenderTable(rows);
    }

    // ---- 備註 ----

    private static string NotesText(SlidePart part)
    {
        var tree = part.NotesSlidePart?.NotesSlide?.CommonSlideData?.ShapeTree;
        if (tree is null)
        {
            return string.Empty;
        }

        var lines = new List<string>();
        foreach (var shape in tree.Elements<P.Shape>())
        {
            var type = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value;
            if (type is null || type.Value != P.PlaceholderValues.Body)
            {
                continue;
            }

            foreach (var paragraph in shape.TextBody?.Elements<A.Paragraph>() ?? [])
            {
                var text = MarkdownWriter.EscapeLineStarts(MarkdownWriter.RenderInline(Pieces(paragraph, part.NotesSlidePart!), emphasis: false, lineBreak: " ", inTable: false).Trim());
                if (text.Length > 0)
                {
                    lines.Add(text);
                }
            }
        }

        return string.Join("\n", lines);
    }

    // ---- 其他 ----

    private void AddConversionNotes(List<string> notes)
    {
        if (_missingSlides > 0)
        {
            notes.Add($"有 {_missingSlides} 張投影片的內容部件遺失，已略過");
        }

        if (_hiddenSlides > 0)
        {
            notes.Add($"有 {_hiddenSlides} 張隱藏投影片（標題後標註「隱藏」）");
        }

        if (_totalPictures > 0)
        {
            notes.Add($"簡報有 {_totalPictures} 張圖片，只輸出替代文字（無替代文字時為 [圖片]）");
        }

        if (_totalCharts > 0)
        {
            notes.Add($"簡報有 {_totalCharts} 個圖表，輸出的是檔案內儲存的快取數據");
        }

        if (_diagrams > 0)
        {
            notes.Add($"簡報有 {_diagrams} 個 SmartArt，內容未轉換");
        }

        if (_embedded > 0)
        {
            notes.Add($"簡報有 {_embedded} 個嵌入物件，內容未轉換");
        }
    }

    /// <summary>GetPartById 找不到關聯會丟例外；壞檔裡的關聯指向不存在的部件很常見，這裡當成「沒有」。</summary>
    private static OpenXmlPart? PartOrNull(OpenXmlPartContainer container, string relationshipId) =>
        container.Parts.Where(p => p.RelationshipId == relationshipId).Select(p => p.OpenXmlPart).FirstOrDefault();

    private static OpenXmlElement Chosen(OpenXmlElement alternate) =>
        alternate.ChildElements.FirstOrDefault(e => e.LocalName == "Choice")
        ?? alternate.ChildElements.FirstOrDefault(e => e.LocalName == "Fallback")
        ?? alternate;
}
