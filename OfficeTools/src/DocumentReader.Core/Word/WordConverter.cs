using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentReader.Core.Markdown;
using DocumentReader.Core.Models;

namespace DocumentReader.Core.Word;

/// <summary>把 Word 文件本體轉成 Markdown 區塊。每次轉換建一個新的實例（內含編號計數等狀態）。</summary>
internal sealed class WordConverter
{
    private readonly MainDocumentPart _main;
    private readonly WordReadOptions _options;
    private readonly WordStyles _styles;
    private readonly WordNumbering _numbering;
    private readonly MarkdownDocumentBuilder _builder = new();

    private int _images;
    private int _charts;
    private int _diagrams;
    private int _equations;
    private int _tocParagraphs;
    private bool _hasRevisions;

    public WordConverter(MainDocumentPart main, WordReadOptions options)
    {
        _main = main;
        _options = options;
        _styles = WordStyles.Load(main);
        _numbering = WordNumbering.Load(main);
    }

    public ParsedDocument Convert()
    {
        if (_main.Document?.Body is { } body)
        {
            Blocks(body, _main);
        }

        var notes = new List<string>();
        AppendAppendices(notes);
        AddConversionNotes(notes);
        return _builder.Build(notes);
    }

    // ---- 區塊 ----

    private void Blocks(OpenXmlElement container, OpenXmlPartContainer part)
    {
        foreach (var child in container.ChildElements)
        {
            switch (child.LocalName)
            {
                case "p":
                    Paragraph((Paragraph)child, part);
                    break;
                case "tbl":
                    var table = TableMarkdown((Table)child, part);
                    if (table is not null)
                    {
                        _builder.AddTable(table);
                    }

                    break;
                case "sdt":
                    if (!IsTocControl(child))
                    {
                        Blocks(ContentOf(child), part);
                    }
                    else
                    {
                        _tocParagraphs++;
                    }

                    break;
                case "ins" or "moveTo" or "customXml" or "smartTag":
                    _hasRevisions |= child.LocalName is "ins" or "moveTo";
                    Blocks(child, part);
                    break;
                case "del" or "moveFrom":
                    _hasRevisions = true;
                    break;
                case "AlternateContent":
                    Blocks(Chosen(child), part);
                    break;
            }
        }
    }

    private void Paragraph(Paragraph p, OpenXmlPartContainer part)
    {
        var properties = p.ParagraphProperties;
        var styleId = properties?.ParagraphStyleId?.Val?.Value;
        if (_styles.IsToc(styleId))
        {
            _tocParagraphs++;
            return;
        }

        var pieces = new List<InlinePiece>();
        var side = new List<string>();
        Inline(p, part, link: null, pieces, side);

        var hasText = pieces.Any(x => x.Text.Trim().Length > 0);
        if (!hasText && side.Count == 0)
        {
            return; // 空段落（包含整段被刪除的段落）不消耗編號
        }

        var level = _styles.HeadingLevel(styleId, properties?.OutlineLevel?.Val?.Value);
        var number = _styles.Numbering(styleId, properties?.NumberingProperties);
        var label = number is { } n ? _numbering.Next(n.NumId, n.Ilvl) : null;

        if (hasText && level is { } headingLevel)
        {
            var title = MarkdownWriter.RenderInline(pieces.Select(x => x with { Link = null }).ToList(), emphasis: false, lineBreak: " ", inTable: false);
            var prefix = label is { IsBullet: false, Text.Length: > 0 } ? label.Text + " " : string.Empty;
            _builder.AddHeading(headingLevel, prefix + title);
        }
        else if (hasText && label is not null)
        {
            _builder.AddListItem(ListItem(label, number!.Value.Ilvl, MarkdownWriter.RenderInline(pieces, emphasis: true, lineBreak: "\n", inTable: false)), number.Value.NumId);
        }
        else if (hasText)
        {
            _builder.AddParagraph(MarkdownWriter.EscapeLineStarts(MarkdownWriter.RenderInline(pieces, emphasis: true, lineBreak: "  \n", inTable: false)));
        }

        foreach (var text in side)
        {
            _builder.AddParagraph("[文字方塊] " + text);
        }
    }

    private static string ListItem(NumberingLabel label, int ilvl, string text)
    {
        var indent = new string(' ', 2 * Math.Clamp(ilvl, 0, 8));
        var marker = label.IsBullet ? "- " : $"{label.Counter.ToString(CultureInfo.InvariantCulture)}. ";
        var continuation = "\n" + indent + new string(' ', marker.Length); // 項目內的換行要縮排，才會留在同一個項目裡
        return indent + marker + MarkdownWriter.EscapeLineStarts(text).Replace("\n", continuation, StringComparison.Ordinal);
    }

    // ---- 行內 ----

    private void Inline(OpenXmlElement container, OpenXmlPartContainer part, string? link, List<InlinePiece> pieces, List<string> side)
    {
        foreach (var child in container.ChildElements)
        {
            switch (child.LocalName)
            {
                case "r":
                    RunContent((Run)child, part, link, pieces, side);
                    break;
                case "hyperlink":
                    var relationship = ((Hyperlink)child).Id?.Value;
                    Inline(child, part, relationship is null ? null : ResolveLink(part, relationship), pieces, side);
                    break;
                case "ins" or "moveTo":
                    _hasRevisions = true;
                    Inline(child, part, link, pieces, side);
                    break;
                case "del" or "moveFrom":
                    _hasRevisions = true; // 輸出接受全部修訂後的版本：被刪除的內容不輸出
                    break;
                case "fldSimple" or "smartTag" or "customXml":
                    Inline(child, part, link, pieces, side);
                    break;
                case "sdt":
                    Inline(ContentOf(child), part, link, pieces, side);
                    break;
                case "oMath" or "oMathPara":
                    _equations++;
                    var math = string.Concat(child.Descendants().Where(e => e.LocalName == "t").Select(e => e.InnerText)).Trim();
                    if (math.Length > 0)
                    {
                        pieces.Add(new InlinePiece($"[公式: {math}]", Link: link));
                    }

                    break;
                case "AlternateContent":
                    Inline(Chosen(child), part, link, pieces, side);
                    break;
            }
        }
    }

    private void RunContent(Run run, OpenXmlPartContainer part, string? link, List<InlinePiece> pieces, List<string> side)
    {
        if (run.RunProperties?.Vanish is { } hidden && (hidden.Val?.Value ?? true))
        {
            return; // 隱藏文字
        }

        var format = _styles.Format(run.RunProperties, run.RunProperties?.RunStyle?.Val?.Value);
        void Text(string text) => pieces.Add(new InlinePiece(text, format.Bold, format.Italic, format.Strike, link));

        foreach (var child in run.ChildElements)
        {
            switch (child.LocalName)
            {
                case "t":
                    Text(((Text)child).Text);
                    break;
                case "tab" or "ptab":
                    Text(" ");
                    break;
                case "br":
                    var breakType = ((Break)child).Type?.Value;
                    if (breakType != BreakValues.Page && breakType != BreakValues.Column) // 分頁、分欄不是文字換行
                    {
                        pieces.Add(new InlinePiece("\n"));
                    }

                    break;
                case "cr":
                    pieces.Add(new InlinePiece("\n"));
                    break;
                case "noBreakHyphen":
                    Text("-");
                    break;
                case "drawing":
                    Drawing(child, part, pieces, side, link);
                    break;
                case "pict":
                    Picture(child, part, pieces, side, link);
                    break;
                case "object":
                    pieces.Add(new InlinePiece("[嵌入物件]", Link: link));
                    break;
                case "footnoteReference" when _options.IncludeFootnotes:
                    pieces.Add(new InlinePiece($"[^f{((FootnoteReference)child).Id?.Value}]"));
                    break;
                case "endnoteReference" when _options.IncludeFootnotes:
                    pieces.Add(new InlinePiece($"[^e{((EndnoteReference)child).Id?.Value}]"));
                    break;
                case "AlternateContent":
                    var holder = new Run(Chosen(child).ChildElements.Select(e => e.CloneNode(true)));
                    RunContent(holder, part, link, pieces, side);
                    break;
            }
        }
    }

    private void Drawing(OpenXmlElement drawing, OpenXmlPartContainer part, List<InlinePiece> pieces, List<string> side, string? link)
    {
        foreach (var box in TextBoxes(drawing))
        {
            AddTextBox(box, part, side);
        }

        var uri = Attribute(drawing.Descendants().FirstOrDefault(e => e.LocalName == "graphicData"), "uri") ?? string.Empty;
        var alt = AltText(Attribute(drawing.Descendants().FirstOrDefault(e => e.LocalName == "docPr"), "descr")
            ?? Attribute(drawing.Descendants().FirstOrDefault(e => e.LocalName == "docPr"), "title"));

        if (uri.Contains("chart", StringComparison.OrdinalIgnoreCase))
        {
            _charts++;
            pieces.Add(new InlinePiece(alt.Length > 0 ? $"[圖表: {alt}]" : "[圖表]", Link: link));
        }
        else if (uri.Contains("diagram", StringComparison.OrdinalIgnoreCase))
        {
            _diagrams++;
            pieces.Add(new InlinePiece(alt.Length > 0 ? $"[SmartArt: {alt}]" : "[SmartArt]", Link: link));
        }
        else if (uri.Contains("picture", StringComparison.OrdinalIgnoreCase) || drawing.Descendants().Any(e => e.LocalName == "blip"))
        {
            _images++;
            pieces.Add(new InlinePiece(alt.Length > 0 ? $"[圖片: {alt}]" : "[圖片]", Link: link));
        }
        // 其他（沒有圖片也沒有文字的圖形，例如裝飾線）不輸出
    }

    private void Picture(OpenXmlElement picture, OpenXmlPartContainer part, List<InlinePiece> pieces, List<string> side, string? link)
    {
        foreach (var box in TextBoxes(picture))
        {
            AddTextBox(box, part, side);
        }

        if (picture.Descendants().Any(e => e.LocalName == "imagedata"))
        {
            _images++;
            var shape = picture.Descendants().FirstOrDefault(e => e.LocalName == "shape");
            var alt = AltText(Attribute(shape, "alt") ?? Attribute(shape, "title"));
            pieces.Add(new InlinePiece(alt.Length > 0 ? $"[圖片: {alt}]" : "[圖片]", Link: link));
        }
    }

    private static IEnumerable<OpenXmlElement> TextBoxes(OpenXmlElement root) =>
        root.Descendants().Where(e => e.LocalName == "txbxContent" && !e.Ancestors().Any(a => a.LocalName == "txbxContent" && root.Descendants().Contains(a)));

    private void AddTextBox(OpenXmlElement box, OpenXmlPartContainer part, List<string> side)
    {
        var lines = new List<string>();
        foreach (var paragraph in box.ChildElements.OfType<Paragraph>())
        {
            var pieces = new List<InlinePiece>();
            Inline(paragraph, part, link: null, pieces, []);
            var text = MarkdownWriter.RenderInline(pieces, emphasis: false, lineBreak: " ", inTable: false).Trim();
            if (text.Length > 0)
            {
                lines.Add(text);
            }
        }

        if (lines.Count > 0)
        {
            side.Add(string.Join(" ", lines));
        }
    }

    // ---- 表格 ----

    private string? TableMarkdown(Table table, OpenXmlPartContainer part)
    {
        var rows = TableRows(table, part);
        if (rows.Count == 0 || rows.All(r => r.All(c => c.Trim().Length == 0)))
        {
            return null;
        }

        return MarkdownWriter.RenderTable(rows);
    }

    /// <summary>
    /// 合併儲存格：橫向合併（gridSpan）與直向合併（vMerge）都把值重複填進被合併的每一格，
    /// 讓每一列都有完整的欄位，agent 不必推算合併關係。
    /// </summary>
    private List<List<string>> TableRows(Table table, OpenXmlPartContainer part)
    {
        var rows = new List<List<string>>();
        foreach (var tr in table.ChildElements.OfType<TableRow>())
        {
            var row = new List<string>();
            var before = tr.TableRowProperties?.Elements<GridBefore>().FirstOrDefault()?.Val?.Value ?? 0;
            row.AddRange(Enumerable.Repeat(string.Empty, before));

            foreach (var cell in tr.ChildElements.OfType<TableCell>())
            {
                var properties = cell.TableCellProperties;
                var span = Math.Max(1, properties?.GridSpan?.Val?.Value ?? 1);
                var text = CellText(cell, part);

                var vertical = properties?.VerticalMerge;
                if (vertical is not null && (vertical.Val is null || vertical.Val.Value == MergedCellValues.Continue))
                {
                    var above = rows.Count > 0 && row.Count < rows[^1].Count ? rows[^1][row.Count] : string.Empty;
                    text = text.Length > 0 ? text : above;
                }

                row.AddRange(Enumerable.Repeat(text, span));
            }

            var after = tr.TableRowProperties?.Elements<GridAfter>().FirstOrDefault()?.Val?.Value ?? 0;
            row.AddRange(Enumerable.Repeat(string.Empty, after));
            rows.Add(row);
        }

        return rows;
    }

    private string CellText(TableCell cell, OpenXmlPartContainer part)
    {
        var parts = new List<string>();
        CellBlocks(cell, part, parts);
        return string.Join("<br>", parts);
    }

    private void CellBlocks(OpenXmlElement container, OpenXmlPartContainer part, List<string> parts)
    {
        foreach (var child in container.ChildElements)
        {
            switch (child.LocalName)
            {
                case "p":
                    var paragraph = (Paragraph)child;
                    var pieces = new List<InlinePiece>();
                    var side = new List<string>();
                    Inline(paragraph, part, link: null, pieces, side);

                    var text = MarkdownWriter.RenderInline(pieces, emphasis: true, lineBreak: "<br>", inTable: true).Trim();
                    var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    if (text.Length > 0 && _styles.Numbering(styleId, paragraph.ParagraphProperties?.NumberingProperties) is { } n)
                    {
                        var label = _numbering.Next(n.NumId, n.Ilvl);
                        text = (label.IsBullet ? "• " : $"{label.Counter.ToString(CultureInfo.InvariantCulture)}. ") + text;
                    }

                    if (text.Length > 0)
                    {
                        parts.Add(text);
                    }

                    parts.AddRange(side.Select(s => "[文字方塊] " + MarkdownWriter.Escape(s, inTable: true)));
                    break;
                case "tbl":
                    var nested = TableRows((Table)child, part);
                    if (nested.Count > 0)
                    {
                        // 巢狀表格降級為文字：列以 "; " 分隔、格以 " / " 分隔
                        parts.Add("（巢狀表格：" + string.Join("; ", nested.Select(r => string.Join(" / ", r.Where(c => c.Length > 0)))) + "）");
                    }

                    break;
                case "sdt":
                    CellBlocks(ContentOf(child), part, parts);
                    break;
                case "ins" or "moveTo" or "customXml" or "smartTag":
                    CellBlocks(child, part, parts);
                    break;
            }
        }
    }

    // ---- 附錄（頁首頁尾、註腳、註解）與轉換說明 ----

    private void AppendAppendices(List<string> notes)
    {
        var headerFooter = HeaderFooterLines();
        if (_options.IncludeHeadersFooters && headerFooter.Count > 0)
        {
            _builder.AddHeading(1, "附錄：頁首頁尾");
            foreach (var line in headerFooter)
            {
                _builder.AddListItem("- " + line);
            }
        }
        else if (headerFooter.Count > 0)
        {
            notes.Add("文件有頁首或頁尾，預設不輸出（可用 IncludeHeadersFooters 開啟）");
        }

        var footnotes = NoteLines();
        if (_options.IncludeFootnotes && footnotes.Count > 0)
        {
            _builder.AddHeading(1, "附錄：註腳");
            foreach (var line in footnotes)
            {
                _builder.AddParagraph(line);
            }
        }
        else if (footnotes.Count > 0)
        {
            notes.Add($"文件有 {footnotes.Count} 則註腳 / 章節附註，預設不輸出（可用 IncludeFootnotes 開啟）");
        }

        var comments = CommentLines();
        if (_options.IncludeComments && comments.Count > 0)
        {
            _builder.AddHeading(1, "附錄：註解");
            foreach (var line in comments)
            {
                _builder.AddListItem("- " + line);
            }
        }
        else if (comments.Count > 0)
        {
            notes.Add($"文件有 {comments.Count} 則註解，預設不輸出（可用 IncludeComments 開啟）");
        }
    }

    private List<string> HeaderFooterLines()
    {
        var lines = new List<string>();
        void Collect(string label, IEnumerable<OpenXmlElement> roots, OpenXmlPartContainer part)
        {
            foreach (var root in roots)
            {
                var text = PlainParagraphs(root, part);
                var line = $"{label}：{text}";
                if (text.Length > 0 && !lines.Contains(line))
                {
                    lines.Add(line);
                }
            }
        }

        Collect("頁首", _main.HeaderParts.Where(h => h.Header is not null).Select(h => (OpenXmlElement)h.Header!), _main);
        Collect("頁尾", _main.FooterParts.Where(f => f.Footer is not null).Select(f => (OpenXmlElement)f.Footer!), _main);
        return lines;
    }

    private List<string> NoteLines()
    {
        var lines = new List<string>();
        if (_main.FootnotesPart?.Footnotes is { } footnotes)
        {
            foreach (var note in footnotes.Elements<Footnote>().Where(f => f.Type is null || f.Type.Value == FootnoteEndnoteValues.Normal))
            {
                var text = PlainParagraphs(note, _main.FootnotesPart);
                if (text.Length > 0)
                {
                    lines.Add($"[^f{note.Id?.Value}]: {text}");
                }
            }
        }

        if (_main.EndnotesPart?.Endnotes is { } endnotes)
        {
            foreach (var note in endnotes.Elements<Endnote>().Where(f => f.Type is null || f.Type.Value == FootnoteEndnoteValues.Normal))
            {
                var text = PlainParagraphs(note, _main.EndnotesPart);
                if (text.Length > 0)
                {
                    lines.Add($"[^e{note.Id?.Value}]: {text}");
                }
            }
        }

        return lines;
    }

    private List<string> CommentLines()
    {
        var lines = new List<string>();
        if (_main.WordprocessingCommentsPart?.Comments is { } comments)
        {
            foreach (var comment in comments.Elements<Comment>())
            {
                var text = PlainParagraphs(comment, _main.WordprocessingCommentsPart);
                if (text.Length > 0)
                {
                    var author = comment.Author?.Value;
                    lines.Add(string.IsNullOrWhiteSpace(author) ? text : $"{author}：{text}");
                }
            }
        }

        return lines;
    }

    /// <summary>附錄用的純文字：每個段落轉成一行（不含強調標記），段落之間以空白連接。</summary>
    private string PlainParagraphs(OpenXmlElement root, OpenXmlPartContainer part)
    {
        var texts = new List<string>();
        foreach (var paragraph in root.Descendants<Paragraph>().Where(p => !p.Ancestors().Any(a => a.LocalName == "txbxContent")))
        {
            var pieces = new List<InlinePiece>();
            Inline(paragraph, part, link: null, pieces, []);
            var text = MarkdownWriter.RenderInline(pieces, emphasis: false, lineBreak: " ", inTable: false).Trim();
            if (text.Length > 0)
            {
                texts.Add(text);
            }
        }

        return string.Join(" ", texts);
    }

    private void AddConversionNotes(List<string> notes)
    {
        if (_tocParagraphs > 0)
        {
            notes.Add("已略過目錄（目錄只是標題加頁碼）");
        }

        if (_images > 0)
        {
            notes.Add($"文件有 {_images} 張圖片，只輸出替代文字（無替代文字時為 [圖片]）");
        }

        if (_charts > 0)
        {
            notes.Add($"文件有 {_charts} 個圖表，數據未轉換");
        }

        if (_diagrams > 0)
        {
            notes.Add($"文件有 {_diagrams} 個 SmartArt，內容未轉換");
        }

        if (_equations > 0)
        {
            notes.Add($"文件有 {_equations} 個數學公式，只轉成純文字");
        }

        if (_hasRevisions)
        {
            notes.Add("文件有追蹤修訂，輸出的是接受全部修訂後的版本");
        }
    }

    // ---- 小工具 ----

    private static string? ResolveLink(OpenXmlPartContainer part, string relationshipId) =>
        part.HyperlinkRelationships.FirstOrDefault(r => r.Id == relationshipId)?.Uri?.OriginalString; // 不用 ToString()：它會補上尾端斜線並把 %20 還原成空白

    private static OpenXmlElement ContentOf(OpenXmlElement sdt) =>
        sdt.ChildElements.FirstOrDefault(e => e.LocalName == "sdtContent") ?? sdt;

    private static bool IsTocControl(OpenXmlElement sdt) =>
        sdt.ChildElements.FirstOrDefault(e => e.LocalName == "sdtPr")?.Descendants()
            .Any(e => e.LocalName == "docPartGallery" && (Attribute(e, "val")?.Contains("Table of Contents", StringComparison.OrdinalIgnoreCase) ?? false)) ?? false;

    /// <summary>mc:AlternateContent 只處理第一個 Choice（Fallback 是同樣內容的相容版本，會重複）。</summary>
    private static OpenXmlElement Chosen(OpenXmlElement alternate) =>
        alternate.ChildElements.FirstOrDefault(e => e.LocalName == "Choice")
        ?? alternate.ChildElements.FirstOrDefault(e => e.LocalName == "Fallback")
        ?? alternate;

    private static string? Attribute(OpenXmlElement? element, string localName) =>
        element?.GetAttributes().Where(a => a.LocalName == localName).Select(a => a.Value).FirstOrDefault();

    private static string AltText(string? alt) =>
        alt is null ? string.Empty : System.Text.RegularExpressions.Regex.Replace(alt, @"\s+", " ", System.Text.RegularExpressions.RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1)).Trim();
}
