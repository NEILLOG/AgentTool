using System.IO.Compression;
using System.Security;
using System.Text;

namespace DocumentReader.Core.Tests;

/// <summary>用原始 XML 組出符合規格的 .docx，才能精確控制每一種結構（樣式、編號、修訂、文字方塊…）。</summary>
internal sealed class DocxBuilder
{
    public const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private readonly StringBuilder _body = new();
    private readonly List<(string Path, string Xml, string ContentType, string RelType)> _parts = [];
    private readonly List<(string Id, string Url)> _links = [];
    private int _nextRel = 10;

    public bool ChineseStyleIds { get; init; } = true;

    public string? ExtraStyles { get; set; }

    public string? NumberingXml { get; set; } = DefaultNumbering;

    // ---- 建構 ----

    public DocxBuilder Add(params string[] blocks)
    {
        foreach (var block in blocks)
        {
            _body.Append(block);
        }

        return this;
    }

    public string Link(string url)
    {
        var id = $"rId{_nextRel++}";
        _links.Add((id, url));
        return id;
    }

    public DocxBuilder Header(string text, string name = "header1") =>
        Part($"word/{name}.xml", $"<w:hdr xmlns:w=\"{W}\">{Para(text)}</w:hdr>", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", "header");

    public DocxBuilder Footer(string text, string name = "footer1") =>
        Part($"word/{name}.xml", $"<w:ftr xmlns:w=\"{W}\">{Para(text)}</w:ftr>", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml", "footer");

    public DocxBuilder Footnotes(params (int Id, string Text)[] notes) =>
        Part(
            "word/footnotes.xml",
            $"<w:footnotes xmlns:w=\"{W}\"><w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
            + string.Concat(notes.Select(n => $"<w:footnote w:id=\"{n.Id}\"><w:p><w:r><w:footnoteRef/></w:r>{Run(" " + n.Text)}</w:p></w:footnote>"))
            + "</w:footnotes>",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            "footnotes");

    public DocxBuilder Comments(params (int Id, string Author, string Text)[] comments) =>
        Part(
            "word/comments.xml",
            $"<w:comments xmlns:w=\"{W}\">" + string.Concat(comments.Select(c => $"<w:comment w:id=\"{c.Id}\" w:author=\"{Esc(c.Author)}\">{Para(c.Text)}</w:comment>")) + "</w:comments>",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
            "comments");

    public DocxBuilder Part(string path, string xml, string contentType, string relType)
    {
        _parts.Add((path, xml, contentType, relType));
        return this;
    }

    public byte[] Build()
    {
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Entry(string name, string xml)
            {
                using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + xml);
            }

            var overrides = new StringBuilder()
                .Append("<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>")
                .Append("<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>");
            if (NumberingXml is not null)
            {
                overrides.Append("<Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>");
            }

            foreach (var part in _parts)
            {
                overrides.Append($"<Override PartName=\"/{part.Path}\" ContentType=\"{part.ContentType}\"/>");
            }

            Entry("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>" + overrides + "</Types>");
            Entry("_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"" + R + "/officeDocument\" Target=\"word/document.xml\"/></Relationships>");

            var rels = new StringBuilder("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            rels.Append("<Relationship Id=\"rId1\" Type=\"" + R + "/styles\" Target=\"styles.xml\"/>");
            if (NumberingXml is not null)
            {
                rels.Append("<Relationship Id=\"rId2\" Type=\"" + R + "/numbering\" Target=\"numbering.xml\"/>");
            }

            for (var i = 0; i < _parts.Count; i++)
            {
                rels.Append($"<Relationship Id=\"rIdP{i}\" Type=\"{R}/{_parts[i].RelType}\" Target=\"{_parts[i].Path["word/".Length..]}\"/>");
            }

            foreach (var (id, url) in _links)
            {
                rels.Append($"<Relationship Id=\"{id}\" Type=\"{R}/hyperlink\" Target=\"{Esc(url)}\" TargetMode=\"External\"/>");
            }

            Entry("word/_rels/document.xml.rels", rels + "</Relationships>");
            Entry("word/document.xml", Document());
            Entry("word/styles.xml", Styles());
            if (NumberingXml is not null)
            {
                Entry("word/numbering.xml", NumberingXml);
            }

            foreach (var part in _parts)
            {
                Entry(part.Path, part.Xml);
            }
        }

        return ms.ToArray();
    }

    public string Save(string path)
    {
        File.WriteAllBytes(path, Build());
        return path;
    }

    private string Document() =>
        $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" " +
        "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" " +
        "xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" " +
        $"xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><w:body>{_body}<w:sectPr/></w:body></w:document>";

    private string Styles()
    {
        var sb = new StringBuilder($"<w:styles xmlns:w=\"{W}\">");
        sb.Append("<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>");
        for (var level = 1; level <= 9; level++)
        {
            var id = ChineseStyleIds ? level.ToString(System.Globalization.CultureInfo.InvariantCulture) : $"Heading{level}";
            // 中文版 Word 的內建標題樣式 ID 是 "1"、"2"…，名稱才是 "heading 1"
            sb.Append($"<w:style w:type=\"paragraph\" w:styleId=\"{id}\"><w:name w:val=\"heading {level}\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:outlineLvl w:val=\"{level - 1}\"/></w:pPr></w:style>");
        }

        sb.Append("<w:style w:type=\"paragraph\" w:styleId=\"Title\"><w:name w:val=\"Title\"/><w:basedOn w:val=\"Normal\"/></w:style>");
        sb.Append("<w:style w:type=\"character\" w:styleId=\"Strong\"><w:name w:val=\"Strong\"/><w:rPr><w:b/></w:rPr></w:style>");
        sb.Append("<w:style w:type=\"character\" w:styleId=\"Emphasis\"><w:name w:val=\"Emphasis\"/><w:rPr><w:i/></w:rPr></w:style>");
        sb.Append("<w:style w:type=\"paragraph\" w:styleId=\"ListBullet\"><w:name w:val=\"List Bullet\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:numPr><w:numId w:val=\"1\"/></w:numPr></w:pPr></w:style>");
        sb.Append("<w:style w:type=\"paragraph\" w:styleId=\"TOC1\"><w:name w:val=\"toc 1\"/><w:basedOn w:val=\"Normal\"/></w:style>");
        sb.Append("<w:style w:type=\"paragraph\" w:styleId=\"TOCHeading\"><w:name w:val=\"TOC Heading\"/><w:basedOn w:val=\"Normal\"/></w:style>");
        sb.Append("<w:style w:type=\"paragraph\" w:styleId=\"BodyOutline9\"><w:name w:val=\"Body Outline9\"/><w:basedOn w:val=\"1\"/><w:pPr><w:outlineLvl w:val=\"9\"/></w:pPr></w:style>");
        sb.Append("<w:style w:type=\"paragraph\" w:styleId=\"MyHeading\"><w:name w:val=\"My Heading\"/><w:basedOn w:val=\"2\"/></w:style>");
        sb.Append("<w:style w:type=\"paragraph\" w:styleId=\"CustomChinese\"><w:name w:val=\"標題 1\"/><w:basedOn w:val=\"Normal\"/></w:style>");
        sb.Append(ExtraStyles);
        return sb.Append("</w:styles>").ToString();
    }

    // ---- XML 小工具（測試中直接呼叫）----

    public static string Esc(string text) => SecurityElement.Escape(text) ?? string.Empty;

    public static string Run(string text, bool b = false, bool i = false, bool strike = false, string? style = null, bool vanish = false, bool boldOff = false)
    {
        var props = (style is null ? string.Empty : $"<w:rStyle w:val=\"{style}\"/>")
            + (b ? "<w:b/>" : string.Empty) + (boldOff ? "<w:b w:val=\"0\"/>" : string.Empty)
            + (i ? "<w:i/>" : string.Empty) + (strike ? "<w:strike/>" : string.Empty) + (vanish ? "<w:vanish/>" : string.Empty);
        return $"<w:r>{(props.Length > 0 ? $"<w:rPr>{props}</w:rPr>" : string.Empty)}<w:t xml:space=\"preserve\">{Esc(text)}</w:t></w:r>";
    }

    /// <summary>以原始 XML（多個 run 等）組段落。</summary>
    public static string P(string inner, string? style = null, int? numId = null, int ilvl = 0, int? outline = null)
    {
        var props = (style is null ? string.Empty : $"<w:pStyle w:val=\"{style}\"/>")
            + (numId is null ? string.Empty : $"<w:numPr><w:ilvl w:val=\"{ilvl}\"/><w:numId w:val=\"{numId}\"/></w:numPr>")
            + (outline is null ? string.Empty : $"<w:outlineLvl w:val=\"{outline}\"/>");
        return $"<w:p>{(props.Length > 0 ? $"<w:pPr>{props}</w:pPr>" : string.Empty)}{inner}</w:p>";
    }

    /// <summary>純文字段落。</summary>
    public static string Para(string text, string? style = null, int? numId = null, int ilvl = 0, int? outline = null) =>
        P(Run(text), style, numId, ilvl, outline);

    public static string H(int level, string text, bool chineseIds = true) =>
        Para(text, chineseIds ? level.ToString(System.Globalization.CultureInfo.InvariantCulture) : $"Heading{level}");

    public static string Br() => "<w:r><w:br/></w:r>";

    public static string PageBreak() => "<w:r><w:br w:type=\"page\"/></w:r>";

    public static string Tab() => "<w:r><w:tab/></w:r>";

    public static string Hyperlink(string relationshipId, string inner) => $"<w:hyperlink r:id=\"{relationshipId}\">{inner}</w:hyperlink>";

    public static string Anchor(string anchor, string inner) => $"<w:hyperlink w:anchor=\"{anchor}\">{inner}</w:hyperlink>";

    public static string Ins(string inner) => $"<w:ins w:id=\"1\" w:author=\"a\" w:date=\"2026-01-01T00:00:00Z\">{inner}</w:ins>";

    public static string Del(string text) => $"<w:del w:id=\"2\" w:author=\"a\" w:date=\"2026-01-01T00:00:00Z\"><w:r><w:delText xml:space=\"preserve\">{Esc(text)}</w:delText></w:r></w:del>";

    // ---- 表格 ----

    public static string Cell(string text, int span = 1, string? vMerge = null)
    {
        var props = (span > 1 ? $"<w:gridSpan w:val=\"{span}\"/>" : string.Empty)
            + (vMerge is null ? string.Empty : vMerge.Length == 0 ? "<w:vMerge/>" : $"<w:vMerge w:val=\"{vMerge}\"/>");
        return $"<w:tc>{(props.Length > 0 ? $"<w:tcPr>{props}</w:tcPr>" : string.Empty)}{(text.StartsWith("<", StringComparison.Ordinal) ? text : Para(text))}</w:tc>";
    }

    public static string Row(params string[] cells) => $"<w:tr>{string.Concat(cells)}</w:tr>";

    public static string Table(params string[] rows) => $"<w:tbl><w:tblPr/>{string.Concat(rows)}</w:tbl>";

    /// <summary>每一格是純文字的簡單表格。</summary>
    public static string SimpleTable(params string[][] rows) => Table(rows.Select(r => Row(r.Select(c => Cell(c)).ToArray())).ToArray());

    // ---- 圖片與其他 ----

    public static string Image(string? alt, string uri = "http://schemas.openxmlformats.org/drawingml/2006/picture") =>
        $"<w:r><w:drawing><wp:inline><wp:docPr id=\"1\" name=\"Picture 1\"{(alt is null ? string.Empty : $" descr=\"{Esc(alt)}\"")}/><a:graphic><a:graphicData uri=\"{uri}\"><pic:pic><pic:blipFill><a:blip/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

    public static string Chart(string? alt = null) =>
        $"<w:r><w:drawing><wp:inline><wp:docPr id=\"2\" name=\"Chart 1\"{(alt is null ? string.Empty : $" descr=\"{Esc(alt)}\"")}/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

    public static string SmartArt() =>
        "<w:r><w:drawing><wp:inline><wp:docPr id=\"3\" name=\"Diagram 1\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/></a:graphic></wp:inline></w:drawing></w:r>";

    /// <summary>Word 的文字方塊：Choice 是新式 (wps)，Fallback 是同樣內容的 VML 相容版本。</summary>
    public static string TextBox(string text) =>
        "<w:r><mc:AlternateContent>"
        + $"<mc:Choice Requires=\"wps\"><w:drawing><wp:anchor><wp:docPr id=\"4\" name=\"Text Box 1\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:wsp><wps:txbx><w:txbxContent>{Para(text)}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>"
        + $"<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>{Para(text)}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>"
        + "</mc:AlternateContent></w:r>";

    public static string VmlImage(string? alt) =>
        $"<w:r><w:pict><v:shape{(alt is null ? string.Empty : $" alt=\"{Esc(alt)}\"")}><v:imagedata r:id=\"rId99\"/></v:shape></w:pict></w:r>";

    public static string Math(string text) => $"<m:oMath><m:r><m:t>{Esc(text)}</m:t></m:r></m:oMath>";

    public static string SimpleField(string result) => $"<w:fldSimple w:instr=\" PAGE \">{Run(result)}</w:fldSimple>";

    public static string ComplexField(string instruction, string result) =>
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
        + $"<w:r><w:instrText xml:space=\"preserve\"> {Esc(instruction)} </w:instrText></w:r>"
        + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
        + Run(result)
        + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    public static string InlineSdt(string inner) => $"<w:sdt><w:sdtPr/><w:sdtContent>{inner}</w:sdtContent></w:sdt>";

    public static string BlockSdt(string blocks, string? gallery = null) =>
        $"<w:sdt><w:sdtPr>{(gallery is null ? string.Empty : $"<w:docPartObj><w:docPartGallery w:val=\"{gallery}\"/></w:docPartObj>")}</w:sdtPr><w:sdtContent>{blocks}</w:sdtContent></w:sdt>";

    public static string FootnoteRef(int id) => $"<w:r><w:footnoteReference w:id=\"{id}\"/></w:r>";

    // ---- 編號定義 ----

    /// <summary>
    /// numId 1：項目符號；2、3：十進位多層清單（兩個獨立的清單）；4：標題編號（%1.%2）；
    /// 5：起始值覆寫為 5；6：小寫字母；7：大寫羅馬；8：法律式編號（%1.%2，上層一律十進位）。
    /// </summary>
    public static readonly string DefaultNumbering =
        $"<w:numbering xmlns:w=\"{W}\">"
        + "<w:abstractNum w:abstractNumId=\"0\">" + string.Join(string.Empty, Enumerable.Range(0, 9).Select(l => $"<w:lvl w:ilvl=\"{l}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"•\"/></w:lvl>")) + "</w:abstractNum>"
        + "<w:abstractNum w:abstractNumId=\"1\">" + string.Join(string.Empty, Enumerable.Range(0, 9).Select(l => $"<w:lvl w:ilvl=\"{l}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%{l + 1}.\"/></w:lvl>")) + "</w:abstractNum>"
        + "<w:abstractNum w:abstractNumId=\"2\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.%2\"/></w:lvl><w:lvl w:ilvl=\"2\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.%2.%3\"/></w:lvl></w:abstractNum>"
        + "<w:abstractNum w:abstractNumId=\"3\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerLetter\"/><w:lvlText w:val=\"%1)\"/></w:lvl></w:abstractNum>"
        + "<w:abstractNum w:abstractNumId=\"4\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"upperRoman\"/><w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum>"
        + "<w:abstractNum w:abstractNumId=\"5\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"upperLetter\"/><w:lvlText w:val=\"%1\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"upperLetter\"/><w:isLgl/><w:lvlText w:val=\"%1.%2\"/></w:lvl></w:abstractNum>"
        + "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>"
        + "<w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num>"
        + "<w:num w:numId=\"3\"><w:abstractNumId w:val=\"1\"/></w:num>"
        + "<w:num w:numId=\"4\"><w:abstractNumId w:val=\"2\"/></w:num>"
        + "<w:num w:numId=\"5\"><w:abstractNumId w:val=\"1\"/><w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"5\"/></w:lvlOverride></w:num>"
        + "<w:num w:numId=\"6\"><w:abstractNumId w:val=\"3\"/></w:num>"
        + "<w:num w:numId=\"7\"><w:abstractNumId w:val=\"4\"/></w:num>"
        + "<w:num w:numId=\"8\"><w:abstractNumId w:val=\"5\"/></w:num>"
        + "</w:numbering>";
}
