using System.IO.Compression;
using System.Security;
using System.Text;

namespace DocumentReader.Core.Tests;

/// <summary>用原始 XML 組出 .pptx，才能精確控制形狀位置、預留位置、備註、圖表等結構。</summary>
internal sealed class PptxBuilder
{
    private const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private const string C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    private sealed class SlideDef
    {
        public string Shapes { get; init; } = string.Empty;

        public bool Hidden { get; init; }

        public string? Notes { get; init; }

        public List<(string Id, string Url)> Links { get; } = [];

        public List<(string Id, string Xml)> Charts { get; } = [];
    }

    private readonly List<SlideDef> _slides = [];
    private readonly List<(string Path, string Content)> _extraEntries = [];

    public long SlideHeight { get; init; } = 5_143_500;

    /// <summary>sldIdLst 的順序（元素是投影片檔案編號，從 1 開始）；null 為依序。用來驗證順序以簡報為準、不是檔名。</summary>
    public int[]? SlideOrder { get; init; }

    // ---- 建構 ----

    /// <summary>加一張投影片；<paramref name="links"/> 是這張投影片用到的超連結（rId → 網址），<paramref name="charts"/> 是圖表部件（rId → chartSpace XML）。</summary>
    public PptxBuilder Slide(string shapes, bool hidden = false, string? notes = null, (string Id, string Url)[]? links = null, (string Id, string Xml)[]? charts = null)
    {
        var slide = new SlideDef { Shapes = shapes, Hidden = hidden, Notes = notes };
        slide.Links.AddRange(links ?? []);
        slide.Charts.AddRange(charts ?? []);
        _slides.Add(slide);
        return this;
    }

    public PptxBuilder Entry(string path, string content)
    {
        _extraEntries.Add((path, content));
        return this;
    }

    public static string Esc(string text) => SecurityElement.Escape(text) ?? string.Empty;

    /// <summary>形狀。<paramref name="ph"/> 例如 "type=\"title\"" 或 "idx=\"1\""。</summary>
    public static string Sp(int id, string paragraphs, string? ph = null, long? x = 0, long y = 0, long cx = 1000000, long cy = 500000, string? extraCNvPr = null, bool txBox = false) =>
        $"<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"Shape {id}\" {extraCNvPr}/><p:cNvSpPr{(txBox ? " txBox=\"1\"" : string.Empty)}/><p:nvPr>{(ph is null ? string.Empty : $"<p:ph {ph}/>")}</p:nvPr></p:nvSpPr>"
        + (x is null ? "<p:spPr/>" : $"<p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm></p:spPr>")
        + $"<p:txBody><a:bodyPr/><a:lstStyle/>{paragraphs}</p:txBody></p:sp>";

    public static string Title(string text, int id = 2) => Sp(id, Para(text), ph: "type=\"title\"", x: 0, y: 0);

    public static string Para(string text) => $"<a:p><a:r><a:rPr lang=\"en\"/><a:t>{Esc(text)}</a:t></a:r></a:p>";

    /// <summary>由多個 run 組成的段落（用 <see cref="Run"/> 產生 run）。</summary>
    public static string Paragraph(params string[] runs) => $"<a:p>{string.Concat(runs)}</a:p>";

    /// <summary>帶段落屬性的段落：<paramref name="pPrXml"/> 是完整的 a:pPr 元素。</summary>
    public static string ParaPr(string pPrXml, string text) => $"<a:p>{pPrXml}<a:r><a:t>{Esc(text)}</a:t></a:r></a:p>";

    public static string Bullet(string text, int level = 0) => ParaPr($"<a:pPr lvl=\"{level}\"><a:buChar char=\"•\"/></a:pPr>", text);

    public static string Numbered(string text, int level = 0) => ParaPr($"<a:pPr lvl=\"{level}\"><a:buAutoNum type=\"arabicPeriod\"/></a:pPr>", text);

    public static string Run(string text, string rPrAttributes = "", string rPrChildren = "") =>
        $"<a:r><a:rPr lang=\"en\" {rPrAttributes}>{rPrChildren}</a:rPr><a:t>{Esc(text)}</a:t></a:r>";

    public static string Pic(int id, string? descr, long x = 0, long y = 0, bool hidden = false) =>
        $"<p:pic><p:nvPicPr><p:cNvPr id=\"{id}\" name=\"Picture {id}\"{(descr is null ? string.Empty : $" descr=\"{Esc(descr)}\"")}{(hidden ? " hidden=\"1\"" : string.Empty)}/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>"
        + $"<p:blipFill><a:blip/></p:blipFill><p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"100\" cy=\"100\"/></a:xfrm></p:spPr></p:pic>";

    public static string Grp(int id, string children, long x = 0, long y = 0) =>
        $"<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"{id}\" name=\"Group {id}\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
        + $"<p:grpSpPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"1000\" cy=\"1000\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"1000\" cy=\"1000\"/></a:xfrm></p:grpSpPr>{children}</p:grpSp>";

    public static string Frame(int id, string graphicDataInner, string uri, long x = 0, long y = 0) =>
        $"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"{id}\" name=\"Frame {id}\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
        + $"<p:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"1000\" cy=\"1000\"/></p:xfrm><a:graphic><a:graphicData uri=\"{uri}\">{graphicDataInner}</a:graphicData></a:graphic></p:graphicFrame>";

    /// <summary>表格。儲存格內容是 a:tc 的完整 XML（用 <see cref="Cell"/> 產生）。</summary>
    public static string Table(int id, long x, long y, params string[][] rows) =>
        Frame(id, $"<a:tbl><a:tblGrid/>{string.Concat(rows.Select(r => $"<a:tr h=\"100\">{string.Concat(r)}</a:tr>"))}</a:tbl>", "http://schemas.openxmlformats.org/drawingml/2006/table", x, y);

    public static string Cell(string text, string attributes = "") =>
        $"<a:tc {attributes}><a:txBody><a:bodyPr/><a:lstStyle/>{(text.Length > 0 ? Para(text) : "<a:p/>")}</a:txBody><a:tcPr/></a:tc>";

    public static string ChartFrame(int id, string rId, long x = 0, long y = 0) =>
        Frame(id, $"<c:chart xmlns:c=\"{C}\" xmlns:r=\"{R}\" r:id=\"{rId}\"/>", C, x, y);

    /// <summary>圖表部件。<paramref name="series"/>：每個系列的名稱與數值；<paramref name="categories"/> 是共用的分類。</summary>
    public static string ChartXml(string? title, string[] categories, (string Name, double?[] Values)[] series, string kind = "barChart", string extraInGroup = "") =>
        $"<c:chartSpace xmlns:c=\"{C}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\"><c:chart>"
        + (title is null ? string.Empty : $"<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>{Esc(title)}</a:t></a:r></a:p></c:rich></c:tx></c:title>")
        + $"<c:plotArea><c:{kind}>{extraInGroup}"
        + string.Concat(series.Select((s, i) =>
            $"<c:ser><c:idx val=\"{i}\"/><c:order val=\"{i}\"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{Esc(s.Name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>"
            + $"<c:cat><c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache><c:ptCount val=\"{categories.Length}\"/>{string.Concat(categories.Select((c, j) => $"<c:pt idx=\"{j}\"><c:v>{Esc(c)}</c:v></c:pt>"))}</c:strCache></c:strRef></c:cat>"
            + $"<c:val><c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"{s.Values.Length}\"/>{string.Concat(s.Values.Select((v, j) => v is null ? string.Empty : $"<c:pt idx=\"{j}\"><c:v>{v.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}</c:v></c:pt>"))}</c:numCache></c:numRef></c:val></c:ser>"))
        + $"</c:{kind}></c:plotArea></c:chart></c:chartSpace>";

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

            var overrides = new StringBuilder("<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>");
            var presentationRels = new StringBuilder();
            var sldIds = new StringBuilder();
            var chartNumber = 0;

            for (var i = 0; i < _slides.Count; i++)
            {
                var n = i + 1;
                var slide = _slides[i];
                overrides.Append($"<Override PartName=\"/ppt/slides/slide{n}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>");
                presentationRels.Append($"<Relationship Id=\"rId{n}\" Type=\"{R}/slide\" Target=\"slides/slide{n}.xml\"/>");

                var rels = new StringBuilder();
                foreach (var (id, url) in slide.Links)
                {
                    rels.Append($"<Relationship Id=\"{id}\" Type=\"{R}/hyperlink\" Target=\"{Esc(url)}\" TargetMode=\"External\"/>");
                }

                foreach (var (id, xml) in slide.Charts)
                {
                    chartNumber++;
                    overrides.Append($"<Override PartName=\"/ppt/charts/chart{chartNumber}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>");
                    rels.Append($"<Relationship Id=\"{id}\" Type=\"{R}/chart\" Target=\"../charts/chart{chartNumber}.xml\"/>");
                    Entry($"ppt/charts/chart{chartNumber}.xml", xml);
                }

                if (slide.Notes is not null)
                {
                    overrides.Append($"<Override PartName=\"/ppt/notesSlides/notesSlide{n}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml\"/>");
                    rels.Append($"<Relationship Id=\"rIdNotes\" Type=\"{R}/notesSlide\" Target=\"../notesSlides/notesSlide{n}.xml\"/>");
                    Entry(
                        $"ppt/notesSlides/notesSlide{n}.xml",
                        $"<p:notes xmlns:a=\"{A}\" xmlns:r=\"{R}\" xmlns:p=\"{P}\"><p:cSld><p:spTree>{GroupHeader}"
                        + "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Slide Image\"/><p:cNvSpPr/><p:nvPr><p:ph type=\"sldImg\" idx=\"2\"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>"
                        + $"<p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Notes\"/><p:cNvSpPr/><p:nvPr><p:ph type=\"body\" idx=\"1\"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{string.Concat(slide.Notes.Split('\n').Select(Para))}</p:txBody></p:sp>"
                        + "<p:sp><p:nvSpPr><p:cNvPr id=\"4\" name=\"Number\"/><p:cNvSpPr/><p:nvPr><p:ph type=\"sldNum\" idx=\"5\"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id=\"{00000000-0000-0000-0000-000000000000}\" type=\"slidenum\"><a:t>7</a:t></a:fld></a:p></p:txBody></p:sp>"
                        + "</p:spTree></p:cSld></p:notes>");
                }

                Entry($"ppt/slides/_rels/slide{n}.xml.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{rels}</Relationships>");
                Entry(
                    $"ppt/slides/slide{n}.xml",
                    $"<p:sld xmlns:a=\"{A}\" xmlns:r=\"{R}\" xmlns:p=\"{P}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"{(slide.Hidden ? " show=\"0\"" : string.Empty)}><p:cSld><p:spTree>{GroupHeader}{slide.Shapes}</p:spTree></p:cSld></p:sld>");
            }

            foreach (var n in SlideOrder ?? Enumerable.Range(1, _slides.Count).ToArray())
            {
                sldIds.Append($"<p:sldId id=\"{255 + n}\" r:id=\"rId{n}\"/>");
            }

            foreach (var (path, content) in _extraEntries)
            {
                Entry(path, content);
            }

            Entry(
                "[Content_Types].xml",
                $"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>{overrides}</Types>");
            Entry("_rels/.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{R}/officeDocument\" Target=\"ppt/presentation.xml\"/></Relationships>");
            Entry("ppt/_rels/presentation.xml.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{presentationRels}</Relationships>");
            Entry("ppt/presentation.xml", $"<p:presentation xmlns:a=\"{A}\" xmlns:r=\"{R}\" xmlns:p=\"{P}\"><p:sldIdLst>{sldIds}</p:sldIdLst><p:sldSz cx=\"9144000\" cy=\"{SlideHeight}\"/></p:presentation>");
        }

        return ms.ToArray();
    }

    public string Save(string path)
    {
        File.WriteAllBytes(path, Build());
        return path;
    }

    private const string GroupHeader = "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>";
}
