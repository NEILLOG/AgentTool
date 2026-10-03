using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentReader.Core.Markdown;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace DocumentReader.Core.Ppt;

/// <summary>
/// 從圖表部件的快取值（c:strCache / c:numCache）取出標題與數據系列。
/// 快取是 PowerPoint 存檔時寫進去的結果，不需要內嵌的 xlsx，也不會因為外部連結失效而讀不到。
/// </summary>
internal static class PptChart
{
    private const int MaxRows = 100;

    private sealed record Series(string Name, List<string> Categories, List<string> Values);

    /// <returns>圖表的 Markdown 區塊（標題行，有數據時再加一張表）。</returns>
    public static IReadOnlyList<(bool IsTable, string Text)> Render(ChartPart chartPart)
    {
        var chart = chartPart.ChartSpace?.GetFirstChild<C.Chart>();
        var title = chart?.GetFirstChild<C.Title>() is { } t ? string.Concat(t.Descendants<A.Text>().Select(x => x.Text)).Trim() : string.Empty;
        var groups = chart?.PlotArea?.ChildElements.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal)).ToList() ?? [];
        var kind = groups.Count > 0 ? TypeName(groups) : null;

        var label = "[圖表" + (title.Length > 0 || kind is not null ? ": " : string.Empty)
            + MarkdownWriter.Escape(title) + (title.Length > 0 && kind is not null ? "（" + kind + "）" : kind is not null ? kind : string.Empty) + "]";
        var result = new List<(bool, string)> { (false, label) };

        var series = groups.SelectMany(g => g.ChildElements.Where(e => e.LocalName == "ser")).Select(ReadSeries).ToList();
        if (series.Count == 0 || series.All(s => s.Values.Count == 0))
        {
            return result;
        }

        var categories = series.Select(s => s.Categories).FirstOrDefault(c => c.Count > 0) ?? [];
        var rowCount = Math.Max(categories.Count, series.Max(s => s.Values.Count));
        var shown = Math.Min(rowCount, MaxRows);

        var rows = new List<List<string>> { new[] { string.Empty }.Concat(series.Select(s => MarkdownWriter.Escape(s.Name, inTable: true))).ToList() };
        for (var i = 0; i < shown; i++)
        {
            var row = new List<string> { i < categories.Count ? MarkdownWriter.Escape(categories[i], inTable: true) : (i + 1).ToString(CultureInfo.InvariantCulture) };
            row.AddRange(series.Select(s => i < s.Values.Count ? MarkdownWriter.Escape(s.Values[i], inTable: true) : string.Empty));
            rows.Add(row);
        }

        result.Add((true, MarkdownWriter.RenderTable(rows)));
        if (rowCount > shown)
        {
            result.Add((false, $"（圖表共 {rowCount} 筆資料，只列出前 {shown} 筆）"));
        }

        return result;
    }

    private static Series ReadSeries(OpenXmlElement ser)
    {
        var name = ser.ChildElements.FirstOrDefault(e => e.LocalName == "tx") is { } tx ? SeriesName(tx) : string.Empty;
        var categories = Points(ser.ChildElements.FirstOrDefault(e => e.LocalName is "cat" or "xVal")).ToList();
        var values = Points(ser.ChildElements.FirstOrDefault(e => e.LocalName is "val" or "yVal")).ToList();
        return new Series(name, categories, values);
    }

    private static string SeriesName(OpenXmlElement tx) =>
        Points(tx).FirstOrDefault(n => n.Length > 0) ?? string.Concat(tx.Descendants().Where(e => e.LocalName == "v").Select(v => v.InnerText)).Trim();

    /// <summary>依 idx 排列快取點的文字；多層分類（multiLvlStrRef）只取第一層。</summary>
    private static List<string> Points(OpenXmlElement? source)
    {
        var cache = source?.Descendants().FirstOrDefault(e => e.LocalName is "strCache" or "numCache" or "lvl");
        if (cache is null)
        {
            return [];
        }

        var count = cache.ChildElements.FirstOrDefault(e => e.LocalName == "ptCount") is { } c
            && int.TryParse(Attribute(c, "val"), NumberStyles.None, CultureInfo.InvariantCulture, out var n) ? n : 0;
        var byIndex = new Dictionary<int, string>();
        foreach (var pt in cache.ChildElements.Where(e => e.LocalName == "pt"))
        {
            if (int.TryParse(Attribute(pt, "idx"), NumberStyles.None, CultureInfo.InvariantCulture, out var idx))
            {
                byIndex[idx] = string.Concat(pt.ChildElements.Where(e => e.LocalName == "v").Select(v => v.InnerText)).Trim();
            }
        }

        var length = Math.Max(count, byIndex.Count == 0 ? 0 : byIndex.Keys.Max() + 1);
        return Enumerable.Range(0, length).Select(i => byIndex.TryGetValue(i, out var v) ? v : string.Empty).ToList();
    }

    private static string? Attribute(OpenXmlElement element, string localName) =>
        element.GetAttributes().Where(a => a.LocalName == localName).Select(a => a.Value).FirstOrDefault();

    private static string TypeName(List<OpenXmlElement> groups)
    {
        var names = groups.Select(Name).Distinct().ToList();
        return string.Join("＋", names);

        static string Name(OpenXmlElement g) => g.LocalName.ToLowerInvariant() switch
        {
            "barchart" or "bar3dchart" => g.ChildElements.FirstOrDefault(e => e.LocalName == "barDir")?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value == "bar" ? "橫條圖" : "直條圖",
            "linechart" or "line3dchart" => "折線圖",
            "piechart" or "pie3dchart" or "ofpiechart" => "圓餅圖",
            "doughnutchart" => "環圈圖",
            "areachart" or "area3dchart" => "區域圖",
            "scatterchart" => "散佈圖",
            "radarchart" => "雷達圖",
            "bubblechart" => "泡泡圖",
            "stockchart" => "股價圖",
            _ => g.LocalName,
        };
    }
}
