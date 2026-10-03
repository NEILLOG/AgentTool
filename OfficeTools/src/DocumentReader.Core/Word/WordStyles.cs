using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentReader.Core.Word;

internal readonly record struct RunFormat(bool Bold, bool Italic, bool Strike);

/// <summary>樣式定義：判斷標題層級、段落樣式帶的編號、字元樣式帶的粗體 / 斜體。</summary>
internal sealed partial class WordStyles
{
    private const int MaxChainDepth = 20;

    private sealed record StyleInfo(
        string Id,
        string? Name,
        string? BasedOn,
        int? OutlineLevel,
        int? NumId,
        int? Ilvl,
        bool? Bold,
        bool? Italic,
        bool? Strike);

    private readonly Dictionary<string, StyleInfo> _styles = new(StringComparer.Ordinal);

    // 內建標題樣式的「名稱」是 heading 1（中文版 Word 的樣式 ID 是 "1"、"2"，所以不能用 ID 判斷）；也接受別的工具產生的「標題 1」
    [GeneratedRegex(@"^(heading|標題|标题)\s*([1-9])$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 1000)]
    private static partial Regex HeadingName();

    [GeneratedRegex(@"^(toc|目錄|目录)(\s*\d)?(heading)?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 1000)]
    private static partial Regex TocName();

    public static WordStyles Load(MainDocumentPart main)
    {
        var result = new WordStyles();
        var styles = main.StyleDefinitionsPart?.Styles;
        if (styles is null)
        {
            return result;
        }

        foreach (var style in styles.Elements<Style>())
        {
            var id = style.StyleId?.Value;
            if (string.IsNullOrEmpty(id))
            {
                continue;
            }

            var pPr = style.StyleParagraphProperties;
            var rPr = style.StyleRunProperties;
            result._styles[id] = new StyleInfo(
                id,
                style.StyleName?.Val?.Value,
                style.BasedOn?.Val?.Value,
                pPr?.OutlineLevel?.Val?.Value,
                pPr?.NumberingProperties?.NumberingId?.Val?.Value,
                pPr?.NumberingProperties?.NumberingLevelReference?.Val?.Value,
                OnOff(rPr?.Bold),
                OnOff(rPr?.Italic),
                OnOff(rPr?.Strike));
        }

        return result;
    }

    /// <summary>
    /// 標題層級（1 到 9），不是標題回傳 null。依序看：段落直接設定的大綱層級、樣式（含 basedOn 鏈）的大綱層級、
    /// 樣式名稱是 heading N / 標題 N（或 Title，視為 1）。大綱層級 10（內文）明確表示不是標題。
    /// </summary>
    public int? HeadingLevel(string? styleId, int? directOutlineLevel)
    {
        if (directOutlineLevel is { } direct)
        {
            return direct is >= 0 and <= 8 ? direct + 1 : null;
        }

        foreach (var style in Chain(styleId))
        {
            if (style.OutlineLevel is { } level)
            {
                return level is >= 0 and <= 8 ? level + 1 : null;
            }

            if (style.Name is { } name)
            {
                var match = HeadingName().Match(name.Trim());
                if (match.Success)
                {
                    return int.Parse(match.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
                }

                if (name.Trim().Equals("Title", StringComparison.OrdinalIgnoreCase))
                {
                    return 1;
                }
            }
        }

        return null;
    }

    /// <summary>目錄樣式（TOC 1、目錄標題…）。目錄只是標題的重複加上頁碼，對理解內文沒有幫助。</summary>
    public bool IsToc(string? styleId) =>
        Chain(styleId).Any(s => TocName().IsMatch((s.Name ?? s.Id).Trim().Replace(" ", string.Empty, StringComparison.Ordinal)));

    /// <summary>段落的編號（numId, ilvl）：直接設定優先，沒有就看樣式；numId 為 0 表示明確取消編號。</summary>
    public (int NumId, int Ilvl)? Numbering(string? styleId, NumberingProperties? direct)
    {
        var numId = direct?.NumberingId?.Val?.Value;
        var ilvl = direct?.NumberingLevelReference?.Val?.Value;

        foreach (var style in Chain(styleId))
        {
            numId ??= style.NumId;
            ilvl ??= style.Ilvl;
        }

        return numId is > 0 ? (numId.Value, ilvl ?? 0) : null;
    }

    /// <summary>run 的粗體 / 斜體 / 刪除線：字元樣式（含 basedOn 鏈）再被直接設定覆蓋。</summary>
    public RunFormat Format(RunProperties? direct, string? runStyleId)
    {
        bool? bold = null, italic = null, strike = null;
        foreach (var style in Chain(runStyleId).Reverse()) // 由最基底的樣式往下套，越接近 run 的越優先
        {
            bold = style.Bold ?? bold;
            italic = style.Italic ?? italic;
            strike = style.Strike ?? strike;
        }

        bold = OnOff(direct?.Bold) ?? bold;
        italic = OnOff(direct?.Italic) ?? italic;
        strike = OnOff(direct?.Strike) ?? OnOff(direct?.DoubleStrike) ?? strike;
        return new RunFormat(bold ?? false, italic ?? false, strike ?? false);
    }

    private IEnumerable<StyleInfo> Chain(string? styleId)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (var id = styleId; id is not null && seen.Count < MaxChainDepth && seen.Add(id) && _styles.TryGetValue(id, out var style); id = style.BasedOn)
        {
            yield return style;
        }
    }

    private static bool? OnOff(OnOffType? element) =>
        element is null ? null : element.Val is null ? true : element.Val.Value;
}
