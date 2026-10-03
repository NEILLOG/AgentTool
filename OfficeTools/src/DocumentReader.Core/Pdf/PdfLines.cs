using System.Text;
using DocumentReader.Core.Markdown;

namespace DocumentReader.Core.Pdf;

internal sealed record PdfLine(IReadOnlyList<PdfWord> Words, Box Box, double FontSize);

internal static class PdfLines
{
    /// <summary>
    /// 把字詞依基線分行：基線差在較小字級的 0.4 倍以內就是同一行（項目符號、標點的字形外框中心與文字不同，但基線相同）。
    /// 行內由左而右，行由上而下。
    /// </summary>
    public static List<PdfLine> Cluster(IEnumerable<PdfWord> words)
    {
        var lines = new List<List<PdfWord>>();
        var baselines = new List<double>();
        var sizes = new List<double>();
        foreach (var word in words.OrderByDescending(w => w.Baseline).ThenBy(w => w.Box.Left))
        {
            var index = -1;
            for (var i = lines.Count - 1; i >= 0 && i >= lines.Count - 3; i--)
            {
                if (Math.Abs(word.Baseline - baselines[i]) <= Math.Max(0.4 * Math.Min(sizes[i], word.FontSize), 0.5))
                {
                    index = i;
                    break;
                }
            }

            if (index < 0)
            {
                lines.Add([word]);
                baselines.Add(word.Baseline);
                sizes.Add(word.FontSize);
            }
            else
            {
                lines[index].Add(word);
            }
        }

        return lines
            .Select(l => l.OrderBy(w => w.Box.Left).ToList())
            .Select(l => new PdfLine(l, Box.Of(l.Select(w => w.Box)), Median(l.Select(w => w.FontSize))))
            .OrderByDescending(l => l.Words.Max(w => w.Baseline))
            .ToList();
    }

    public static double Median(IEnumerable<double> values)
    {
        var sorted = values.OrderBy(v => v).ToList();
        return sorted.Count == 0 ? 0 : sorted[sorted.Count / 2];
    }

    /// <summary>一行的文字。中日文之間不加空白（字詞是因為字型切換才被拆開的），其他情況加一個空白；連續的同一個連結包成一個 Markdown 連結。</summary>
    public static string Join(PdfLine line, bool inTable = false)
    {
        var sb = new StringBuilder();
        var words = line.Words;
        for (var i = 0; i < words.Count;)
        {
            var link = words[i].Link;
            var j = i + 1;
            while (link is not null && j < words.Count && words[j].Link == link)
            {
                j++;
            }

            var text = new StringBuilder();
            for (var k = i; k < j; k++)
            {
                if (k > i)
                {
                    text.Append(Separator(words[k - 1], words[k]));
                }

                text.Append(MarkdownWriter.Escape(words[k].Text, inTable));
            }

            if (sb.Length > 0)
            {
                sb.Append(Separator(words[i - 1], words[i]));
            }

            sb.Append(link is null ? text.ToString() : $"[{text.ToString().Replace("[", "\\[", StringComparison.Ordinal).Replace("]", "\\]", StringComparison.Ordinal)}]({EscapeUrl(link)})");
            i = j;
        }

        return sb.ToString();
    }

    public static string Separator(PdfWord left, PdfWord right)
    {
        var gap = right.Box.Left - left.Box.Right;
        var fontSize = Math.Max(left.FontSize, right.FontSize);
        return PdfTextNormalizer.IsCjk(left.Text[^1]) && PdfTextNormalizer.IsCjk(right.Text[0]) && gap < 0.6 * fontSize ? string.Empty : " ";
    }

    /// <summary>兩行接成同一段時的連接字串：任一邊是中日文就直接接；行尾是連字號（英文斷字）也直接接；否則加空白。</summary>
    public static string LineSeparator(string previous, string next)
    {
        if (previous.Length == 0 || next.Length == 0)
        {
            return string.Empty;
        }

        var cjk = PdfTextNormalizer.IsCjk(previous[^1]) || PdfTextNormalizer.IsCjk(next[0]);
        var hyphenated = previous[^1] == '-' && previous.Length > 1 && char.IsLetter(previous[^2]) && char.IsLetter(next[0]);
        return cjk || hyphenated ? string.Empty : " ";
    }

    public static string JoinLines(IEnumerable<string> lines)
    {
        var sb = new StringBuilder();
        foreach (var line in lines)
        {
            sb.Append(sb.Length == 0 ? string.Empty : LineSeparator(sb.ToString(), line)).Append(line);
        }

        return sb.ToString();
    }

    private static string EscapeUrl(string url) =>
        url.Replace(" ", "%20", StringComparison.Ordinal).Replace("(", "%28", StringComparison.Ordinal).Replace(")", "%29", StringComparison.Ordinal);
}
