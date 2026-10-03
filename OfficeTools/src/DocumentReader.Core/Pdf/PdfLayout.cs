using System.Text;
using System.Text.RegularExpressions;
using DocumentReader.Core.Markdown;

namespace DocumentReader.Core.Pdf;

/// <summary>
/// 把一頁的字詞與表格排成 Markdown：先用 XY-cut 決定閱讀順序（找出整頁空白的水平帶與垂直帶遞迴切開，
/// 所以雙欄、側邊欄、說明框會一塊一塊讀），再把每一塊的行組成段落、清單、標題。
/// </summary>
internal static partial class PdfLayout
{
    private const double ParagraphGap = 0.5; // 垂直空白 ≥ 這個倍數的字級就切成不同區塊
    private const double ColumnGap = 1.5; // 水平空白 ≥ 這個倍數的字級就視為不同欄
    private const double SingleLineColumnGap = 6; // 單行文字內，水平空白 ≥ 這個倍數的字級才切成不同欄
    private const double FullLine = 2.5; // 行尾離版面右緣在這個倍數的字級內，視為「寫滿」（下一行是同一段的延續）

    private sealed record Element(Box Box, PdfWord? Word, PdfTable? Table);

    private enum BlockKind
    {
        Heading,
        Paragraph,
        ListItem,
        Table,
    }

    private sealed record Block(BlockKind Kind, string Text);

    public static string RenderPage(PageContent page, IReadOnlyList<PdfWord> words, double bodySize)
    {
        var tables = PdfTableFinder.Find(page with { Words = words });
        var claimed = new HashSet<PdfWord>(tables.SelectMany(t => t.Words), ReferenceEqualityComparer.Instance as IEqualityComparer<PdfWord>);

        var free = words.Where(w => !claimed.Contains(w)).ToList();
        var elements = free.Select(w => new Element(w.Box, w, null)).Concat(tables.Select(t => new Element(t.Box, null, t))).ToList();
        if (elements.Count == 0)
        {
            return string.Empty;
        }

        var fontSize = free.Count > 0 ? PdfLines.Median(free.Select(w => w.FontSize)) : bodySize;
        var leaves = new List<List<Element>>();
        Cut(elements, Math.Max(fontSize, 1), leaves);

        var blocks = new List<Block>();
        foreach (var leaf in leaves)
        {
            EmitLeaf(leaf, bodySize, blocks);
        }

        return Join(blocks);
    }

    // ---- XY-cut ----

    private static void Cut(List<Element> items, double fontSize, List<List<Element>> leaves)
    {
        if (items.Count <= 1)
        {
            leaves.Add(items);
            return;
        }

        // 水平切（上下分段）與垂直切（左右分欄）都可行時，先切空白較寬的那個：
        // 雙欄頁面欄間距比段落間距寬，會先分欄（一欄讀完再讀下一欄）；跨欄的標題因為沒有貫穿整頁的垂直空白，只能先水平切
        var rows = Split(items, byY: true, ParagraphGap * fontSize);
        var singleLine = Box.Of(items.Select(i => i.Box)).Height <= 1.6 * fontSize && items.All(i => i.Table is null);
        var columns = Split(items, byY: false, (singleLine ? SingleLineColumnGap : ColumnGap) * fontSize);

        var cutRows = rows.Groups.Count > 1;
        var cutColumns = columns.Groups.Count > 1;
        if (cutRows && cutColumns)
        {
            cutRows = rows.MaxGap >= columns.MaxGap;
            cutColumns = !cutRows;
        }

        if (cutRows || cutColumns)
        {
            foreach (var group in cutRows ? rows.Groups : columns.Groups)
            {
                Cut(group, fontSize, leaves);
            }

            return;
        }

        leaves.Add(items);
    }

    /// <summary>沿著一個方向找出空白帶，在每一條夠寬的空白帶切開；回傳依閱讀方向排好的群組（Y：由上而下；X：由左而右）。</summary>
    private static (List<List<Element>> Groups, double MaxGap) Split(List<Element> items, bool byY, double minGap)
    {
        var maxGap = 0.0;
        var sorted = byY ? items.OrderByDescending(i => i.Box.Top).ThenBy(i => i.Box.Left).ToList() : items.OrderBy(i => i.Box.Left).ThenByDescending(i => i.Box.Top).ToList();
        var groups = new List<List<Element>>();
        var edge = 0.0;
        foreach (var item in sorted)
        {
            if (groups.Count == 0)
            {
                groups.Add([item]);
                edge = byY ? item.Box.Bottom : item.Box.Right;
                continue;
            }

            var gap = byY ? edge - item.Box.Top : item.Box.Left - edge;
            if (gap >= minGap)
            {
                maxGap = Math.Max(maxGap, gap);
                groups.Add([item]);
                edge = byY ? item.Box.Bottom : item.Box.Right;
            }
            else
            {
                groups[^1].Add(item);
                edge = byY ? Math.Min(edge, item.Box.Bottom) : Math.Max(edge, item.Box.Right);
            }
        }

        return (groups, maxGap);
    }

    // ---- 區塊內容 ----

    private static void EmitLeaf(List<Element> leaf, double bodySize, List<Block> blocks)
    {
        var run = new List<PdfWord>();
        foreach (var element in leaf.OrderByDescending(e => e.Box.Top).ThenBy(e => e.Box.Left))
        {
            if (element.Word is { } word)
            {
                run.Add(word);
                continue;
            }

            Paragraphs(run, bodySize, blocks);
            run.Clear();
            var table = element.Table!;
            blocks.AddRange(table.Title.Where(t => t.Length > 0).Select(t => new Block(BlockKind.Paragraph, t)));
            blocks.Add(new Block(BlockKind.Table, MarkdownWriter.RenderTable(table.Rows.Select(r => r.ToList()).ToList())));
        }

        Paragraphs(run, bodySize, blocks);
    }

    private static void Paragraphs(List<PdfWord> words, double bodySize, List<Block> blocks)
    {
        if (words.Count == 0)
        {
            return;
        }

        var lines = PdfLines.Cluster(words);
        var right = lines.Max(l => l.Box.Right);

        List<PdfLine> current = [];
        var currentBullet = default((string Marker, bool Ordered)?);

        void Flush()
        {
            if (current.Count > 0)
            {
                blocks.Add(MakeBlock(current, currentBullet, bodySize));
            }

            current = [];
            currentBullet = null;
        }

        foreach (var original in lines)
        {
            var (line, bullet) = StripBullet(original);
            var fontSize = Math.Max(line.FontSize, 1);

            var startNew = current.Count == 0 || bullet is not null;
            if (!startNew)
            {
                var previous = current[^1];
                startNew = Math.Abs(fontSize - previous.FontSize) > 0.15 * previous.FontSize; // 垂直空白大的段落已在 XY-cut 切開，這裡只處理字級改變（緊貼的標題與內文）
            }

            if (startNew)
            {
                Flush();
                currentBullet = bullet;
            }

            current.Add(line);
        }

        Flush();

        Block MakeBlock(List<PdfLine> group, (string Marker, bool Ordered)? bullet, double body)
        {
            // 行尾沒寫滿的行後面保留換行（硬換行），寫滿的行與下一行接成同一段
            var text = new StringBuilder();
            var lineTexts = group.Select(l => PdfLines.Join(l)).ToList();
            for (var i = 0; i < group.Count; i++)
            {
                var lineText = lineTexts[i];
                if (i > 0)
                {
                    var previous = group[i - 1];
                    var full = previous.Box.Right >= right - (FullLine * Math.Max(previous.FontSize, 1)) && !EndsWithLabel(lineTexts[i - 1]);
                    text.Append(full ? PdfLines.LineSeparator(text.ToString(), lineText) : "  \n");
                }

                text.Append(lineText);
            }

            var content = MarkdownWriter.EscapeLineStarts(text.ToString());
            if (bullet is { } b)
            {
                return new Block(BlockKind.ListItem, b.Marker + content.Replace("\n", "\n  ", StringComparison.Ordinal));
            }

            var size = group.Max(l => l.FontSize);
            if (group.Count == 1 && body > 0 && size >= 1.3 * body && text.Length <= 80)
            {
                var level = size >= 2.0 * body ? 3 : size >= 1.6 * body ? 4 : 5;
                return new Block(BlockKind.Heading, new string('#', level) + " " + content);
            }

            return new Block(BlockKind.Paragraph, content);
        }
    }

    /// <summary>行尾是冒號：這一行是「標籤」，後面不是同一句的延續。</summary>
    private static bool EndsWithLabel(string line) => line.TrimEnd().EndsWith(':') || line.TrimEnd().EndsWith('：');

    /// <summary>行首的項目符號（• · ● 等）或「1.」「2)」編號：移出文字，回傳清單標記。</summary>
    private static (PdfLine Line, (string Marker, bool Ordered)? Bullet) StripBullet(PdfLine line)
    {
        var first = line.Words[0];
        string? marker = null;
        var rest = first.Text;
        var ordered = false;

        if (BulletChars.Contains(first.Text[0]))
        {
            marker = "- ";
            rest = first.Text[1..];
        }
        else if (Numbered().Match(first.Text) is { Success: true } m)
        {
            marker = m.Groups[1].Value + ". ";
            rest = m.Groups[2].Value;
            ordered = true;
        }

        if (marker is null)
        {
            return (line, null);
        }

        var words = line.Words.ToList();
        if (rest.Trim().Length == 0)
        {
            words.RemoveAt(0);
        }
        else
        {
            words[0] = first with { Text = rest.TrimStart() };
        }

        if (words.Count == 0)
        {
            return (line, null);
        }

        return (new PdfLine(words, Box.Of(words.Select(w => w.Box)), line.FontSize), (marker, ordered));
    }

    private const string BulletChars = "•·●▪■◦○◆▶►‣⁃";

    // "1." "2)" "3、" 單獨成字，或後面直接接非數字的文字（避免把 "3.5"、"2026/5" 當成編號）
    [GeneratedRegex(@"^(\d{1,3})[.)、](?!\d)(.*)$", RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 1000)]
    private static partial Regex Numbered();

    private static string Join(List<Block> blocks)
    {
        var sb = new StringBuilder();
        for (var i = 0; i < blocks.Count; i++)
        {
            if (i > 0)
            {
                sb.Append(blocks[i].Kind == BlockKind.ListItem && blocks[i - 1].Kind == BlockKind.ListItem ? "\n" : "\n\n");
            }

            sb.Append(blocks[i].Text);
        }

        return sb.ToString();
    }
}
