namespace DocumentReader.Core.Pdf;

/// <param name="Title">表格最上面橫跨整個寬度的列（通常是表格標題），不放進表格本體。</param>
/// <param name="Words">被這張表格的儲存格認領的字詞（排版時要從一般文字中拿掉）。</param>
internal sealed record PdfTable(Box Box, IReadOnlyList<string> Title, IReadOnlyList<IReadOnlyList<string>> Rows, IReadOnlyList<PdfWord> Words);

/// <summary>
/// 用框線找表格：水平與垂直框線互相接觸的一群就是一張表格，框線圍出儲存格，字詞依中心點歸入儲存格。
/// 橫向合併的儲存格（該列沒有中間的垂直框線）把值重複填進每一欄，但橫跨整個表格寬度的列（分組標題）只放在第一格；
/// 最上面橫跨整個寬度的列當成表格標題，放在表格外。直向合併目前不處理（值只出現在它所在的那一列）。
/// 沒有框線的表格找不到，會當成一般文字排版。
/// </summary>
internal static class PdfTableFinder
{
    private const double Tolerance = 2.0;
    private const double MinRowHeight = 2.5;

    public static List<PdfTable> Find(PageContent page)
    {
        var horizontal = Merge(page.HorizontalRules);
        var vertical = Merge(page.VerticalRules);
        if (horizontal.Count < 2 || vertical.Count < 2)
        {
            return [];
        }

        var tables = new List<PdfTable>();
        foreach (var (h, v) in Components(horizontal, vertical))
        {
            var table = Build(page, h, v);
            if (table is not null)
            {
                tables.Add(table);
            }
        }

        return tables;
    }

    // ---- 框線整理 ----

    /// <summary>位置相同（容許誤差內）的線合併成一條，首尾相接或重疊的線段接成一段。</summary>
    private static List<Rule> Merge(IReadOnlyList<Rule> rules)
    {
        var merged = new List<Rule>();
        foreach (var group in GroupByPosition(rules))
        {
            var position = group.Average(r => r.Position);
            Rule? current = null;
            foreach (var rule in group.OrderBy(r => r.Start))
            {
                if (current is not null && rule.Start <= current.End + Tolerance)
                {
                    current = current with { End = Math.Max(current.End, rule.End) };
                }
                else
                {
                    if (current is not null)
                    {
                        merged.Add(current);
                    }

                    current = new Rule(position, rule.Start, rule.End);
                }
            }

            if (current is not null)
            {
                merged.Add(current);
            }
        }

        return merged;
    }

    private static List<List<Rule>> GroupByPosition(IReadOnlyList<Rule> rules)
    {
        var groups = new List<List<Rule>>();
        foreach (var rule in rules.OrderBy(r => r.Position))
        {
            if (groups.Count > 0 && rule.Position - groups[^1][^1].Position <= 1.5)
            {
                groups[^1].Add(rule);
            }
            else
            {
                groups.Add([rule]);
            }
        }

        return groups;
    }

    /// <summary>互相接觸的水平 / 垂直框線分成一群一群。</summary>
    private static List<(List<Rule> H, List<Rule> V)> Components(List<Rule> horizontal, List<Rule> vertical)
    {
        var parent = Enumerable.Range(0, horizontal.Count + vertical.Count).ToArray();
        int Find(int x) => parent[x] == x ? x : parent[x] = Find(parent[x]);

        for (var i = 0; i < horizontal.Count; i++)
        {
            for (var j = 0; j < vertical.Count; j++)
            {
                var h = horizontal[i];
                var v = vertical[j];
                if (v.Position >= h.Start - Tolerance && v.Position <= h.End + Tolerance && h.Position >= v.Start - Tolerance && h.Position <= v.End + Tolerance)
                {
                    parent[Find(i)] = Find(horizontal.Count + j);
                }
            }
        }

        var result = new Dictionary<int, (List<Rule> H, List<Rule> V)>();
        for (var i = 0; i < horizontal.Count; i++)
        {
            var root = Find(i);
            if (!result.TryGetValue(root, out var entry))
            {
                result[root] = entry = ([], []);
            }

            entry.H.Add(horizontal[i]);
        }

        for (var j = 0; j < vertical.Count; j++)
        {
            var root = Find(horizontal.Count + j);
            if (!result.TryGetValue(root, out var entry))
            {
                result[root] = entry = ([], []);
            }

            entry.V.Add(vertical[j]);
        }

        return result.Values.Where(c => c.H.Count >= 2 && c.V.Count >= 2).ToList();
    }

    // ---- 建表 ----

    private static PdfTable? Build(PageContent page, List<Rule> horizontal, List<Rule> vertical)
    {
        var ys = Distinct(horizontal.Select(h => h.Position)).OrderByDescending(y => y).ToList();
        if (ys.Count < 2)
        {
            return null;
        }

        var bounds = new Box(vertical.Min(v => v.Position), ys[^1], vertical.Max(v => v.Position), ys[0]);
        var inside = page.Words.Where(w => Inflate(bounds, 1).ContainsCenter(w.Box)).ToList();

        // 每一列（兩條水平框線之間）的儲存格邊界：只用涵蓋這一列的垂直框線
        var bands = new List<(double Top, double Bottom, List<double> Boundaries)>();
        for (var i = 0; i + 1 < ys.Count; i++)
        {
            double top = ys[i], bottom = ys[i + 1];
            if (top - bottom < MinRowHeight)
            {
                continue;
            }

            var covering = vertical.Where(v => Math.Min(v.End, top) - Math.Max(v.Start, bottom) >= 0.6 * (top - bottom)).Select(v => v.Position);
            var boundaries = Distinct(covering).OrderBy(x => x).ToList();
            if (boundaries.Count >= 2)
            {
                bands.Add((top, bottom, boundaries));
            }
        }

        var columns = Distinct(bands.SelectMany(b => b.Boundaries)).OrderBy(x => x).ToList();
        if (columns.Count < 3 || bands.Count == 0)
        {
            return null; // 少於兩欄：只是個框，不是表格
        }

        var rows = new List<IReadOnlyList<string>>();
        var title = new List<string>();
        var claimed = new List<PdfWord>();
        foreach (var (top, bottom, boundaries) in bands)
        {
            var row = new string[columns.Count - 1];
            for (var c = 0; c + 1 < boundaries.Count; c++)
            {
                var cell = new Box(boundaries[c], bottom, boundaries[c + 1], top);
                var words = inside.Where(w => cell.ContainsCenter(w.Box)).ToList();
                claimed.AddRange(words);
                var text = PdfLines.JoinLines(PdfLines.Cluster(words).Select(l => PdfLines.Join(l, inTable: true))).Trim();

                // 把這個儲存格涵蓋的每一個全域欄位都填上同樣的值（橫向合併）
                for (var k = 0; k + 1 < columns.Count; k++)
                {
                    var center = (columns[k] + columns[k + 1]) / 2;
                    if (center >= boundaries[c] && center < boundaries[c + 1])
                    {
                        row[k] = text;
                    }
                }
            }

            if (row.All(string.IsNullOrEmpty))
            {
                continue;
            }

            if (boundaries.Count == 2)
            {
                // 橫跨整個寬度：分組標題。最上面的當表格標題，其餘只放第一格
                if (rows.Count == 0)
                {
                    title.Add(row[0]);
                    continue;
                }

                rows.Add(row.Select((x, k) => k == 0 ? x : string.Empty).ToList());
                continue;
            }

            rows.Add(row.Select(x => x ?? string.Empty).ToList());
        }

        return rows.Count == 0 ? null : new PdfTable(bounds, title, rows, claimed);
    }

    private static List<double> Distinct(IEnumerable<double> values)
    {
        var result = new List<double>();
        foreach (var value in values.OrderBy(v => v))
        {
            if (result.Count == 0 || value - result[^1] > Tolerance)
            {
                result.Add(value);
            }
        }

        return result;
    }

    private static Box Inflate(Box box, double by) => new(box.Left - by, box.Bottom - by, box.Right + by, box.Top + by);
}
