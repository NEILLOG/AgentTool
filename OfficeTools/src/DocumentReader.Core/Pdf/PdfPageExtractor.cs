using UglyToad.PdfPig.Actions;
using UglyToad.PdfPig.Annotations;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Graphics;

namespace DocumentReader.Core.Pdf;

/// <summary>用 PdfPig 取出一頁的字詞、框線、超連結與品質指標。</summary>
internal static class PdfPageExtractor
{
    private const double ThinRule = 2.5; // 比這細的矩形當成線
    private const double MinRuleLength = 6;

    public static PageContent Extract(Page page)
    {
        var letters = page.Letters;
        var garbage = letters.Count(l => l.Value.Any(PdfTextNormalizer.IsGarbage));
        var images = page.GetImages().Select(i => i.BoundingBox).ToList();
        var pageArea = Math.Max(1, page.Width * page.Height);
        var imageArea = images.Sum(b => Math.Max(0, b.Width) * Math.Max(0, b.Height));

        var words = new List<PdfWord>();
        foreach (var word in page.GetWords())
        {
            var text = PdfTextNormalizer.Normalize(WordText(word)).Trim();
            if (text.Length == 0)
            {
                continue;
            }

            var size = word.Letters.Count > 0 ? word.Letters.Average(l => l.PointSize) : word.BoundingBox.Height;
            size = size > 0 ? size : word.BoundingBox.Height;
            var bold = word.Letters.Count > 0 && word.Letters.Count(l => IsBoldFont(l.FontName)) * 2 > word.Letters.Count;

            Box box;
            double baseline;
            if (word.TextOrientation == TextOrientation.Horizontal && word.Letters.Count > 0)
            {
                baseline = PdfLines.Median(word.Letters.Select(l => l.StartBaseLine.Y));
                box = new Box(word.BoundingBox.Left, baseline - (0.2 * size), word.BoundingBox.Right, baseline + (0.8 * size));
            }
            else
            {
                box = ToBox(word.BoundingBox);
                baseline = box.Bottom;
            }

            words.Add(new PdfWord(text, box, size, bold, baseline));
        }

        AttachLinks(page, words);
        var (horizontal, vertical) = ExtractRules(page);
        return new PageContent(
            page.Number,
            page.Width,
            page.Height,
            letters.Count,
            images.Count,
            Math.Min(1, imageArea / pageArea),
            letters.Count == 0 ? 0 : (double)garbage / letters.Count,
            words,
            horizontal,
            vertical,
            page.Rotation.Value != 0);
    }

    /// <summary>
    /// 水平文字依字的位置由左而右重組：PdfPig 預設照內容串流順序，標點之類後畫的字會跑到最後（「報價單 -」被讀成「報價單-」）。
    /// </summary>
    private static string WordText(UglyToad.PdfPig.Content.Word word) =>
        word.TextOrientation == TextOrientation.Horizontal && word.Letters.Count > 1
            ? string.Concat(word.Letters.OrderBy(l => l.BoundingBox.Left).Select(l => l.Value))
            : word.Text;

    private static Box ToBox(PdfRectangle r) => new(r.Left, r.Bottom, r.Right, r.Top);

    private static bool IsBoldFont(string? name) =>
        name is not null && (name.Contains("Bold", StringComparison.OrdinalIgnoreCase) || name.Contains("Black", StringComparison.OrdinalIgnoreCase) || name.Contains("Heavy", StringComparison.OrdinalIgnoreCase));

    // ---- 超連結 ----

    private static void AttachLinks(Page page, List<PdfWord> words)
    {
        List<Annotation> annotations;
        try
        {
            annotations = page.GetAnnotations().Where(a => a.Type == AnnotationType.Link).ToList();
        }
        catch (Exception ex) when (ex is InvalidOperationException or ArgumentException or NotSupportedException or IndexOutOfRangeException)
        {
            return; // 註解壞掉不該讓整頁讀不出來
        }

        foreach (var annotation in annotations)
        {
            if (annotation.Action is not UriAction { Uri: { Length: > 0 } uri } || !IsSafeLink(uri))
            {
                continue;
            }

            var area = ToBox(annotation.Rectangle);
            foreach (var word in words.Where(w => w.Link is null && area.ContainsCenter(w.Box)))
            {
                word.Link = uri;
            }
        }
    }

    private static bool IsSafeLink(string uri) =>
        uri.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || uri.StartsWith("https://", StringComparison.OrdinalIgnoreCase) || uri.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase);

    // ---- 框線 ----

    /// <summary>
    /// 表格框線：細長的矩形（很多 PDF 用填色的細矩形畫線）、直線、以及有描邊的矩形四邊。
    /// 只收水平與垂直的線；整頁大小的矩形（背景、裁切範圍）忽略。
    /// </summary>
    private static (List<Rule> Horizontal, List<Rule> Vertical) ExtractRules(Page page)
    {
        var horizontal = new List<Rule>();
        var vertical = new List<Rule>();

        void AddH(double y, double x0, double x1)
        {
            if (Math.Abs(x1 - x0) >= MinRuleLength)
            {
                horizontal.Add(new Rule(y, Math.Min(x0, x1), Math.Max(x0, x1)));
            }
        }

        void AddV(double x, double y0, double y1)
        {
            if (Math.Abs(y1 - y0) >= MinRuleLength)
            {
                vertical.Add(new Rule(x, Math.Min(y0, y1), Math.Max(y0, y1)));
            }
        }

        foreach (var path in page.Paths)
        {
            foreach (var sub in path)
            {
                var points = Points(sub);
                if (points.Count < 2)
                {
                    continue;
                }

                var box = Box.Of(points.Select(p => new Box(p.X, p.Y, p.X, p.Y)));
                if (box.Width >= page.Width * 0.95 && box.Height >= page.Height * 0.95)
                {
                    continue;
                }

                var axisAligned = points.Zip(points.Skip(1), (a, b) => Math.Abs(a.X - b.X) < 0.01 || Math.Abs(a.Y - b.Y) < 0.01).All(x => x);
                if (!axisAligned)
                {
                    continue;
                }

                if (box.Height <= ThinRule && box.Width >= MinRuleLength)
                {
                    AddH(box.CenterY, box.Left, box.Right);
                }
                else if (box.Width <= ThinRule && box.Height >= MinRuleLength)
                {
                    AddV(box.CenterX, box.Bottom, box.Top);
                }
                else if (path.IsStroked && points.Count >= 4)
                {
                    // 有描邊的矩形：四邊都是框線
                    AddH(box.Bottom, box.Left, box.Right);
                    AddH(box.Top, box.Left, box.Right);
                    AddV(box.Left, box.Bottom, box.Top);
                    AddV(box.Right, box.Bottom, box.Top);
                }
            }
        }

        return (horizontal, vertical);
    }

    private static List<PdfPoint> Points(PdfSubpath sub)
    {
        var points = new List<PdfPoint>();
        foreach (var command in sub.Commands)
        {
            switch (command)
            {
                case PdfSubpath.Move move:
                    points.Add(move.Location);
                    break;
                case PdfSubpath.Line line:
                    if (points.Count == 0)
                    {
                        points.Add(line.From);
                    }

                    points.Add(line.To);
                    break;
                case PdfSubpath.BezierCurve:
                    return []; // 曲線不是框線
            }
        }

        return points;
    }
}
