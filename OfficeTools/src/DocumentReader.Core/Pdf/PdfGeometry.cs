namespace DocumentReader.Core.Pdf;

/// <summary>PDF 座標（原點在左下角，Y 向上）的矩形。</summary>
internal readonly record struct Box(double Left, double Bottom, double Right, double Top)
{
    public double Width => Right - Left;

    public double Height => Top - Bottom;

    public double CenterX => (Left + Right) / 2;

    public double CenterY => (Bottom + Top) / 2;

    public bool ContainsCenter(Box other) => other.CenterX >= Left && other.CenterX < Right && other.CenterY >= Bottom && other.CenterY < Top;

    public Box Union(Box other) => new(Math.Min(Left, other.Left), Math.Min(Bottom, other.Bottom), Math.Max(Right, other.Right), Math.Max(Top, other.Top));

    public static Box Of(IEnumerable<Box> boxes) => boxes.Aggregate((a, b) => a.Union(b));
}

/// <summary>
/// 頁面上的一個字詞（已正規化文字）。<see cref="Box"/> 是依字級與基線估出來的「緊」外框（基線下 0.2 個字級到基線上 0.8 個字級），
/// 不是 PdfPig 的字形外框：CJK 字型的字形外框常比字級高很多，會讓緊鄰的兩行看起來重疊。
/// <see cref="Link"/> 是蓋在它上面的超連結。
/// </summary>
internal sealed record PdfWord(string Text, Box Box, double FontSize, bool Bold, double Baseline)
{
    public string? Link { get; set; }
}

/// <summary>水平框線（Y 固定、X 從 X0 到 X1）或垂直框線（X 固定、Y 從 Y0 到 Y1），Position 是固定的那個座標。</summary>
internal sealed record Rule(double Position, double Start, double End);

internal sealed record PageContent(
    int Number,
    double Width,
    double Height,
    int Letters,
    int Images,
    double ImageCoverage,
    double GarbageRatio,
    IReadOnlyList<PdfWord> Words,
    IReadOnlyList<Rule> HorizontalRules,
    IReadOnlyList<Rule> VerticalRules,
    bool Rotated);
