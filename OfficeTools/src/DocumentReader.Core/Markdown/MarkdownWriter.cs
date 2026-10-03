using System.Text;
using System.Text.RegularExpressions;

namespace DocumentReader.Core.Markdown;

internal enum BlockKind
{
    Heading,
    Paragraph,
    ListItem,
    Table,
}

/// <param name="Start">這一節（含標題行）在全文中的起點。</param>
/// <param name="OwnEnd">不含子節的終點（不含）。</param>
/// <param name="TotalEnd">含子節的終點（不含）。</param>
internal sealed record SectionInfo(string Id, int Level, string Title, int Start, int OwnEnd, int TotalEnd, string? ParentId);

internal sealed record ParsedDocument(string Text, IReadOnlyList<SectionInfo> Sections, IReadOnlyList<string> Notes);

internal readonly record struct InlinePiece(string Text, bool Bold = false, bool Italic = false, bool Strike = false, string? Link = null);

internal static partial class MarkdownWriter
{
    /// <summary>跳脫會被 Markdown 當成格式的字元；底線與中括號不跳脫，因為識別字與 [1] 這類文字很常見，跳脫只會增加雜訊。</summary>
    public static string Escape(string text, bool inTable = false)
    {
        var sb = new StringBuilder(text.Length + 8);
        foreach (var c in text)
        {
            if (c is '\\' or '*' or '`' || (inTable && c == '|'))
            {
                sb.Append('\\');
            }

            sb.Append(c);
        }

        return sb.ToString();
    }

    /// <summary>行首看起來像標題、清單、引用的文字加反斜線，避免被當成區塊語法。</summary>
    public static string EscapeLineStarts(string text)
    {
        if (text.Length == 0)
        {
            return text;
        }

        var lines = text.Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            var line = OrderedListStart().Replace(lines[i], "$1\\$2$3");
            lines[i] = BlockStart().Replace(line, m => m.Groups["lead"].Value + "\\" + m.Groups["mark"].Value + m.Groups["rest"].Value);
        }

        return string.Join('\n', lines);
    }

    // 有序清單：數字後面的 . 或 ) 要跳脫（"1. x" → "1\. x"），不是跳脫數字
    [GeneratedRegex(@"^(\s*\d+)([.)])(\s|$)", RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 1000)]
    private static partial Regex OrderedListStart();

    // 標題 (#)、無序清單 (- + )、引用 (>)、分隔線 / 設定式標題（行首連續的 - 或 =）
    [GeneratedRegex(@"^(?<lead>\s*)(?<mark>#{1,6}(?=\s)|={3,}|-{3,}|[-+>])(?<rest>.*)$", RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 1000)]
    private static partial Regex BlockStart();

    /// <summary>把行內片段組成一行 Markdown：合併相同格式的相鄰片段、空白移到強調標記之外、連結包成 [文字](網址)。</summary>
    /// <param name="lineBreak">片段中的換行（"\n"）要換成什麼：段落用 Markdown 的硬換行，表格儲存格用 &lt;br&gt;。</param>
    public static string RenderInline(IReadOnlyList<InlinePiece> pieces, bool emphasis, string lineBreak, bool inTable)
    {
        var merged = Merge(pieces);
        var sb = new StringBuilder();
        for (var i = 0; i < merged.Count;)
        {
            var link = merged[i].Link;
            var j = i;
            while (j < merged.Count && merged[j].Link == link)
            {
                j++;
            }

            var group = merged.GetRange(i, j - i);
            if (link is null)
            {
                sb.Append(RenderRuns(group, emphasis, lineBreak, inTable));
            }
            else
            {
                var text = RenderRuns(group, emphasis, lineBreak, inTable);
                // 連結文字裡的中括號要成對或跳脫，否則會被誤判成連結語法
                var escaped = text.Replace("[", "\\[", StringComparison.Ordinal).Replace("]", "\\]", StringComparison.Ordinal);
                sb.Append(text.Trim().Length == 0 ? text : $"[{escaped}]({EscapeUrl(link)})");
            }

            i = j;
        }

        return sb.ToString();
    }

    private static List<InlinePiece> Merge(IReadOnlyList<InlinePiece> pieces)
    {
        var result = new List<InlinePiece>();
        foreach (var piece in pieces)
        {
            if (piece.Text.Length == 0)
            {
                continue;
            }

            if (result.Count > 0 && Same(result[^1], piece))
            {
                result[^1] = result[^1] with { Text = result[^1].Text + piece.Text };
            }
            else
            {
                result.Add(piece);
            }
        }

        return result;
    }

    private static bool Same(InlinePiece a, InlinePiece b) => a.Bold == b.Bold && a.Italic == b.Italic && a.Strike == b.Strike && a.Link == b.Link;

    private static string RenderRuns(List<InlinePiece> pieces, bool emphasis, string lineBreak, bool inTable)
    {
        var sb = new StringBuilder();
        foreach (var piece in pieces)
        {
            var lines = piece.Text.Split('\n');
            for (var i = 0; i < lines.Length; i++)
            {
                if (i > 0)
                {
                    sb.Append(lineBreak);
                }

                sb.Append(Wrap(Escape(lines[i], inTable), piece, emphasis));
            }
        }

        return sb.ToString();
    }

    private static string Wrap(string text, InlinePiece piece, bool emphasis)
    {
        if (!emphasis || (!piece.Bold && !piece.Italic && !piece.Strike) || text.Trim().Length == 0)
        {
            return text;
        }

        var inner = text.Trim();
        var leading = text[..(text.Length - text.TrimStart().Length)];
        var trailing = text[(text.TrimEnd().Length)..];

        var marker = (piece.Bold ? "**" : string.Empty) + (piece.Italic ? "*" : string.Empty);
        inner = marker + inner + new string(marker.Reverse().ToArray());
        if (piece.Strike)
        {
            inner = "~~" + inner + "~~";
        }

        return leading + inner + trailing;
    }

    private static string EscapeUrl(string url) =>
        url.Replace(" ", "%20", StringComparison.Ordinal).Replace("(", "%28", StringComparison.Ordinal).Replace(")", "%29", StringComparison.Ordinal);

    /// <summary>第一列當標題列；各列補齊到相同欄數；儲存格要先跳脫過 | 並把換行換成 &lt;br&gt;。</summary>
    public static string RenderTable(List<List<string>> rows)
    {
        var width = rows.Max(r => r.Count);
        string Line(IReadOnlyList<string> cells) =>
            "| " + string.Join(" | ", Enumerable.Range(0, width).Select(c => c < cells.Count && cells[c].Length > 0 ? cells[c] : " ")) + " |";

        var lines = new List<string> { Line(rows[0]), "| " + string.Join(" | ", Enumerable.Repeat("---", width)) + " |" };
        lines.AddRange(rows.Skip(1).Select(Line));
        return string.Join('\n', lines);
    }
}

/// <summary>依序累積 Markdown 區塊，同時記錄各標題的位置，最後算出每一節的範圍。</summary>
internal sealed class MarkdownDocumentBuilder
{
    private const int SeparatorLength = 2; // "\n\n"

    private readonly StringBuilder _text = new();
    private readonly List<(int Level, string Title, int Start)> _headings = [];
    private BlockKind? _last;
    private int? _lastListKey;

    public int Length => _text.Length;

    public void AddHeading(int level, string title)
    {
        var clean = Regex.Replace(title, @"\s+", " ", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1)).Trim();
        if (clean.Length == 0)
        {
            return;
        }

        Begin(BlockKind.Heading);
        _headings.Add((level, clean, _text.Length));
        _text.Append(new string('#', Math.Min(level, 6))).Append(' ').Append(clean);
    }

    public void AddParagraph(string markdown) => Add(BlockKind.Paragraph, markdown);

    /// <param name="listKey">清單的識別（Word 的 numId）；相鄰的項目屬於不同清單時用空行分開，同一清單的項目只用換行。</param>
    public void AddListItem(string markdown, int? listKey = null)
    {
        if (markdown.Trim().Length == 0)
        {
            return;
        }

        var newList = _last == BlockKind.ListItem && _lastListKey != listKey;
        Begin(BlockKind.ListItem, forceBlankLine: newList);
        _lastListKey = listKey;
        _text.Append(markdown);
    }

    public void AddTable(string markdown) => Add(BlockKind.Table, markdown);

    public ParsedDocument Build(IReadOnlyList<string> notes)
    {
        var text = _text.ToString();
        var sections = new List<SectionInfo>();

        if (_headings.Count > 0 && _headings[0].Start > 0 || _headings.Count == 0 && text.Length > 0)
        {
            var end = _headings.Count > 0 ? _headings[0].Start - SeparatorLength : text.Length;
            sections.Add(new SectionInfo("s0", 0, "(前言)", 0, end, end, null));
        }

        var parents = new Stack<(int Level, string Id)>();
        for (var i = 0; i < _headings.Count; i++)
        {
            var (level, title, start) = _headings[i];
            var ownEnd = i + 1 < _headings.Count ? _headings[i + 1].Start - SeparatorLength : text.Length;
            var next = _headings.Skip(i + 1).Where(h => h.Level <= level).Select(h => (int?)h.Start).FirstOrDefault();
            var totalEnd = next is { } n ? n - SeparatorLength : text.Length;

            while (parents.Count > 0 && parents.Peek().Level >= level)
            {
                parents.Pop();
            }

            var id = $"s{i + 1}";
            sections.Add(new SectionInfo(id, level, title, start, ownEnd, totalEnd, parents.Count > 0 ? parents.Peek().Id : null));
            parents.Push((level, id));
        }

        return new ParsedDocument(text, sections, notes);
    }

    private void Add(BlockKind kind, string markdown)
    {
        if (markdown.Trim().Length == 0)
        {
            return;
        }

        Begin(kind);
        _text.Append(markdown);
    }

    private void Begin(BlockKind kind, bool forceBlankLine = false)
    {
        if (_last is not null)
        {
            _text.Append(kind == BlockKind.ListItem && _last == BlockKind.ListItem && !forceBlankLine ? "\n" : "\n\n");
        }

        _last = kind;
    }
}
