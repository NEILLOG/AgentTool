using DocumentReader.Core.Models;
using OfficeTools.Common.Errors;

namespace DocumentReader.Core.Markdown;

/// <summary>把已轉好的 Markdown 全文依字元數切段；Word 與 PPT 共用。</summary>
internal static class TextPaging
{
    /// <summary>
    /// 讀 [<paramref name="scopeStart"/>, <paramref name="scopeEnd"/>) 範圍內從 <paramref name="offset"/> 開始的一段。
    /// 被截斷時盡量停在段落邊界，並用 NextOffset 告訴呼叫端下一次從哪裡接著讀。
    /// </summary>
    public static ReadResult Slice(
        string text,
        int scopeStart,
        int scopeEnd,
        int offset,
        int? maxChars,
        int defaultLimit,
        string? sectionId,
        IReadOnlyList<string> notes)
    {
        var length = Math.Max(0, scopeEnd - scopeStart);
        if (offset < 0 || offset > length)
        {
            throw new OfficeToolException(ErrorCodes.InvalidValue, $"offset {offset} 超出範圍（0 到 {length}）", "offset 要用上一次回傳的 NextOffset");
        }

        if (maxChars is < 1)
        {
            throw new OfficeToolException(ErrorCodes.InvalidValue, $"maxChars {maxChars} 無效", "maxChars 至少是 1");
        }

        var limit = Math.Min(maxChars ?? defaultLimit, defaultLimit);
        var start = scopeStart + offset;
        var end = Math.Min(scopeEnd, start + limit);
        if (end < scopeEnd)
        {
            end = BreakPoint(text, start, end);
        }

        var truncated = end < scopeEnd;
        return new ReadResult(
            text[start..end].TrimEnd('\n'),
            offset,
            end - scopeStart,
            length,
            truncated,
            truncated ? end - scopeStart : null,
            sectionId,
            notes);
    }

    /// <summary>被迫截斷時盡量停在段落邊界（空行），其次是換行；前半段內找不到才硬切，且不切在代理對中間。</summary>
    private static int BreakPoint(string text, int start, int hardEnd)
    {
        var earliest = start + ((hardEnd - start) / 2);
        var paragraph = text.LastIndexOf("\n\n", hardEnd - 1, hardEnd - earliest, StringComparison.Ordinal);
        if (paragraph >= earliest)
        {
            return paragraph + 2;
        }

        var line = text.LastIndexOf('\n', hardEnd - 1, hardEnd - earliest);
        if (line >= earliest)
        {
            return line + 1;
        }

        return hardEnd > start + 1 && char.IsHighSurrogate(text[hardEnd - 1]) ? hardEnd - 1 : hardEnd;
    }
}
