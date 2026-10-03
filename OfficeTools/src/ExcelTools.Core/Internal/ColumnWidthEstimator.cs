using System.Globalization;
using System.Text;
using ClosedXML.Excel;

namespace ExcelTools.Core.Internal;

/// <summary>
/// 估算欄寬（Excel 的欄寬單位 ≈ 預設字型一個數字「0」的寬度）。不依賴作業系統的字型：
/// ClosedXML 的 AdjustToContents 在沒有 CJK 字型的系統（例如 Mac）會把中文量成極窄（實測 18 個中文字只算出 9.75），
/// 所以自己估算，讓 Mac 與 Windows 的結果一致。Excel 實際的自動調整依字型，會有些微差異。
/// </summary>
internal static class ColumnWidthEstimator
{
    public const double DefaultFontSize = 11;

    private const double Padding = 1.0;

    public static double Estimate(string text, double fontSize, bool bold)
    {
        var longestLine = 0.0;
        foreach (var line in text.Split('\n'))
        {
            longestLine = Math.Max(longestLine, Units(line.TrimEnd('\r')));
        }

        if (longestLine == 0)
        {
            return 0;
        }

        return (longestLine * (fontSize / DefaultFontSize) * (bold ? 1.1 : 1.0)) + Padding;
    }

    public static double Units(string line)
    {
        var total = 0.0;
        foreach (var rune in line.EnumerateRunes())
        {
            total += Width(rune);
        }

        return total;
    }

    private static double Width(Rune rune)
    {
        if (IsWide(rune.Value))
        {
            return 2.0;
        }

        if (rune.Value > 0xFFFF)
        {
            return 1.0;
        }

        var c = (char)rune.Value;
        return c switch
        {
            ' ' => 0.5,
            '.' or ',' or ':' or ';' or '!' or '|' or '\'' or '`' or 'i' or 'j' or 'l' or 'I' => 0.5,
            'm' or 'w' or 'M' or 'W' => 1.5,
            _ when c is >= 'A' and <= 'Z' => 1.1,
            _ when c is >= 'a' and <= 'z' => 0.95,
            _ when char.GetUnicodeCategory(c) == UnicodeCategory.Control => 0,
            _ => 1.0,
        };
    }

    /// <summary>全形字（中日韓文字與標點、全形英數、表情符號）佔兩個字元寬度。</summary>
    private static bool IsWide(int cp) =>
        cp is >= 0x1100 and <= 0x115F          // 韓文字母
            or >= 0x2E80 and <= 0x303E         // CJK 部首與符號、標點
            or >= 0x3041 and <= 0x33FF         // 日文假名、注音、CJK 相容
            or >= 0x3400 and <= 0x4DBF         // CJK 擴充 A
            or >= 0x4E00 and <= 0x9FFF         // CJK 統一表意文字
            or >= 0xA000 and <= 0xA4CF         // 彝文
            or >= 0xAC00 and <= 0xD7A3         // 韓文音節
            or >= 0xF900 and <= 0xFAFF         // CJK 相容表意文字
            or >= 0xFE30 and <= 0xFE6F         // CJK 相容形式
            or >= 0xFF00 and <= 0xFF60         // 全形英數與標點
            or >= 0xFFE0 and <= 0xFFE6         // 全形符號
            or >= 0x1F300 and <= 0x1FAFF       // 表情符號
            or >= 0x20000 and <= 0x3FFFD;      // CJK 擴充 B 以後

    /// <summary>儲存格的字型（大小、粗體）→ 估算用的參數。</summary>
    public static (double FontSize, bool Bold) FontOf(IXLCell cell) =>
        (cell.Style.Font.FontSize > 0 ? cell.Style.Font.FontSize : DefaultFontSize, cell.Style.Font.Bold);
}
