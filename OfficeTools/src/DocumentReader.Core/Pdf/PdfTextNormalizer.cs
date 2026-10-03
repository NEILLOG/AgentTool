using System.Text;

namespace DocumentReader.Core.Pdf;

internal static class PdfTextNormalizer
{
    // 部首補充區沒有相容對應，Unicode 不會幫忙換；只列常見的（0x2EC4 西、0x2EA0 民），其餘維持原樣
    private static readonly Dictionary<char, char> Supplement = new() { [(char)0x2EC4] = (char)0x897F, [(char)0x2EA0] = (char)0x6C11 };

    /// <summary>
    /// 把 PDF 字型常見的「相容字元」換回一般字元：康熙部首（U+2F8F → 行、U+2F47 → 日）、相容漢字、字母連字（U+FB01 → fi）。
    /// 只處理這幾個區段，不用整串 NFKC，因為 NFKC 也會把全形英數、圈號數字之類改掉，那不是這裡要的。
    /// </summary>
    public static string Normalize(string text)
    {
        StringBuilder? sb = null;
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];
            string? replacement = null;
            if (Supplement.TryGetValue(c, out var mapped))
            {
                replacement = mapped.ToString();
            }
            else if (NeedsMapping(c))
            {
                replacement = c.ToString().Normalize(NormalizationForm.FormKC);
            }
            else if (c == (char)0xA0 || c < ' ')
            {
                // 空白字元（NBSP）與控制字元：有些字型把空格編成 \0，抽出來是 NUL，一律當空格
                replacement = " ";
            }

            if (replacement is null)
            {
                sb?.Append(c);
                continue;
            }

            sb ??= new StringBuilder(text, 0, i, text.Length);
            sb.Append(replacement);
        }

        return sb?.ToString() ?? text;
    }

    // 部首（康熙部首 0x2F00–0x2FDF、補充 0x2E80–0x2EFF）、相容漢字（0xF900–0xFAFF）、字母連字（0xFB00–0xFB06）
    private static bool NeedsMapping(char c) => c is >= (char)0x2E80 and <= (char)0x2FDF or >= (char)0xF900 and <= (char)0xFAFF or >= (char)0xFB00 and <= (char)0xFB06;

    public static bool IsCjk(char c) =>
        c is >= (char)0x2E80 and <= (char)0x9FFF
            or >= (char)0xA000 and <= (char)0xA4CF
            or >= (char)0xAC00 and <= (char)0xD7AF
            or >= (char)0xF900 and <= (char)0xFAFF
            or >= (char)0xFE30 and <= (char)0xFE4F
            or >= (char)0xFF00 and <= (char)0xFFEF;

    /// <summary>私用區、替代字元、控制字元：PDF 字型缺 Unicode 對應表時抽出來的就是這些。NUL 不算（有些字型把空格編成 NUL）。</summary>
    public static bool IsGarbage(char c) =>
        c is >= (char)0xE000 and <= (char)0xF8FF or (char)0xFFFD || (char.IsControl(c) && c is not ('\0' or '\n' or '\t' or '\r'));
}
