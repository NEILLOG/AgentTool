using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentReader.Core.Word;

/// <param name="IsBullet">項目符號（不是編號）。</param>
/// <param name="Counter">這一層目前的編號值。</param>
/// <param name="Text">套用 lvlText 後的標籤，例如 "1.2"、"第一章"、"a)"；項目符號為空。</param>
internal sealed record NumberingLabel(bool IsBullet, int Counter, string Text);

/// <summary>
/// 編號定義與計數：每個編號清單（numId）各自維護每一層的計數，遇到某一層的項目就把更深的層歸零。
/// 用來產生「1.」「2.」的清單標記，以及標題前面的自動編號（例如 1.2）。
/// </summary>
internal sealed class WordNumbering
{
    private const int Levels = 9;

    private sealed record LevelDef(NumberFormatValues? Format, string? Text, int Start, bool IsLegal);

    private readonly Dictionary<int, int> _abstractOf = [];
    private readonly Dictionary<(int Abstract, int Level), LevelDef> _levels = [];
    private readonly Dictionary<(int NumId, int Level), int> _startOverrides = [];
    private readonly Dictionary<int, int?[]> _counters = [];

    public static WordNumbering Load(MainDocumentPart main)
    {
        var result = new WordNumbering();
        var numbering = main.NumberingDefinitionsPart?.Numbering;
        if (numbering is null)
        {
            return result;
        }

        foreach (var abs in numbering.Elements<AbstractNum>())
        {
            var absId = abs.AbstractNumberId?.Value;
            if (absId is null)
            {
                continue;
            }

            foreach (var level in abs.Elements<Level>())
            {
                if (level.LevelIndex?.Value is { } index)
                {
                    result._levels[(absId.Value, index)] = new LevelDef(
                        level.NumberingFormat?.Val?.Value,
                        level.LevelText?.Val?.Value,
                        level.StartNumberingValue?.Val?.Value ?? 1,
                        level.IsLegalNumberingStyle is not null && (level.IsLegalNumberingStyle.Val?.Value ?? true));
                }
            }
        }

        foreach (var instance in numbering.Elements<NumberingInstance>())
        {
            var numId = instance.NumberID?.Value;
            var absRef = instance.AbstractNumId?.Val?.Value;
            if (numId is null || absRef is null)
            {
                continue;
            }

            result._abstractOf[numId.Value] = absRef.Value;
            foreach (var over in instance.Elements<LevelOverride>())
            {
                if (over.LevelIndex?.Value is { } index && over.StartOverrideNumberingValue?.Val?.Value is { } start)
                {
                    result._startOverrides[(numId.Value, index)] = start;
                }
            }
        }

        return result;
    }

    /// <summary>消耗一個編號項目：這一層的計數加一、更深的層歸零，回傳標籤。</summary>
    public NumberingLabel Next(int numId, int level)
    {
        level = Math.Clamp(level, 0, Levels - 1);
        if (!_counters.TryGetValue(numId, out var counters))
        {
            counters = new int?[Levels];
            _counters[numId] = counters;
        }

        counters[level] = counters[level] is { } current ? current + 1 : StartOf(numId, level);
        for (var deeper = level + 1; deeper < Levels; deeper++)
        {
            counters[deeper] = null;
        }

        var def = Definition(numId, level);
        if (def.Format == NumberFormatValues.Bullet)
        {
            return new NumberingLabel(true, counters[level]!.Value, string.Empty);
        }

        if (def.Format == NumberFormatValues.None)
        {
            return new NumberingLabel(false, counters[level]!.Value, string.Empty);
        }

        return new NumberingLabel(false, counters[level]!.Value, Substitute(numId, level, counters, def));
    }

    private string Substitute(int numId, int level, int?[] counters, LevelDef def)
    {
        var template = def.Text ?? $"%{level + 1}.";
        var sb = new StringBuilder();
        for (var i = 0; i < template.Length; i++)
        {
            if (template[i] == '%' && i + 1 < template.Length && char.IsAsciiDigit(template[i + 1]))
            {
                var referenced = template[i + 1] - '1';
                if (referenced is >= 0 and < Levels)
                {
                    var value = counters[referenced] ?? StartOf(numId, referenced);
                    var format = def.IsLegal ? NumberFormatValues.Decimal : Definition(numId, referenced).Format; // isLgl：這個層級顯示的所有編號（含自己）一律用阿拉伯數字
                    sb.Append(Format(value, format));
                    i++;
                    continue;
                }
            }

            sb.Append(template[i]);
        }

        return sb.ToString().Trim();
    }

    private int StartOf(int numId, int level) =>
        _startOverrides.TryGetValue((numId, level), out var over) ? over : Definition(numId, level).Start;

    private LevelDef Definition(int numId, int level) =>
        _abstractOf.TryGetValue(numId, out var abs) && _levels.TryGetValue((abs, level), out var def)
            ? def
            : new LevelDef(NumberFormatValues.Decimal, $"%{level + 1}.", 1, false);

    private static string Format(int value, NumberFormatValues? format)
    {
        if (format == NumberFormatValues.DecimalZero)
        {
            return value.ToString("00", CultureInfo.InvariantCulture);
        }

        if (format == NumberFormatValues.LowerLetter)
        {
            return Letters(value).ToLowerInvariant();
        }

        if (format == NumberFormatValues.UpperLetter)
        {
            return Letters(value);
        }

        if (format == NumberFormatValues.LowerRoman)
        {
            return Roman(value).ToLowerInvariant();
        }

        if (format == NumberFormatValues.UpperRoman)
        {
            return Roman(value);
        }

        return value.ToString(CultureInfo.InvariantCulture); // decimal，以及其他不支援的格式（中文數字等）一律用阿拉伯數字
    }

    private static string Letters(int value)
    {
        var sb = new StringBuilder();
        for (var n = Math.Max(value, 1); n > 0; n = (n - 1) / 26)
        {
            sb.Insert(0, (char)('A' + ((n - 1) % 26)));
        }

        return sb.ToString();
    }

    private static string Roman(int value)
    {
        (int Value, string Symbol)[] table = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")];
        var sb = new StringBuilder();
        var n = Math.Clamp(value, 1, 3999);
        foreach (var (v, symbol) in table)
        {
            while (n >= v)
            {
                sb.Append(symbol);
                n -= v;
            }
        }

        return sb.ToString();
    }
}
