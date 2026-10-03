using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ExcelTools.Core.Models;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Internal;

/// <summary>解析並驗證 <see cref="FormatSpec"/>。ClosedXML 對這些設定幾乎不驗證（無效的數字格式、不認得的顏色名稱會悄悄變成透明），所以在套用前自己檢查。</summary>
internal static partial class FormatRules
{
    public const double MaxFontSize = 409;

    private static readonly string[] NamedColors =
        ["black", "white", "red", "green", "blue", "yellow", "orange", "purple", "pink", "gray", "grey", "brown", "cyan", "magenta", "lightgray", "darkgray", "lightblue", "darkblue", "lightgreen", "darkgreen", "gold", "silver", "navy", "teal"];

    [GeneratedRegex(@"^#?([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$", RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 1000)]
    private static partial Regex HexColor();

    // 數字格式中中括號內允許的內容：顏色、Color1-56、條件、地區 / 貨幣 [$...]、經過時間 [h] [mm] [ss]、DBNum
    [GeneratedRegex(@"\[([^\]]*)\]", RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 1000)]
    private static partial Regex Bracket();

    [GeneratedRegex(@"^(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\s?\d{1,2}|[<>=]{1,2}.+|\$.*|h+|m+|s+|DBNum[1-4]|Natural.*)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 1000)]
    private static partial Regex AllowedBracket();

    /// <summary>已驗證、可直接套用的格式；null 欄位代表不改。</summary>
    internal sealed record Parsed(
        bool? Bold,
        bool? Italic,
        bool? Underline,
        bool? Strikethrough,
        string? FontName,
        double? FontSize,
        XLColor? FontColor,
        XLColor? FillColor,
        bool ClearFill,
        XLAlignmentHorizontalValues? Horizontal,
        XLAlignmentVerticalValues? Vertical,
        bool? WrapText,
        string? NumberFormat,
        XLBorderStyleValues? BorderStyle,
        XLColor? BorderColor,
        bool OutlineOnly);

    public static Parsed Parse(FormatSpec spec)
    {
        ArgumentNullException.ThrowIfNull(spec);

        if (spec.Bold is null && spec.Italic is null && spec.Underline is null && spec.Strikethrough is null
            && spec.FontName is null && spec.FontSize is null && spec.FontColor is null && spec.FillColor is null
            && spec.HorizontalAlignment is null && spec.VerticalAlignment is null && spec.WrapText is null
            && spec.NumberFormat is null && spec.BorderStyle is null && spec.BorderColor is null && spec.BorderSides is null)
        {
            throw Invalid("沒有指定任何格式", "請至少指定一個屬性，例如 Bold、FillColor、NumberFormat");
        }

        if (spec.FontName is not null && (string.IsNullOrWhiteSpace(spec.FontName) || spec.FontName.Length > 31 || spec.FontName.Any(char.IsControl)))
        {
            throw Invalid($"無效的 FontName「{spec.FontName}」", "字型名稱不可為空、最長 31 字元、不可含控制字元");
        }

        if (spec.FontSize is { } size && (double.IsNaN(size) || size < 1 || size > MaxFontSize))
        {
            throw Invalid($"無效的 FontSize {spec.FontSize}", $"字型大小要介於 1 到 {MaxFontSize}");
        }

        var clearFill = spec.FillColor is not null && spec.FillColor.Trim().Equals("none", StringComparison.OrdinalIgnoreCase);
        var borderStyle = spec.BorderStyle is null ? (XLBorderStyleValues?)null : ParseEnum<XLBorderStyleValues>(spec.BorderStyle, nameof(spec.BorderStyle));
        if (spec.BorderColor is not null && spec.BorderStyle is null)
        {
            throw Invalid("BorderColor 必須搭配 BorderStyle", "例如 BorderStyle = \"thin\"、BorderColor = \"#808080\"");
        }

        var outline = false;
        if (spec.BorderSides is not null)
        {
            if (spec.BorderStyle is null)
            {
                throw Invalid("BorderSides 必須搭配 BorderStyle", "例如 BorderStyle = \"thin\"、BorderSides = \"outline\"");
            }

            outline = spec.BorderSides.Trim().ToUpperInvariant() switch
            {
                "ALL" => false,
                "OUTLINE" => true,
                _ => throw Invalid($"無效的 BorderSides「{spec.BorderSides}」", "可用的值：all（每個儲存格四邊）、outline（範圍外框）"),
            };
        }

        return new Parsed(
            spec.Bold,
            spec.Italic,
            spec.Underline,
            spec.Strikethrough,
            spec.FontName?.Trim(),
            spec.FontSize,
            spec.FontColor is null ? null : ParseColor(spec.FontColor, nameof(spec.FontColor)),
            spec.FillColor is null || clearFill ? null : ParseColor(spec.FillColor, nameof(spec.FillColor)),
            clearFill,
            spec.HorizontalAlignment is null ? null : ParseEnum<XLAlignmentHorizontalValues>(spec.HorizontalAlignment, nameof(spec.HorizontalAlignment)),
            spec.VerticalAlignment is null ? null : ParseEnum<XLAlignmentVerticalValues>(spec.VerticalAlignment, nameof(spec.VerticalAlignment)),
            spec.WrapText,
            spec.NumberFormat is null ? null : ValidateNumberFormat(spec.NumberFormat),
            borderStyle,
            spec.BorderColor is null ? null : ParseColor(spec.BorderColor, nameof(spec.BorderColor)),
            outline);
    }

    public static XLColor ParseColor(string text, string field)
    {
        var value = text.Trim();
        if (HexColor().IsMatch(value))
        {
            return XLColor.FromHtml(value.StartsWith('#') ? value : "#" + value);
        }

        // XLColor.FromName 對不認得的名稱不會丟例外，而是悄悄回傳透明色，所以只接受白名單裡的名稱
        if (NamedColors.Contains(value, StringComparer.OrdinalIgnoreCase))
        {
            return XLColor.FromName(value);
        }

        throw Invalid($"無效的 {field}「{text}」", "請用 #RRGGBB（例如 #FF0000）、#RGB，或名稱：" + string.Join("、", NamedColors));
    }

    /// <summary>檢查 Excel 數字格式碼。無效的格式可能讓 Excel 開檔時要求修復。</summary>
    public static string ValidateNumberFormat(string format)
    {
        var trimmed = format.Trim();
        if (trimmed.Length == 0 || trimmed.Length > 255)
        {
            throw Invalid($"無效的 NumberFormat「{format}」", "數字格式不可為空（要還原請用 General），最長 255 字元。" + NumberFormatHint);
        }

        if (!new ExcelNumberFormat.NumberFormat(trimmed).IsValid)
        {
            throw Invalid($"無效的 NumberFormat「{format}」", NumberFormatHint);
        }

        // 雙引號與括號外的內容：中括號內只允許已知的內容
        foreach (Match m in Bracket().Matches(StripQuoted(trimmed)))
        {
            if (!AllowedBracket().IsMatch(m.Groups[1].Value))
            {
                throw Invalid($"無效的 NumberFormat「{format}」：不認得中括號內的「{m.Groups[1].Value}」", NumberFormatHint);
            }
        }

        return trimmed;
    }

    private const string NumberFormatHint =
        "數字格式碼範例：0.00、#,##0、0.0%、yyyy-mm-dd、yyyy/mm/dd hh:mm、@（文字）、General、[Red]0.00、#,##0;[Red]-#,##0；要顯示文字請用雙引號包起來";

    private static string StripQuoted(string format) => Regex.Replace(format, "\"[^\"]*\"", string.Empty, RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));

    private static T ParseEnum<T>(string text, string field)
        where T : struct, Enum
    {
        var name = text.Trim();
        var match = Enum.GetNames<T>().FirstOrDefault(n => n.Equals(name, StringComparison.OrdinalIgnoreCase));
        if (match is null)
        {
            throw Invalid($"無效的 {field}「{text}」", "可用的值：" + string.Join("、", Enum.GetNames<T>().Select(n => char.ToLowerInvariant(n[0]) + n[1..])));
        }

        return Enum.Parse<T>(match);
    }

    private static OfficeToolException Invalid(string message, string hint) => new(ErrorCodes.InvalidValue, message, hint);
}
