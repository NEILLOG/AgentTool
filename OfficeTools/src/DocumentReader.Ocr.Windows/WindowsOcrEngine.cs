using System.Runtime.InteropServices.WindowsRuntime;
using DocumentReader.Core.Pdf;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

namespace DocumentReader.Ocr.Windows;

/// <summary>
/// 用 Windows 內建的 OCR（Windows.Media.Ocr）辨識頁面圖片。語言依序嘗試 <c>languageTags</c>，用第一個已安裝的；
/// 預設是繁體中文、再退到使用者設定語言。沒有可用語言時 <see cref="IsAvailable"/> 回 false 並說明如何安裝。
/// 未在 Windows 實機驗證（見 plan/05 里程碑 3）。
/// </summary>
public sealed class WindowsOcrEngine : IOcrEngine
{
    private readonly string[] _languageTags;

    public WindowsOcrEngine(params string[] languageTags) =>
        _languageTags = languageTags.Length > 0 ? languageTags : ["zh-Hant-TW", "zh-TW", "zh-Hant"];

    public bool IsAvailable(out string? reason)
    {
        if (Create() is not null)
        {
            reason = null;
            return true;
        }

        reason = $"沒有可用的 OCR 語言。請在「設定 → 時間與語言 → 語言與區域」安裝繁體中文（含 OCR 功能），目前嘗試的語言：{string.Join("、", _languageTags)}";
        return false;
    }

    public string Recognize(byte[] png)
    {
        ArgumentNullException.ThrowIfNull(png);
        var engine = Create() ?? throw new InvalidOperationException("沒有可用的 OCR 語言");

        // WinRT 的非同步 API 在這裡同步等待：呼叫端（PdfReader）是同步介面，OCR 本來就在背景工作中執行
        return Task.Run(async () =>
        {
            using var stream = new InMemoryRandomAccessStream();
            await stream.WriteAsync(png.AsBuffer());
            stream.Seek(0);

            var decoder = await BitmapDecoder.CreateAsync(stream);
            var width = (int)decoder.PixelWidth;
            var height = (int)decoder.PixelHeight;
            if (Math.Max(width, height) > OcrEngine.MaxImageDimension)
            {
                throw new InvalidOperationException($"圖片 {width}×{height} 超過 OCR 的上限 {OcrEngine.MaxImageDimension} 像素，請降低 dpi");
            }

            using var bitmap = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);
            var result = await engine.RecognizeAsync(bitmap);

            // 中日文之間 Windows OCR 會在字詞間插空格；一行一行輸出，行之間用換行
            return string.Join("\n", result.Lines.Select(l => CleanLine(l.Text)));
        }).GetAwaiter().GetResult();
    }

    private OcrEngine? Create()
    {
        foreach (var tag in _languageTags)
        {
            var language = new Language(tag);
            if (OcrEngine.IsLanguageSupported(language))
            {
                return OcrEngine.TryCreateFromLanguage(language);
            }
        }

        return OcrEngine.TryCreateFromUserProfileLanguages();
    }

    /// <summary>去掉中日文字元之間多餘的空格（Windows OCR 逐字輸出時會在字之間加空格）。</summary>
    private static string CleanLine(string line) =>
        System.Text.RegularExpressions.Regex.Replace(line, @"(?<=[⺀-鿿＀-￯]) (?=[⺀-鿿＀-￯])", string.Empty, System.Text.RegularExpressions.RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
}
