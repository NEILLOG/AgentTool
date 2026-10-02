# 03 DocumentReader.Core：Word / PPT / PDF 唯讀

全部轉成 Markdown 給 agent。唯讀不需 Session，做成無狀態。不支援 .doc / .ppt 舊格式。

函式庫：Open XML SDK（MIT）；PDF 用 PdfPig（Apache 2.0）。PPT 不用 ShapeCrawler，直接用 Open XML SDK，少一個依賴。

## 設定與共通行為

```csharp
public sealed class DocumentReaderOptions
{
    public int MaxCharsPerRead { get; init; } = 20_000;
    public int OcrDpi { get; init; } = 200;
    public int CacheMaxEntries { get; init; } = 20;
}
```

- 每次讀取最多回傳 `MaxCharsPerRead` 字元；超過就截斷，回傳 `Truncated` 與續讀位置（與 Excel 的 `NextRange` 同一模式）。
- 快取：key 為「路徑 + 最後修改時間 + 檔案大小」，LRU，上限 `CacheMaxEntries`，避免 WPF 常駐程序記憶體持續增長。
- 加密檔回 `PASSWORD_PROTECTED`；結構損壞回 `CORRUPT_FILE`。
- 文件內容視為不可信資料（可能含提示注入），見 04。

## Word

**轉換規則**

- 標題：沿段落樣式的 `basedOn` 鏈找 `outlineLvl`，找不到再看內建 styleId（`Heading1`~`Heading6`），轉成 `#` 到 `######`。**不可用樣式名稱比對**，中文 Word 的樣式名稱是「標題 1」。
- 清單：依編號定義轉成 `-` 或 `1.`，保留縮排層級。
- 表格：轉 Markdown 表格；合併儲存格重複填值，巢狀表格降級為文字。
- 圖片：輸出 `[圖片: 替代文字]`。
- 修訂追蹤：輸出接受全部修訂後的版本（略過 `w:del`，保留 `w:ins`）。
- 頁首頁尾、註腳、註解：預設不輸出，用選項開啟。

**分段讀取**（docx 沒有頁碼資訊，改用標題切段）

- `GetOutline(path)`：標題樹，每節帶 sectionId 與字數。
- `ReadSection(path, sectionId)`：讀一節。
- `Read(path, offset, maxChars)`：依字數分頁，給沒有標題的文件用。

## PPT

**轉換規則**

- 每張投影片一個 `## 投影片 N：標題` 區塊。
- 文字順序：先預留位置，其他形狀依位置由上而下、由左而右排序。
- 群組形狀遞迴展開；表格轉 Markdown 表格。
- 圖表輸出標題與數據系列：優先讀 chart XML 內的快取值（`c:strCache` / `c:numCache`），內嵌 xlsx 只當備援（可能是外部連結或不存在）。
- 演講者備註預設輸出；隱藏投影片加標記。
- SmartArt 第一版先跳過。

**分段讀取**

- `GetOutline(path)`：每張投影片的編號與標題。
- `ReadSlides(path, from, to)`。

## PDF：逐頁混合策略

1. 用 PdfPig 抽每頁文字，利用字的座標處理雙欄順序，並以「跨頁重複出現的行」去除頁首頁尾。
2. 判斷抽取品質：字數過少，或亂碼比例高（私用區字元、大量問號），標記為需要 OCR。中文字型缺 Unicode 對應表時會出現這種情況。
3. 需要 OCR 的頁面用 PDFium 包裝（例如 PDFtoImage）轉成圖片，交給 `IOcrEngine`。
4. 另提供「把指定頁面轉成圖片」的工具，讓 agent 遇到複雜表格或圖表時自行交給多模態模型看（library 先做到輸出圖片；如何交給 agent 待里程碑 2 確認 ITool 能否回傳圖片）。

**介面**

- `GetOutline(path)`：頁數、每頁是否需要 OCR、PDF 書籤（若有）。
- `ReadPages(path, from, to)`：回傳 Markdown，標註哪些頁面來自 OCR。
- `RenderPages(path, pages, dpi)`：輸出圖片給多模態模型。

**OCR 實作**

- `IOcrEngine` 定義在 `DocumentReader.Core`。
- Windows 實作在 `DocumentReader.Ocr.Windows`（`Windows.Media.Ocr`），TargetFramework `net8.0-windows10.0.19041.0`。
- 繁體中文需在使用者電腦安裝中文語言套件：啟動時用 `OcrEngine.AvailableRecognizerLanguages` 偵測，沒有 zh-Hant 就把頁面標記 `OcrUnavailable`，hint 說明如何安裝。
- 渲染給 OCR 的圖片長寬不得超過 `OcrEngine.MaxImageDimension`，超過就降低 DPI。
- 沒有 OCR 引擎時（例如在 Mac 開發），需要 OCR 的頁面回傳空內容並標記 `OcrUnavailable`，不讓整個讀取失敗。
- Windows OCR 對清楚的印刷文字效果不錯；手寫、低解析度、表格結構較弱。
