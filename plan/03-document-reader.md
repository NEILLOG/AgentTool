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

## Word（已實作：`WordReader`）

唯讀、只支援 .docx。解析結果依「路徑 + 最後修改時間 + 檔案大小 + 讀取選項」做 LRU 快取，同一份文件連續讀多次不會重複解析。開檔時讀進記憶體後立即放開檔案，並用共用的 `PackageChecks` 檢查密碼保護（回 `PASSWORD_PROTECTED`）、損壞（`CORRUPT_FILE`）、解壓縮炸彈與單一部件字元上限（`FILE_TOO_LARGE`）。

**介面**

- `GetOutline(path)`：標題樹。每個節點有 `SectionId`（`s0` 是第一個標題之前的前言，其餘依文件順序 `s1`、`s2`…）、`Level`、`Title`、`Chars`（自己）與 `TotalChars`（含子節）；`HasHeadings` 告訴你文件有沒有標題；`Notes` 說明略過或簡化的內容（目錄、圖片數量、追蹤修訂等）。
- `ReadSection(path, sectionId, includeSubsections = true, offset = 0, maxChars)`：讀一節（含標題行；預設含子節）。`offset` 是這一節範圍內的字元位置。找不到回 `SECTION_NOT_FOUND`，hint 列出現有章節（最多 30 節）。依序讀完所有最上層節點再用 `\n\n` 連起來，等於整份文件。
- `Read(path, offset = 0, maxChars)`：依字元數分頁讀整份文件，給沒有標題的文件用。超過 `MaxCharsPerRead` 時回傳 `Truncated` 與 `NextOffset`；**盡量停在段落邊界**（前半段內找不到空行才退到換行，再不行才硬切，且不切在代理對中間）。`maxChars` 不可超過設定的上限；`offset` 超出範圍回 `INVALID_VALUE`。
- 讀取選項 `WordReadOptions`：`IncludeHeadersFooters`、`IncludeFootnotes`、`IncludeComments`，開啟的內容附在文件最後的「附錄：…」區段（也出現在標題樹裡）；沒開但文件有這些內容時，`Notes` 會提醒。

**轉換規則**

- **標題**：中文版 Word 的內建標題樣式 **ID 是 `"1"`、`"2"`，名稱才是 `heading 1`**，所以不能用樣式 ID 判斷。依序看：段落直接設定的大綱層級 → 樣式（含 `basedOn` 鏈，有循環保護）的大綱層級 → 樣式名稱 `heading N` / `標題 N`（`Title` 視為 1 級）。大綱層級 9（內文）明確表示不是標題。Markdown 最多 6 個 `#`，但標題樹保留原本的層級（到 9）。標題文字去掉格式與連結、換行改空白。**標題的自動編號**（例如 `1.2`）會當前綴加進標題文字，支援多層 `%1.%2`、阿拉伯 / 字母 / 羅馬數字、`isLgl`、起始值覆寫。
- **清單**：項目符號 `-`、編號 `N.`（用計數器算出實際編號，每個編號清單各自計數，高層項目會讓低層歸零，支援起始值覆寫）；巢狀以每層 2 個空格縮排。字母 / 羅馬數字的清單在 Markdown 裡一律用阿拉伯數字當標記。清單也可由段落樣式帶入編號；`numId = 0` 取消編號。相鄰但屬於不同清單的項目之間用空行分開。空的編號段落不消耗編號。
- **行內**：粗體、斜體、刪除線（來自 run 或字元樣式）；相同格式的相鄰 run 合併；**空白移到強調標記外**，避免 `** 文字 **` 這種無效的 Markdown。外部超連結 → `[文字](網址)`，網址用原始字串（`Uri.ToString()` 會補斜線並還原 `%20`）；內部錨點只留文字。隱藏文字略過；換行 → Markdown 硬換行；分頁符號、分欄符號忽略。**只跳脫會被誤判的字元**（`\ * \``，表格內加 `|`），行首像標題 / 清單 / 引用 / 分隔線的文字加反斜線；底線與中括號不跳脫。
- **表格**：Markdown 表格，第一列當標題列。橫向合併（`gridSpan`）與直向合併（`vMerge`）都把值**重複填進被合併的每一格**；`gridBefore` / `gridAfter` 以空格補齊；各列補到同樣欄數。儲存格內多段落以 `<br>` 連接，`|` 跳脫，清單項目加 `•` / `1.` 前綴；**巢狀表格降級為文字**（`（巢狀表格：a / b; c / d）`）。全空的表格略過。
- **圖片與物件**：`[圖片: 替代文字]`（沒有替代文字為 `[圖片]`，替代文字壓成單行）；圖表 `[圖表]`、SmartArt `[SmartArt]`、嵌入物件 `[嵌入物件]`；舊式 VML 圖片也支援；沒有文字也沒有圖片的圖形（裝飾線）忽略。**文字方塊**以 `[文字方塊] …` 獨立段落輸出，且只處理 `mc:AlternateContent` 的 `Choice`（`Fallback` 是同樣內容的相容版本，會重複）。圖片 / 圖表 / 公式數量在 `Notes`。
- **修訂追蹤**：輸出接受全部修訂後的版本（略過 `del`、`moveFrom`，保留 `ins`、`moveTo`）；整段被刪除的段落消失。
- **其他**：內容控制項（`sdt`）展開；目錄（`TOC` 樣式、目錄控制項）略過並在 `Notes` 說明；欄位（頁碼等）只保留結果、不顯示欄位代碼；數學公式 → `[公式: 純文字]`（OMML 的結構如分數會遺失）。
- 頁首頁尾（相同的只列一次）、註腳 / 章節附註（`[^f1]` / `[^e1]`，只有開啟時才加參照標記）、註解（`作者：內容`）：預設不輸出。

**已知限制**：字母 / 羅馬數字等編號格式在 Markdown 清單中以阿拉伯數字呈現；中文數字等其他格式的標題編號也退回阿拉伯數字；OMML 數學公式只轉純文字；圖表數據、SmartArt、圖片內容不轉換（要看圖請用多模態模型）。**目前只用程式產生的 docx 測過**（手寫 XML 與 Open XML SDK 物件模型各一套），尚未用真正的 Word 產出的檔案驗證，見 05。

## PPT

**介面**（`PptReader`，與 `WordReader` 共用 `ParsedDocumentCache` 與分頁邏輯）

- `GetOutline(path)`：每張投影片的編號、標題、字元數、是否隱藏、有無備註、圖片 / 表格 / 圖表數量；`Notes` 說明略過或簡化的內容。
- `ReadSlides(path, from, to = null, offset = 0, maxChars = null)`：讀第 `from` 到 `to` 張（含；省略 `to` 只讀一張）。`offset` 是範圍內的字元位置，被截斷時用 `NextOffset` 續讀。
- `Read(path, offset, maxChars)`：整份簡報依字元數分頁。
- 投影片順序以簡報的 `sldIdLst` 為準，不是檔名；解析結果依「路徑 + 修改時間 + 大小」快取（備註選項也是快取鍵的一部分）；超過 `MaxCharsPerRead` 截斷於段落邊界。

**轉換規則**

- 每張投影片一個 `## 投影片 N：標題`（沒有標題為 `## 投影片 N`，隱藏的加「（隱藏）」）。標題取自標題預留位置（`title` / `ctrTitle`），不會在內文重複；同一張的第二個標題預留位置當一般文字。日期、頁尾、投影片編號預留位置略過。
- **文字順序**：標題 → 其他預留位置（依文件順序）→ 其他形狀依位置排序：先由上而下，垂直距離在投影片高度 1/30 以內視為同一列，列內由左而右。群組形狀遞迴展開、整組當一個單位依群組位置排序。隱藏的形狀與群組略過；`mc:AlternateContent` 只取 `Choice`。
- **段落**：粗體 / 斜體 / 刪除線保留；外部超連結（http / https / mailto）轉 Markdown 連結，其他連結只留文字；換行 `a:br` 為硬換行；`_x000B_` 這類 OOXML 跳脫還原、控制字元移除；Markdown 語法字元與行首的 `#`、`1.`、`-` 都跳脫。
- **項目符號**：明確有符號（`buChar`、`buAutoNum`、`buBlip`）的段落成為清單項目，依 `lvl` 縮排，自動編號按層級計數、被一般段落打斷就重來；內容 / 本文預留位置在沒有明寫 `buNone` 時預設有符號（符號其實由母片決定，這裡**不解析母片**）。不同形狀的清單用空行分開。
- **表格**：轉 Markdown 表格，`hMerge` / `vMerge` 的被合併儲存格填入來源值；全空的表格略過。
- **圖表**：輸出 `[圖表: 標題（類型）]` 加一張「分類 × 系列」的表，數據取自 chart XML 內的快取（`c:strCache` / `c:numCache`，不需要內嵌 xlsx），最多 100 列；組合圖併成一張表；沒有快取只輸出標籤。
- **圖片**：`[圖片: 替代文字]`，沒有替代文字為 `[圖片]`，連續多張無替代文字的合併為 `[圖片 ×N]`；SmartArt `[SmartArt]`、嵌入物件 `[嵌入物件]`，內容不轉換。
- **演講者備註**：預設輸出為 `> **備註：** …` 引用區塊（`IncludeNotes = false` 關閉）；只取備註本文，不含投影片縮圖與頁碼。

**已知限制**：項目符號不看母片；SmartArt、圖片內容、動畫、版面配置圖形（沒有文字的形狀）不轉換；形狀位置沿用投影片自己的座標，**沒有座標、也不是預留位置**的形狀視為在左上角；母片與版面配置上的固定文字（Logo 文字、浮水印）不輸出；多欄閱讀順序是依位置的近似，卡片式或不規則版面的順序可能與設計者的意圖不同（例如標籤與說明被分開）。

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
