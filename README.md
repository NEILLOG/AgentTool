# AgentTool：給 AI agent 用的 Office 文件工具

純 C#（.NET 8）寫的 Office 文件工具 library，讓 agent 能**讀寫 Excel**，以及**讀取 Word / PowerPoint / PDF**（轉成 Markdown）。不依賴 Python、不依賴已安裝的 Office，跑在使用者本機，直接操作本機檔案。

> 目前狀態：library 已完成並有完整測試（1057 個）；**尚未接上 agent**（ITool 介接見 [INTEGRATION-ITool.md](INTEGRATION-ITool.md)）；Windows 專屬項目尚未實機驗證。

## 能做什麼

| 格式 | 能力 | 使用的函式庫 |
| --- | --- | --- |
| Excel `.xlsx` | 完整讀寫：開檔 / 建檔 / 存檔 / 另存、工作表增刪改、讀寫儲存格範圍、公式、插入刪除列欄、尋找、複製、格式、欄寬、合併、凍結窗格 | ClosedXML |
| Excel `.xlsm` | 唯讀（寫入會被拒絕，避免弄壞巨集） | ClosedXML |
| Word `.docx` | 唯讀：標題大綱、依章節讀、依字數分頁；轉成 Markdown（表格、清單、連結、修訂追蹤已接受） | Open XML SDK |
| PowerPoint `.pptx` | 唯讀：投影片大綱、依範圍讀；文字依位置排序、表格、圖表數據、演講者備註 | Open XML SDK |
| PDF | 唯讀：文字層轉 Markdown（依框線重建表格、雙欄閱讀順序、頁首頁尾去除、連結）；掃描頁交給 OCR；頁面可轉成圖片 | PdfPig、PDFtoImage、Windows OCR |

設計重點：

- **保護使用者的檔案**：路徑白名單（含 symlink、UNC、ADS 檢查）、覆寫前自動備份、原子存檔（先寫暫存檔再替換）、含圖表 / 樞紐 / VBA 的活頁簿拒絕覆寫原檔（只能另存）、結構操作失敗自動復原、閒置逾時有未存變更時先存備份。
- **對 agent 友善**：所有輸出有大小上限並附 `Truncated` / `NextRange` / `NextOffset` 續讀位置；錯誤是 17 種有固定代碼的 `OfficeToolException`，每個都帶「下一步該怎麼做」的提示。
- **library 與 agent 解耦**：library 只回傳可序列化的 record，不依賴任何 agent 框架；接 agent 時只需要一層薄 adapter。

## 專案結構

```
AgentTool/
├─ README.md                      本檔
├─ INTEGRATION-ITool.md           給負責接進 agent 的人（或另一個 agent）的介接說明
├─ plan/                          設計文件（架構、各工具行為、安全與錯誤、路線圖）
├─ sample/                        範例文件（docx / pptx / pdf；不進版控）
└─ OfficeTools/
   ├─ OfficeTools.sln
   ├─ src/
   │  ├─ OfficeTools.Common/            PathGuard、錯誤碼、設定、封裝檢查
   │  ├─ ExcelTools.Core/               Excel 讀寫（Session + File/Sheet/Range/Format 操作）
   │  ├─ DocumentReader.Core/           Word / PPT / PDF 唯讀轉 Markdown
   │  └─ DocumentReader.Ocr.Windows/    Windows 內建 OCR 實作（只能在 Windows 執行）
   └─ tests/                            三個測試專案（xUnit）
```

## 快速開始

需求：.NET SDK 8.0.420（由 `OfficeTools/global.json` 固定）。macOS、Windows 都能建置與測試；`DocumentReader.Ocr.Windows` 在 Mac 上只能編譯、不能執行。

```bash
cd OfficeTools
dotnet build
dotnet test
```

程式碼範例：

```csharp
var options = new OfficeToolsOptions { AllowedRoots = [@"C:\Users\me\Documents"] };
var guard = new PathGuard(options);

// Excel
var excelOptions = new ExcelToolsOptions();
using var sessions = new WorkbookSessionManager(guard, excelOptions);
var files = new FileOperations(guard, options, sessions);
var ranges = new RangeOperations(sessions, excelOptions);

var book = files.Open(@"C:\Users\me\Documents\report.xlsx");
var data = ranges.ReadRange(book.WorkbookId, "Sheet1", "A1:D20");
ranges.WriteRange(book.WorkbookId, "Sheet1", "E1", [["合計"], ["=SUM(D2:D20)"]]);
files.Save(book.WorkbookId);            // 覆寫前自動備份

// Word / PPT / PDF（無狀態，唯讀）
var readerOptions = new DocumentReaderOptions();
var word = new WordReader(guard, options, readerOptions);
var outline = word.GetOutline(@"C:\Users\me\Documents\spec.docx");
var section = word.ReadSection(@"C:\Users\me\Documents\spec.docx", "s2");

var pdf = new PdfReader(guard, options, readerOptions /*, ocrEngine */);
var page = pdf.ReadPages(@"C:\Users\me\Documents\invoice.pdf", 1);
```

## 文件

| 文件 | 內容 |
| --- | --- |
| [INTEGRATION-ITool.md](INTEGRATION-ITool.md) | **接進 agent 的完整說明**：工具清單與參數、錯誤對應、生命週期、安全規則 |
| [plan/README.md](plan/README.md) | 設計文件索引與已確認決策 |
| [plan/01-architecture.md](plan/01-architecture.md) | 分層架構 |
| [plan/02-excel-tools.md](plan/02-excel-tools.md) | Excel 各操作的行為與已知限制（含 ClosedXML 的各種陷阱） |
| [plan/03-document-reader.md](plan/03-document-reader.md) | Word / PPT / PDF 的轉換規則與已知限制 |
| [plan/04-common-security-errors.md](plan/04-common-security-errors.md) | 安全設計與錯誤碼表 |
| [plan/05-roadmap-testing.md](plan/05-roadmap-testing.md) | 里程碑、測試、待驗證清單 |

## 進度與已知限制

- ✅ 里程碑 1（library）完成：Excel 全部操作、Word / PPT / PDF reader。
- ⏳ 里程碑 2（agent 介接）：待 ITool 介面確認，說明見 [INTEGRATION-ITool.md](INTEGRATION-ITool.md)。
- ⏳ 里程碑 3（Windows 驗證）：OCR 實測、PDFium 原生檔、Excel 存檔的檔案控制代碼與鎖檔行為，清單見 plan/05。
- 真實文件驗證覆蓋有限：Word / PPT / PDF 各用 1–3 份真實檔案驗過，Excel 只用程式產生的檔案。
- 主要限制：PDF 沒有框線的表格不會被偵測、設計稿式版面的區塊順序是近似；PPT 項目符號不看母片；Word 字母 / 羅馬數字編號會轉成阿拉伯數字；ClosedXML 存檔會丟失圖表 / 樞紐 / VBA（所以拒絕覆寫原檔）。完整清單見各設計文件。
