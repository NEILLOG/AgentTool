# 05 開發順序與驗證

分三個里程碑。里程碑 1 只做 library，先做 Excel 再做文件讀取，在 Mac 上每一步都能用 ConsoleSandbox 和測試驗證。

## 里程碑 1：Library（Mac）

0. **已完成。** 專案骨架：`git init`、`global.json`、`Directory.Build.props`、`Directory.Packages.props`（固定 ClosedXML 與相容的 `DocumentFormat.OpenXml` 版本）、空 solution 與測試專案，跑通 `dotnet build` / `dotnet test`。
1. `OfficeTools.Common`：`PathGuard`、錯誤碼、設定。**已完成。**
2. Excel 底層：`A1Address`（A1 位址解析）、`CellValueConverter`（值轉換），附單元測試。**已完成。**
3. `WorkbookSessionManager` + File / Sheet / Range 操作（**3a File 與 Session、3b Sheet 已完成**；3c-1 Range 讀寫（ReadRange、WriteRange、AppendRows、ClearRange、公式驗證）已完成，3c-2 結構操作（InsertRows、DeleteRows、InsertColumns、DeleteColumns，含檢查點復原）、Find、CopyRange 已完成。**第 3 步的 File / Sheet / Range 全部完成**）（含資料保全偵測、原子存檔、逾時備份），用 `ConsoleSandbox` 手動試。寫入公式前先驗證語法（見 02）。
4. Excel Format 操作（`FormatRange`、`SetColumnWidth`、`AutoFitColumns`、`Merge`、`Unmerge`、`FreezePanes`）、`Find`、`CopyRange`。**已完成**（`Find`、`CopyRange` 在 3c-2 完成）。
5. Word reader：大綱、分節讀取、Markdown 輸出。**已完成**（見 03）。
6. PPT reader：投影片大綱、文字排序、備註、圖表數據。**已完成**（見 03；用使用者提供的真實 sample.pptx 驗證過）。
7. PDF reader：**已完成**（見 03）。PDFtoImage 在 osx-arm64 實測可用（約 0.1–0.5 秒一頁）；PdfPig 文字抽取 + 依框線建表格 + XY-cut 閱讀順序 + 品質判斷 + 頁面轉圖片；`IOcrEngine` 與 Windows 實作（`DocumentReader.Ocr.Windows`，Mac 上只編譯）。

## 里程碑 2：Agent 介接（待使用者確認 ITool 介面後才開始）

1. 取得 ITool 定義（見 01 的待確認清單），必要時抽出 `Agent.Abstractions`。
2. `OfficeTools.AgentAdapters`：工具命名、參數 schema、例外轉文字、內容邊界標記；用 fake ITool 測試。
3. 依 ITool 是否支援圖片，決定 `RenderPages` 的交付方式。

## 里程碑 3：Windows 驗證

- Windows OCR 實測、接進 WPF agent 整合測試、golden file 用 Excel 實際開啟。
- **PDF / OCR（Mac 上無法驗證）**：
  - [ ] `WindowsOcrEngine` 在裝了繁體中文語言套件的機器上辨識掃描頁；沒裝時 `IsAvailable` 回 false 並顯示安裝提示。
  - [ ] 200 dpi 的 A4 頁面不超過 `OcrEngine.MaxImageDimension`（超過時要降 dpi）。
  - [ ] Windows 上 PDFtoImage（`bblanchon.PDFium.Win32`）能轉圖，打包後原生檔有被帶上。
  - [ ] 中日文之間的空格清理（`CleanLine`）對真實辨識結果合理。
- **Excel 存檔相關（Mac 上無法驗證）**：
  - [ ] ClosedXML `SaveAs(path)` 之後暫存檔沒有被持有，`File.ReadAllBytes` 與 `File.Replace` 能成功。
  - [ ] 使用者在 Excel 開著同一個檔案時，`Open` 仍可讀；`Save` 在 `File.Replace` 失敗時回 `FILE_LOCKED` 且原檔不受影響。
  - [ ] `FileShare` 與 symlink 相關測試（目前在沒有權限建立 symlink 時會直接略過）。

## 測試重點

- xUnit，每個測試獨立暫存資料夾。
- 測試檔案盡量用程式產生（含中文、合併儲存格、修訂追蹤、雙欄 PDF）放在 `tests/TestFiles`；真實業務檔案放不進版控的資料夾（加入 `.gitignore`）。
- A1 解析邊界：`XFD1048576`、整欄 `A:A`、整列 `1:1`。
- 值轉換來回一致；截斷與 `NextRange` 正確；鎖檔情境；路徑穿越、symlink、UNC、ADS 被擋。
- 資料保全：含圖表 / 樞紐 / VBA 的檔案 `Save` 回 `UNSAFE_TO_OVERWRITE`；存檔中斷不會破壞原檔。
- 逾時備份：有未存變更的 session 逾時後產生備份、之後呼叫回 `SESSION_EXPIRED` 並附路徑；備份失敗時 session 保留（逾時用可注入的時鐘測試）。
- 輸出的 xlsx 用 `OpenXmlValidator`（Office2019）自動驗證（Mac 可跑，已用於建立與工作表操作；附負面對照確認驗證器真的會報錯）；里程碑 3 再用 Excel 實際開一次，確認沒有「需要修復」的警告。

## Word 真實文件驗證

- [x] 用使用者提供的一份真實 Word 檔（36 頁中文會訊，多層標題、超連結、表格、圖片）驗證：標題樹、章節內容、連結輸出合理。
- [ ] 再收集 5 到 10 份實際會處理的 Word 文件（含多層編號標題、追蹤修訂、文字方塊、目錄），確認標題樹、章節範圍、表格與清單輸出合理。其餘功能目前只用程式產生的 docx 測過，真正由 Word 產出的檔案可能有我沒想到的結構。
- [ ] 讓 agent 用 `GetOutline` → `ReadSection` 讀完文件後回答內容問題，記錄正確率。

## PPT 真實文件驗證

- [x] 用使用者提供的一份真實 pptx（40 張，Google 簡報匯出的範本，含表格、群組形狀、超連結、一張有 5888 個形狀的投影片）驗證：結構與文字輸出合理，整份解析約 2 秒。
- [ ] 這份檔案沒有圖表、備註、SmartArt、隱藏投影片，這幾項只用程式產生的 pptx 測過；請再給一份有圖表與備註的簡報（PowerPoint 存出的）。
- [ ] 讓 agent 用 `GetOutline` → `ReadSlides` 讀完簡報後回答內容問題，記錄正確率。

## PDF 驗證（決定要不要追加 Python）

- [x] 用使用者提供的三份真實 PDF 驗證：Notion 風格行程表（3 頁，表格、連結、康熙部首字元）、含框線表格的報價單（跨欄標題、合併儲存格）、設計稿式複雜排版（標註框、小字清單）。表格與清單輸出正確；複雜排版只能做到各區塊內容完整，區塊間順序是近似。
- [ ] 這三份都有文字層：**掃描檔與亂碼檔只用程式產生的 PDF（整頁圖片）測過**，請再給一份真實掃描檔，並在 Windows 上實測 OCR。
- [ ] 收集 10 到 20 份實際會處理的 PDF，包含掃描檔與表格密集的報表。
- [ ] 讓 agent 讀完後回答內容問題，記錄正確率。
- [ ] 表格頁若普遍答錯，先試「轉圖片給多模態模型」，仍不夠再評估 Docling MCP server。
