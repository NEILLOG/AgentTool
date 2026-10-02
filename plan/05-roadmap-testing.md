# 05 開發順序與驗證

分三個里程碑。里程碑 1 只做 library，先做 Excel 再做文件讀取，在 Mac 上每一步都能用 ConsoleSandbox 和測試驗證。

## 里程碑 1：Library（Mac）

0. **已完成。** 專案骨架：`git init`、`global.json`、`Directory.Build.props`、`Directory.Packages.props`（固定 ClosedXML 與相容的 `DocumentFormat.OpenXml` 版本）、空 solution 與測試專案，跑通 `dotnet build` / `dotnet test`。
1. `OfficeTools.Common`：`PathGuard`、錯誤碼、設定。**已完成。**
2. Excel 底層：`A1Address`（A1 位址解析）、`CellValueConverter`（值轉換），附單元測試。**已完成。**
3. `WorkbookSessionManager` + File / Sheet / Range 操作（含資料保全偵測、原子存檔、逾時備份），用 `ConsoleSandbox` 手動試。寫入公式前先驗證語法（見 02）。
4. Excel Format 操作、`Find`、`CopyRange`。
5. Word reader：大綱、分節讀取、Markdown 輸出。
6. PPT reader：投影片大綱、文字排序、備註、圖表數據。
7. PDF reader：**先做 spike 確認 PDFtoImage 在 osx-arm64 可用**；再做 PdfPig 抽取、品質判斷、頁面轉圖片；Windows OCR 實作（Mac 上只 build）。

## 里程碑 2：Agent 介接（待使用者確認 ITool 介面後才開始）

1. 取得 ITool 定義（見 01 的待確認清單），必要時抽出 `Agent.Abstractions`。
2. `OfficeTools.AgentAdapters`：工具命名、參數 schema、例外轉文字、內容邊界標記；用 fake ITool 測試。
3. 依 ITool 是否支援圖片，決定 `RenderPages` 的交付方式。

## 里程碑 3：Windows 驗證

- Windows OCR 實測、接進 WPF agent 整合測試、golden file 用 Excel 實際開啟。

## 測試重點

- xUnit，每個測試獨立暫存資料夾。
- 測試檔案盡量用程式產生（含中文、合併儲存格、修訂追蹤、雙欄 PDF）放在 `tests/TestFiles`；真實業務檔案放不進版控的資料夾（加入 `.gitignore`）。
- A1 解析邊界：`XFD1048576`、整欄 `A:A`、整列 `1:1`。
- 值轉換來回一致；截斷與 `NextRange` 正確；鎖檔情境；路徑穿越、symlink、UNC、ADS 被擋。
- 資料保全：含圖表 / 樞紐 / VBA 的檔案 `Save` 回 `UNSAFE_TO_OVERWRITE`；存檔中斷不會破壞原檔。
- 逾時備份：有未存變更的 session 逾時後產生備份、之後呼叫回 `SESSION_EXPIRED` 並附路徑；備份失敗時 session 保留（逾時用可注入的時鐘測試）。
- 輸出的 xlsx 用 `OpenXmlValidator` 自動驗證（Mac 可跑）；里程碑 3 再用 Excel 實際開一次，確認沒有「需要修復」的警告。

## PDF 驗證（決定要不要追加 Python）

- [ ] 收集 10 到 20 份實際會處理的 PDF，包含掃描檔與表格密集的報表。
- [ ] 讓 agent 讀完後回答內容問題，記錄正確率。
- [ ] 表格頁若普遍答錯，先試「轉圖片給多模態模型」，仍不夠再評估 Docling MCP server。
