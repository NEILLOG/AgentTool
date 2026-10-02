# 01 架構與資料夾結構

## 分層

```
WPF agent 小工具（UI + Agent Core：對話迴圈、使用者確認）
        │
        ├── ITool adapter（包成現有 ITool 介面）
        └── MCP server adapter（之後；讓 Claude、Copilot 等外部 client 也能用）
                │
        ┌───────┴────────┐
ExcelTools.Core     DocumentReader.Core
（完整讀寫）        （Word / PPT / PDF 唯讀轉 Markdown）
        └───────┬────────┘
        OfficeTools.Common（PathGuard、錯誤碼、設定）
```

原則：library 完全不知道 agent 存在，只提供一般 C# 方法、回傳可序列化的 record。寫入類操作的使用者確認放在 agent 端統一處理。

## 資料夾結構

```
OfficeTools/
├─ OfficeTools.sln
├─ Directory.Build.props          // 共用：net8.0、Nullable、LangVersion、警告設定
├─ Directory.Packages.props       // 集中管理 NuGet 版本
├─ global.json                    // 固定 SDK 8.0.420
├─ README.md
├─ src/
│  ├─ OfficeTools.Common/
│  │  ├─ OfficeToolsOptions.cs
│  │  ├─ Security/PathGuard.cs
│  │  └─ Errors/{OfficeToolException.cs, ErrorCodes.cs}
│  ├─ ExcelTools.Core/
│  │  ├─ ExcelToolsOptions.cs
│  │  ├─ Workspace/{WorkbookSessionManager.cs, WorkbookSession.cs}
│  │  ├─ Operations/{File,Sheet,Range,Format}Operations.cs
│  │  ├─ Models/{RangeData.cs, SheetInfo.cs, ReadOptions.cs, FormatSpec.cs}
│  │  └─ Internal/{A1Address.cs, CellValueConverter.cs}
│  ├─ DocumentReader.Core/        // net8.0，跨平台
│  │  ├─ Word/WordReader.cs
│  │  ├─ Ppt/{PptReader.cs, ShapeOrdering.cs}
│  │  ├─ Pdf/{PdfReader.cs, PageQualityChecker.cs, PdfPageRenderer.cs}
│  │  ├─ Ocr/IOcrEngine.cs        // OCR 抽象
│  │  ├─ Markdown/MarkdownWriter.cs
│  │  ├─ Models/{DocumentOutline.cs, ReadResult.cs}
│  │  └─ Caching/ParsedDocumentCache.cs
│  ├─ DocumentReader.Ocr.Windows/ // net8.0-windows10.0.19041.0，Windows.Media.Ocr 實作
│  ├─ Agent.Abstractions/         // 里程碑 2：ITool 介面（若從 WPF 專案抽出）
│  ├─ OfficeTools.AgentAdapters/  // 里程碑 2：包成現有 ITool
│  │  ├─ Excel/
│  │  ├─ Documents/
│  │  └─ ToolRegistration.cs
│  └─ OfficeTools.McpServer/      // 之後才加，console app，stdio
├─ tests/
│  ├─ OfficeTools.Common.Tests/
│  ├─ ExcelTools.Core.Tests/
│  ├─ DocumentReader.Core.Tests/
│  ├─ OfficeTools.AgentAdapters.Tests/  // 里程碑 2：用 fake ITool 測試
│  └─ TestFiles/                  // 範例 xlsx / docx / pptx / pdf、golden files
└─ samples/
   └─ ConsoleSandbox/             // 不經 agent，直接呼叫 library 手動試
```

## 取捨說明

- **OCR 拆成獨立專案**：Windows OCR 只能在 Windows 執行。拆出 `IOcrEngine` 介面後，`DocumentReader.Core` 維持 `net8.0`，在 Mac / Linux 也能 build 與跑測試；OCR 實作只在 Windows 專案裡。該專案要設定 `<EnableWindowsTargeting>true</EnableWindowsTargeting>` 才能在 Mac 上 build。
- **NuGet 版本相容**：ClosedXML 對 `DocumentFormat.OpenXml` 有版本範圍限制，DocumentReader 也直接用 Open XML SDK。在 `Directory.Packages.props` 固定一個雙方都相容的版本，不各自升級。
- **AgentAdapters 獨立**：它要引用 ITool 介面。若 ITool 目前定義在 WPF 專案內，建議抽成小的 `Agent.Abstractions` 專案，避免 adapter 引用整個 WPF。
- **ConsoleSandbox**：開發時不必啟動 WPF 或呼叫 LLM，直接餵檔案看輸出。
- WPF 專案可不放在這個 solution，用 project reference 或本機 NuGet 引用。

## 里程碑 2 待確認：ITool 介面

里程碑 1 只做 library，不碰 ITool。使用者確認以下事項後才實作 adapter（見 05）：

- [ ] 方法簽章：同步或非同步（`Task`）、是否帶 `CancellationToken`。
- [ ] 參數 schema 的描述方式（JSON Schema、屬性標註或其他）。
- [ ] 回傳型別：只能回文字，還是能回圖片等多模態內容？**`RenderPages` 依賴能回傳圖片**；若不行，改為把圖片存檔並回傳路徑，或第一版先不做。
- [ ] 寫入類操作的使用者確認，是由 ITool 宣告（例如 `RequiresConfirmation`）還是 agent 端另有機制。
