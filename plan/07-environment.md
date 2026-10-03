# 07 開發環境驗證

驗證日期：2026-10-03

## 已驗證

| 項目 | 結果 |
| --- | --- |
| 開發機 | MacBook Air（Apple Silicon，arm64） |
| Claude 執行方式 | Claude Code（VS Code 擴充）直接跑在 Mac 本機，工作目錄 AgentTool |
| Claude 可用的 .NET SDK | 8.0.420、6.0.428（`/usr/local/share/dotnet`），2026-10-03 確認 |

結論：Claude 可以在這台 Mac 上直接 build、跑測試、修到全綠。（先前在隔離 Linux VM 內無 .NET SDK、連不到 NuGet 的限制已不適用。）

## Mac 本機（使用者於 2026-10-03 回報）

| 項目 | 結果 | 狀態 |
| --- | --- | --- |
| 架構 / 系統 | arm64，macOS 26.6.2 | OK |
| .NET SDK | 8.0.420（另有 6.0.428） | OK |
| .NET Runtime | Microsoft.NETCore.App 8.0.26、AspNetCore 8.0.26 | OK |
| NuGet 來源 | nuget.org 已啟用 | OK |
| git | 2.50.1 | OK |
| VS Code `code` 指令 | 沒有輸出 | 待確認 |

待辦：

- [ ] VS Code：若已安裝，在 VS Code 按 Cmd+Shift+P 執行「Shell Command: Install 'code' command in PATH」；若未安裝則安裝，並加裝 C# Dev Kit。
- [ ] 在 solution 根目錄放 `global.json`（列入 05 的第 0 步） 固定 SDK 版本（機器上有 6.0 與 8.0 兩個 SDK，避免日後裝了新版 SDK 造成行為改變）：

```json
{
  "sdk": {
    "version": "8.0.420",
    "rollForward": "latestFeature"
  }
}
```

## 平台相容性

| 元件 | Mac 上可開發與測試 | 備註 |
| --- | --- | --- |
| OfficeTools.Common | 可 | |
| ExcelTools.Core（ClosedXML） | 可 | 欄寬自動調整不依賴作業系統字型（確定性估算，Mac 與 Windows 一致） |
| DocumentReader.Core：Word / PPT | 可 | Open XML SDK 跨平台 |
| DocumentReader.Core：PDF 文字抽取（PdfPig） | 可 | 純 managed |
| PDF 頁面轉圖片（PDFium 包裝） | 大致可 | 需確認套件有 osx-arm64 原生檔（05 第 7 步先做 spike） |
| DocumentReader.Ocr.Windows | 只能 build，不能執行 | Windows.Media.Ocr 僅 Windows；需 `EnableWindowsTargeting` |
| WPF agent 整合 | 不可 | 需在 Windows 上驗證 |

## 建議的開發流程

1. **在 Mac 上用 Claude Code** 開發與測試 Common、ExcelTools、DocumentReader（里程碑 1）；里程碑 2 的 AgentAdapters 也在 Mac 上用 fake ITool 測試；Claude 自己 build、跑測試、修到全綠。
2. **Windows 機器**負責 05 的里程碑 3：Windows OCR 實測、接進 WPF agent 的整合測試、golden file 用 Excel 實際開啟。
