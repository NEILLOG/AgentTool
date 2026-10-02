# Agent 文件工具規劃（C#）

更新日期：2026-10-03

## 決策摘要

Excel、Word、PPT、PDF 工具全部用 C# .NET 8 自己實作，不打包 Python，先以內建工具的形式接進 WPF agent。

- **Excel**：.xlsx 完整讀寫（ClosedXML）；.xlsm 第一版只讀。含 ClosedXML 無法保留內容（圖表、樞紐、VBA）的檔案不覆寫原檔，只能另存。
- **Word / PPT**：只做唯讀，轉成 Markdown 給 agent 理解內文，不做編輯。
- **PDF**：唯讀，PdfPig 抽文字，抽不到的頁面改用 OCR（Windows 上用 Windows 內建 OCR）。
- **部署**：工具跑在使用者本機，直接讀寫本機檔案；公司 IIS / k8s 暫不使用。
- **介接**：工具 library 不依賴 agent，透過薄 adapter 接現有的 ITool 介面（里程碑 2，介面待確認後實作）；之後要支援 MCP 時，同一個 library 再包一層 MCP server。

不選 Python 的理由：Python 唯一明顯勝出的是複雜 PDF 表格辨識（如 Docling），但代價是數百 MB 以上的打包體積、雙語言維護、防毒誤判風險。先用 C# 版本驗證，真的不夠再把 Python 工具掛成獨立 MCP server。

## 文件索引

| 檔案 | 內容 |
| --- | --- |
| [01-architecture.md](01-architecture.md) | 分層架構、資料夾結構 |
| [02-excel-tools.md](02-excel-tools.md) | ExcelTools.Core 完整讀寫 |
| [03-document-reader.md](03-document-reader.md) | Word / PPT / PDF 唯讀轉 Markdown |
| [04-common-security-errors.md](04-common-security-errors.md) | 共用元件、安全、錯誤設計 |
| [05-roadmap-testing.md](05-roadmap-testing.md) | 三個里程碑（Library → Agent 介接 → Windows 驗證）、測試、PDF 驗證 |
| [06-out-of-scope.md](06-out-of-scope.md) | 暫不做與未來選項 |
| [07-environment.md](07-environment.md) | 開發環境驗證結果與平台相容性 |

## 已確認決策（2026-10-03）

- ITool 介接延到里程碑 2，待使用者確認介面格式後實作；里程碑 1 只做 library。
- .xlsm 第一版只讀。
- 閒置逾時遇到未存變更：先存備份再釋放（見 02）。
