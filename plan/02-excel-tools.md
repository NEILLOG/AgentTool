# 02 ExcelTools.Core：完整讀寫

以 ClosedXML（MIT）實作，不需安裝 Excel。第一版：.xlsx 可讀寫；.xlsm 只讀（ClosedXML 存檔無法保證保留 VBA）。

## 設定

```csharp
public sealed class ExcelToolsOptions
{
    public int MaxCellsPerRead { get; init; } = 2000;
    public TimeSpan IdleTimeout { get; init; } = TimeSpan.FromMinutes(30);
}
```

（路徑、檔案大小、覆寫相關設定在 `OfficeToolsOptions`，見 04。）

## 資料保全（ClosedXML 限制）

ClosedXML 載入既有檔案後再存檔，會遺失或損壞它不支援的內容：圖表、樞紐分析、VBA、Slicer、部分圖形。

- `Open` 時掃描套件部件（chart、pivotTable / pivotCache、vbaProject、drawing 內非圖片的物件、slicer），在回傳摘要附上 `PreservationWarnings`。
- 有 `PreservationWarnings` 的活頁簿，`Save` 拒絕覆寫原檔，回 `UNSAFE_TO_OVERWRITE`，只能 `SaveAs` 到新檔。
- .xlsm 一律唯讀，寫入類操作回 `UNSUPPORTED_FORMAT`，hint 建議另存為 .xlsx。

## Session 模型

- `Open(path)` 回傳 `workbookId`，後續操作都帶這個 ID，連續修改不用重複開檔。
- **不長時間佔用檔案**：用 `FileShare.ReadWrite` 讀進 `MemoryStream` 後立即關閉檔案，ClosedXML 從記憶體載入。使用者在 Excel 開著也能讀，session 也不會擋住使用者開檔。
- 每個 session 一把 `SemaphoreSlim`，因為 ClosedXML 不是執行緒安全的。
- 記錄是否有未存檔變更；`Close` 時有變更就回 `UNSAVED_CHANGES`，除非帶 `discardChanges`。
- 同一路徑重複開啟回傳既有 ID。
- **閒置逾時**：沒有未存變更的 session 逾時直接釋放。有未存變更的 session **先存備份再釋放**：
  - 備份存在原檔同資料夾，檔名 `原檔名.autosave-yyyyMMdd-HHmmss.xlsx`（仍受 `PathGuard` 約束；不覆寫原檔，所以不受資料保全限制）。
  - 釋放後保留一筆「已逾時」紀錄（workbookId → 備份路徑）。之後用該 ID 呼叫時回 `SESSION_EXPIRED`，hint 附上備份路徑，請 agent 告知使用者並用 `Open` 開備份接續。
  - 備份失敗（例如沒有寫入權限）時不釋放 session，保留在記憶體，`ListOpen` 標記為逾時待處理。

## 存檔語意

- `Save`：只用於由 `Open` 開啟的既有檔案，視為允許覆寫原檔，**一律先備份**（不受 `AllowOverwrite` 控制）。
- `Create` / `SaveAs` 目的地已存在：依 `AllowOverwrite` 決定，預設回 `FILE_EXISTS`；允許時依 `BackupOnOverwrite` 先備份。
- **原子寫入**：先存到同資料夾的暫存檔，再用 `File.Replace`（或 `File.Move` 覆寫）換掉目標檔，避免寫到一半造成壞檔。目標被鎖時回 `FILE_LOCKED`。

## 操作清單（第一版）

| 類別 | 方法 |
| --- | --- |
| File | `ListFiles`、`Create`、`Open`（順便回工作表摘要與 `PreservationWarnings`）、`Save`、`SaveAs`、`Close`、`ListOpen` |
| Sheet | `ListSheets`、`GetSheetInfo`、`Add`、`Rename`、`Copy`、`Delete`、`Move` |
| Range | `ReadRange`、`WriteRange`、`AppendRows`、`ClearRange`、`InsertRows`、`DeleteRows`、`InsertColumns`、`DeleteColumns`、`Find`、`CopyRange` |
| Format | `FormatRange`、`SetColumnWidth`、`AutoFitColumns`、`Merge`、`Unmerge`、`FreezePanes` |

## 資料表示

```csharp
public sealed record RangeData(
    string Sheet,
    string Range,               // 實際回傳的範圍
    object?[][] Values,         // 數字/字串/布林/null，日期轉 ISO 字串
    string?[][]? Formulas,      // 只有 IncludeFormulas 時才回
    bool Truncated,
    string? NextRange);         // 被截斷時，下一段要讀的範圍
```

- 讀取回傳緊湊的 2D 陣列；可選擇附公式（`IncludeFormulas`）、顯示文字（`UseFormattedText`）。
- 超過 `MaxCellsPerRead` 就截斷，回傳 `Truncated` 與 `NextRange`。
- 寫入時 `=` 開頭當公式、ISO 日期字串轉日期，兩者都可關閉；`null` 代表清空。
- 公式計算失敗時回傳快取值並標記 `CalculationWarnings`，不讓整個操作失敗。
- `SHEET_NOT_FOUND` 錯誤直接附上現有工作表清單，省一輪重試。
