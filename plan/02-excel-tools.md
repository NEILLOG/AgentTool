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
- **閒置從操作結束算起**：長時間的操作不算閒置。沒有背景計時器，每次存取時順便清理（主程式也可定期呼叫 `SweepExpired`）。
- **閒置逾時**：沒有未存變更的 session 逾時直接釋放。有未存變更的 session **先存備份再釋放**：
  - 備份存在原檔同資料夾，檔名 `原檔名.autosave-yyyyMMdd-HHmmss.xlsx`（仍受 `PathGuard` 約束；不覆寫原檔，所以不受資料保全限制）。
  - 釋放後保留一筆「已逾時」紀錄（workbookId → 備份路徑）。之後用該 ID 呼叫時回 `SESSION_EXPIRED`，hint 附上備份路徑，請 agent 告知使用者並用 `Open` 開備份接續。
  - 備份失敗（例如沒有寫入權限）時不釋放 session，保留在記憶體，`ListOpen` 標記為逾時待處理。

## 存檔語意

- `Save`：只用於由 `Open` 開啟的既有檔案，視為允許覆寫原檔，**每個 session 第一次 Save 前備份原檔**（`原檔名.backup-時間戳.xlsx`，不受 `AllowOverwrite` 控制）；同一 session 之後的 Save 不再重複備份，避免產生一堆備份檔。備份失敗（唯讀資料夾、空間不足）時中止存檔並回 `FILE_LOCKED`，原檔不動。
- `Create` / `SaveAs` 目的地已存在：依 `AllowOverwrite` 決定，預設回 `FILE_EXISTS`；允許時依 `BackupOnOverwrite` 先備份。
- **原子寫入**：先存到同資料夾的暫存檔（`.名稱.guid.tmp.xlsx`），再用 `File.Replace`（目標不存在則 `File.Move`）換掉目標檔，避免寫到一半造成壞檔。目標被鎖時回 `FILE_LOCKED`。
- **ClosedXML 每個活頁簿實例只能存一次**（實測 0.105.1）：它會記住上一次存檔的串流或檔案，第二次存檔時重新開啟，目標已釋放或被換名就丟 `ObjectDisposedException` / `FileNotFoundException`。因此每次存檔成功後，用剛寫出的內容重新載入一個乾淨的活頁簿取代 session 內的實例（`WorkbookSession.ReloadFrom`）。副作用：每次存檔多一次解析，檔案大時稍慢。
- 存檔一律用 `SaveAs(string path)`，不用 `SaveAs(Stream)`；路徑的副檔名必須是 .xlsx（ClosedXML 會檢查，所以暫存檔名以 .xlsx 結尾）。

## 操作清單（第一版）

| 類別 | 方法 |
| --- | --- |
| File | `ListFiles`、`Create`、`Open`（順便回工作表摘要與 `PreservationWarnings`）、`Save`、`SaveAs`、`Close`、`ListOpen` |
| Sheet | `ListSheets`、`GetSheetInfo`、`Add`、`Rename`、`Copy`、`Delete`、`Move` |
| Range | `ReadRange`、`WriteRange`、`AppendRows`、`ClearRange`、`InsertRows`、`DeleteRows`、`InsertColumns`、`DeleteColumns`、`Find`、`CopyRange` |
| Format | `FormatRange`、`SetColumnWidth`、`AutoFitColumns`、`Merge`、`Unmerge`、`FreezePanes` |

## 工作表操作的行為（實測 ClosedXML 0.105.1）

- 工作表一律以名稱指定，**不分大小寫**（與 Excel 一致）；位置是 1 起算的分頁順序。找不到回 `SHEET_NOT_FOUND`，hint 列出現有工作表（依分頁順序）。
- `wb.Worksheets` 的列舉順序是**建立順序**，不是分頁順序，要用 `Position` 排序；所有清單輸出都已排序。
- **改名**：ClosedXML 會同步更新其他工作表公式與已定義名稱裡的參照（含需要加引號的名稱）。名稱重複（不分大小寫）、不合法（空白、超過 31 字元、含 `\ / ? * [ ] :`、以單引號開頭或結尾）回 `INVALID_VALUE`；只改大小寫視為合法。
- **刪除**：ClosedXML **不會**處理其他地方對該工作表的參照，Excel 開啟後會變 `#REF!`。因此刪除前掃描公式與已定義名稱，結果以 `SheetChangeResult.Warnings` 回報（列出最多 10 處與總數）。另外 ClosedXML 允許刪到 0 張或只剩隱藏表（Excel 會要求修復），所以自行擋下：至少保留一張工作表，且至少一張可見。
- **複製**：預設命名為「原名 (2)」「原名 (3)」…（超過 31 字元時截斷原名），預設位置緊接在來源之後（與 Excel 一致）。
- 修改類操作都回傳操作後的工作表清單，讓 agent 不必再查一次。
- `GetSheetInfo` 的已使用範圍以「有內容」為準，空的合併儲存格不算。

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
- 公式計算失敗時回傳快取值並標記 `CalculationWarnings`，不讓整個操作失敗。實測 ClosedXML 0.105 讀取公式時會丟例外的情況：語法錯誤（`ExpressionParseException`）、循環參照（`InvalidOperationException`）、不支援的運算（例如範圍交集，`NotImplementedException`），讀取端要逐格捕捉。
- **寫入公式要先驗證語法**：實測 `FormulaA1 = "SUM(A1:"` 設定與存檔都不會報錯，壞公式會被寫進檔案、Excel 開啟時要求修復。寫入前先用 `ClosedXML.Parser` 解析，語法錯誤回 `INVALID_VALUE`（附錯誤位置），不寫入。未知函式名稱（`NOSUCHFN(1)`）不會被擋，會得到 `#NAME?`，這是 Excel 本身的行為，不視為錯誤。
- 以字串寫入時，ISO 日期字串預設轉成日期（可用 `ParseIsoDates` 關閉）；`00123` 這類看起來像數字的字串一律保持文字。
- `SHEET_NOT_FOUND` 錯誤直接附上現有工作表清單，省一輪重試。
