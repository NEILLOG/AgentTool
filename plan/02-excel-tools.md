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

## 範圍操作的行為（`RangeOperations`）

**ReadRange**
- 結果**限縮在工作表有內容的區域內**：起點維持你要求的位置，終點縮到最後有內容的列 / 欄，`Range` 回傳實際範圍。整欄（`A:A`）、整列（`1:1`）、整張表（`A1:XFD1048576`）因此都安全。要求的區域完全在有內容的區域之外時回傳空陣列。
- 超過 `MaxCellsPerRead` 以**整列**為單位截斷，至少回傳一整列（一列比上限寬時仍回整列），`NextRange` 是接著要讀的範圍（一律是儲存格範圍格式）。
- 只走有內容的格子（`CellsUsed`），讀取不會把空格子建出來，也不會讓工作表變成「有變更」。
- `UseFormattedText` 一律用 Invariant culture 格式化。ClosedXML 的 `GetFormattedString()` 不帶參數時會隨使用者的地區設定變化（實測法文變成 `1 234,50`、阿拉伯文用阿拉伯數字），所以固定傳 `CultureInfo.InvariantCulture`。
- 公式算不出來（語法錯誤、循環參照、ClosedXML 解析不了的語法）時，該格回傳檔案中的快取值並在 `CalculationWarnings` 列出（最多 20 筆，其餘合併成一則摘要），不讓整個讀取失敗。公式算出 `#DIV/0!`、`#NAME?` 等錯誤值不算警告，直接當值回傳。
- ClosedXML 會把範圍擴張到涵蓋跨出範圍的合併儲存格，所以讀取時要過濾掉超出要求的格子。

**WriteRange / AppendRows**
- `WriteRange` 的位置可以是單一儲存格（左上角起點），或與資料大小**完全相同**的範圍；整欄整列不可當寫入目標。資料列可長短不一，缺的格子不動；整列是 `null` 就略過；格子是 `null` 代表清空內容（保留格式）。
- **全有或全無**：先把所有值轉換並驗證，全部通過才動工作表。任何一格有問題（型別不支援、NaN、字串太長、公式語法錯誤）都不寫入，錯誤訊息開頭是出問題的儲存格位址。
- **公式語法驗證**（`FormulaValidator`）：在暫存活頁簿裡解析，只攔語法錯誤（`ExpressionParseException`），回 `INVALID_VALUE` 並附公式與位置。用 54 個常見合法語法（陣列常數、結構化參照、`LET`、`XLOOKUP`、動態陣列 `A1#`、跨表 `Sheet1:Sheet3!A1` 等）驗證沒有誤判。不攔：未知函式（得到 `#NAME?`）、不存在的工作表（`#REF!`）、循環參照。
- 寫入後試算剛寫的公式，算出錯誤值或無法計算（含循環參照）時在 `RangeChangeResult.Warnings` 提醒，但資料照常寫入。
- 單次寫入上限 `MaxCellsPerWrite`（預設 20000），超過回 `INVALID_VALUE` 並建議分批；超出工作表邊界（XFD1048576）回 `INVALID_RANGE`。
- `AppendRows` 從有內容區域的下一列、第一個有內容的欄開始；可用 `startColumn` 指定。空工作表從 A1 開始。

**InsertRows / DeleteRows / InsertColumns / DeleteColumns**（實測 ClosedXML 0.105.1 的行為與對策）
- 插入：ClosedXML 會正確調整同表與他表公式、已定義名稱、合併儲存格、表格與資料驗證的範圍（`SUM(A1:A9)` 在中間插入列會變成 `SUM(A1:A11)`）。插入位置是「在該列 / 欄之前」，與 Excel 一致；欄用欄名稱（`"C"`），列用 1 起算的列號。
- **插入會把有內容的儲存格擠出工作表時，ClosedXML 悄悄丟掉內容**（Excel 會拒絕）。所以插入前檢查有內容區域的最後一列 / 欄，會擠出邊界就回 `INVALID_RANGE`。
- **刪除時，指向被刪儲存格的公式 ClosedXML 處理是錯的**：刪列會悄悄把 `=A5*2` 改成 `=A4*2`（指到別格），刪欄則完全不更新。Excel 會變 `#REF!`。因此刪除前掃描所有公式與已定義名稱（`FormulaReferences`，啟發式：先去掉字串常數，排除函式名與表格結構化參照），**整個參照都落在被刪列 / 欄內**的列入 `Warnings`（最多 10 處加總數）；位於被刪區域內、自己會被刪掉的公式不計；範圍只是部分重疊（`SUM(A1:A9)` 刪第 5 列）是正常縮小，不警告。
- **檢查點與復原**：每次結構操作先把活頁簿序列化成檢查點並換一個乾淨的實例來操作，操作後再序列化並用 Open XML 驗證該工作表（含表格部件）與活頁簿部件；操作丟例外或驗證有錯就**復原到檢查點**並回 `UNSAFE_OPERATION`，活頁簿不會被標成有變更。實測會觸發的情況：刪光整個表格（存檔失敗）。每次結構操作約多兩次序列化與解析。
- 另一個實測到的壞檔情境（有資料驗證時在第 1 列插入 / 刪除會產生空的 `sqref`）只發生在「剛用 API 在記憶體中建立」的活頁簿；session 裡的活頁簿一律從檔案載入，實測不會發生，但驗證網仍保留以防其他未知組合。

**Find**
- 比對三種文字：顯示文字（例如 `1,234.50`）、原始值（`1234.5`、ISO 日期、`TRUE`）、以及選用的公式文字（`SearchFormulas`）；`MatchedIn` 說明命中哪一種。預設不分大小寫、子字串；`WholeCell` 要求整格相同。結果依工作表分頁順序、列優先，最多 `MaxResults`（預設 100）筆，超過 `Truncated`。省略 sheet 搜尋全部工作表；指定 range 必須同時指定 sheet。算不出來的公式不影響搜尋（用快取值並附警告）。

**CopyRange**
- 來源縮到實際使用的區域；目標是單一儲存格（貼上範圍的左上角），可跨工作表。`All` 複製值、公式、格式與合併儲存格，相對參照像 Excel 一樣位移（絕對與混合參照維持）；`ValuesOnly` 以計算結果貼值、不帶格式。覆蓋了有內容的儲存格時在 `Warnings` 說明。
- **來源與目標重疊時** ClosedXML 直接複製會讀到已被覆蓋的格子，所以經由暫存工作表快照後貼回（結束後移除）。暫存區必須使用與來源**相同的座標**：放在 A1 時第一步的位移可能把相對參照推出邊界變成 `#REF!`（實測踩到）。

**ClearRange**
- 預設只清內容與公式，保留格式；可選只清格式或全部清除。範圍會縮到「含格式」的使用區域，所以清整張表也很快。

## 資料表示

```csharp
public sealed record RangeData(
    string Sheet,
    string Range,               // 實際回傳的範圍（已限縮在有內容的區域內）
    string? UsedRange,          // 工作表目前有內容的範圍
    object?[][] Values,         // 數字/字串/布林/null，日期轉 ISO 字串，公式錯誤是 "#DIV/0!" 等文字
    string?[][]? Formulas,      // 只有 IncludeFormulas 時才回
    bool Truncated,
    string? NextRange,          // 被截斷時，下一段要讀的範圍
    IReadOnlyList<string> CalculationWarnings);
```

- 讀取回傳緊湊的 2D 陣列；可選擇附公式（`IncludeFormulas`）、顯示文字（`UseFormattedText`）。
- 超過 `MaxCellsPerRead` 就截斷，回傳 `Truncated` 與 `NextRange`。
- 寫入時 `=` 開頭當公式、ISO 日期字串轉日期，兩者都可關閉；`null` 代表清空。
- 公式計算失敗時回傳快取值並標記 `CalculationWarnings`，不讓整個操作失敗。實測 ClosedXML 0.105 讀取公式時會丟例外的情況：語法錯誤（`ExpressionParseException`）、循環參照（`InvalidOperationException`）、不支援的運算（例如範圍交集，`NotImplementedException`），讀取端要逐格捕捉。
- **寫入公式要先驗證語法**：實測 `FormulaA1 = "SUM(A1:"` 設定與存檔都不會報錯，壞公式會被寫進檔案、Excel 開啟時要求修復。寫入前先用 `ClosedXML.Parser` 解析，語法錯誤回 `INVALID_VALUE`（附錯誤位置），不寫入。未知函式名稱（`NOSUCHFN(1)`）不會被擋，會得到 `#NAME?`，這是 Excel 本身的行為，不視為錯誤。
- 以字串寫入時，ISO 日期字串預設轉成日期（可用 `ParseIsoDates` 關閉）；`00123` 這類看起來像數字的字串一律保持文字。
- `SHEET_NOT_FOUND` 錯誤直接附上現有工作表清單，省一輪重試。
