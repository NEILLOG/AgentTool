# 06 暫不做與未來選項

| 項目 | 原因 | 之後怎麼補 |
| --- | --- | --- |
| Word / PPT 編輯 | 流式文件與空間版面難定址，容易改錯或排版難看 | Word 用段落穩定 ID 編輯；PPT 以範本版面與預留位置為主 |
| Excel 樞紐、Power Query、VBA | ClosedXML 做不到，需要真的 Excel | 需要時另用 COM 方案，只限 Windows 桌面環境 |
| Excel 圖表 | ClosedXML 支援有限 | 第二版評估 |
| .xlsm 寫入、保留圖表/樞紐的就地存檔 | ClosedXML 存檔會遺失 VBA、圖表、樞紐 | 第一版只能另存；之後評估改用 Open XML SDK 直接改儲存格部件 |
| 舊格式 .xls / .doc / .ppt | Open XML 不支援 | 回錯誤請使用者另存新格式 |
| 複雜 PDF 表格辨識 | .NET 沒有對等的版面模型 | 先用多模態模型看圖；仍不夠再把 Docling 包成 MCP server |
| MCP server | 目前內建即可 | 同一個 library 加一層 MCP adapter |
| 遠端 API（IIS / k8s） | 檔案在使用者本機，走網路只增加成本 | 檔案集中在伺服器或需多人共用時再做 |
