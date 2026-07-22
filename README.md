以下為「異常圖片彙整專案」投影片內容的彙整文字（繁體中文、台灣用語）：

專案簡述
目標：當 單站偵測到異常（以 Lot / Unit 為查詢主鍵）時，自動由各站（SPI2 / AOI3 / AOI4 / 3DL / FT1）回追並集中輸出相關影像，降低人工找圖時間與漏抓風險。
概念：分散讀取各設備可見影像、集中輸出到共用 NAS/雲端資料夾，由中央端再產生彙整報表。
核心流程（簡要步驟）
單站發現異常，填入 query_list.xlsx（以 Lot_No、Unit_No、Status = Open/New 為處理條件）。
各站 Collector 讀取 query_list.xlsx（站點依各自的 config 判斷要抓哪些圖）。
Collector 在本機搜尋符合規則的圖片，複製至共用 Output，保留原始檔名（先寫入 .part，完成後改正式檔名）。
中央端或報表程式依 Output 建立彙整報表（FOUND/NOT_FOUND/ERROR 等狀態），供人工驗證及下載。
架構與主要檔案
Query_Input：query_list.xlsx（唯一入口）
各站 Collector（可獨立部署）
以 station_config.json 控制來源路徑、圖種、比對規則、資源上限
Output / Log：NAS 共用位置儲存圖片與 log（Output/{Lot_No}/{Station}/）
站點差異化規則（示例）
SPI2：元件圖、整版圖、熱力圖；按 Unit / Lot；OK/NG 分路徑，BOARD 以檔名包含 Lot 判斷
AOI3：元件圖、Map；Before Reflow；以 AOI003 path_include 判定
AOI4：元件圖、Map；After Reflow；以 AOI004 path_include 判定
3DL：2D/3D 元件圖、Defect Map；尋找 Unit.tif、Unit_3D.tif、DefectMap_Die.jpg
FT1：H_DefectMap、Mapping_H；以 Lot folder 與檔名尾碼抓取（by lot 圖種允許 Unit 空白）
輸出檔名與結構設計
資料夾簡化，關鍵資訊放在檔名與 log 中，範例：
Output/{Lot_No}/{Station}/
範例檔名：A326689B90900804_846_SPI2_COMPONENT_OK_20260701160339_846_1.jpg
設計原則：避免資料夾過深，方便報表或人工檢視。
穩定性與容錯設計
複製安全：先寫 .part 暫存檔，完成再改名；若 Output 遺失，下一輪補傳。
資源控管：每輪最多 50 次來源搜尋；已存在 Output 的項目跳過；長清單分輪處理。
常駐監控：single_instance 防重複執行；status_file heartbeat；log / done_record 自動封存。


預期效益
效率：單站異常後不用逐站人工找圖、長清單可背景分輪處理、已完成項目不重複掃描來源。 
人工撈取每片每個站點耗時5分鐘，最大產能A+ 良率85%, 每天處理量約90EA  

協助製作PPT 
