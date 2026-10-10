# excel-lookup-web

內網唯讀查詢網頁:讀取資料夾中的 Excel,每個 .xlsx = 一個分類,提供搜尋。

## 原則
- 唯讀:程式絕不寫入或修改原始 Excel
- 讀檔一律先複製到暫存,讀完立即刪除
- 必須在 macOS 與 Windows 都能執行(路徑用 pathlib,檔案讀寫指定 utf-8)
- 使用者介面文字一律使用繁體中文
- 前端不用框架,保持單一 HTML 檔

## 技術
- Python 3.11+(pandas 3 需要 3.11)、FastAPI、uvicorn、pandas、openpyxl
- requirements.txt 全部用 == 鎖定版本,升級任何套件前先跑 pytest
- 測試:pytest

## 指令
- 啟動:python app.py → http://localhost:8000
- 測試:pytest
- 產生各種格式的測試檔:python tests/fixtures/generate_fixtures.py(輸出到 tests/fixtures/generated/,不進版控)

## 架構
- `app.py`:FastAPI、設定載入、背景工作(預先載入、每日備份)
- `excel_reader.py`:安全讀檔(暫存複本)、讀取失敗原因判斷、`Catalog`(預先載入與搜尋)
- `table_layout.py`:表頭偵測、合併儲存格填值、欄位名稱整理(純邏輯,不碰檔案)
- `cell_format.py`:儲存格值 → 顯示文字(日期、數字格式、貨幣),搜尋用正規化
- `backup.py`:每日備份到 backups/YYYY-MM-DD/,保留 backup_keep_days 天

## 設計決定
- **讀檔用 openpyxl 唯讀模式,不用 pandas.read_excel**:需要每個儲存格的數字格式(還原開頭 0、貨幣、千分位)與工作表可見狀態,pandas 拿不到。pandas 只用於搜尋(Series.str.contains)。
- **公式讀 Excel 存下的計算結果**(data_only=True):沒經過 Excel 存檔的檔案(例如程式產生的)公式欄會是空白。
- **預先載入**:啟動時與每 reload_interval_seconds 秒檢查檔案修改時間與大小,在背景只重讀有變動的檔案(小檔優先)。搜尋絕不讀檔;檔案重新載入完成前沿用舊資料;從未載入過的檔案回報為 loading。
- **壞檔隔離**:任何讀取失敗只影響該檔,回報簡短中文原因(密碼保護、舊版 .xls、損毀、無法存取),不可讓「全部分類」搜尋失敗。
- **隱藏工作表預設不讀**:隱藏與極度隱藏的工作表在 include_hidden_sheets 為 false(預設)時完全不讀取,避免使用者刻意隱藏的資料被搜尋到。
- **表頭偵測**(table_layout.detect_header):在前 10 個非空白列中,選最前面一個同時符合下列條件的列;都不符合時用第一個非空白列:
  - 非空白格數(合併範圍算每一格)≥ 候選列最大值的 80%
  - 非空白格中文字超過一半(數字、日期不算文字)
  - 至少 2 個不同的值(排除跨整列合併的標題),整張表只有 1 欄時例外
  - 下面還有資料列
- 表頭上方的列(標題、說明)不列入資料,也搜尋不到;全空白的列與欄略過。
- **合併儲存格**:只有起點在表頭下方的合併範圍會把值填入每一格;表頭的合併不處理,不支援多層表頭。
- **欄位名稱**:重複的顯示為「電話」「電話(2)」,空白的顯示為「欄位N」(N 為顯示後的欄位序號)。
- **搜尋規則**:不分大小寫、/ 與 - 視為相同、忽略數字間的千分位逗號;儲存格之間以分隔字元串接,不會跨儲存格比對。前端醒目標示的規則需與後端一致(static/index.html 的 highlightPattern)。
- **數字格式**:支援補 0、小數位數、千分位、百分比、貨幣與文字前後綴、正負零分段;分數與科學記號等無法處理的格式以一般格式顯示。

## 風險:openpyxl 內部屬性
- openpyxl 唯讀模式不提供合併範圍,程式透過內部屬性 `_worksheet_path`(excel_reader.SHEET_PATH_ATTR)找到工作表 XML,自行解析 `<mergeCell>`。
- 這是非公開 API,升級 openpyxl 後可能改名。取不到時程式不會當掉,但合併儲存格不會填值;啟動時會檢查(merge_support_ok)並在記錄中印出警告。
- tests/test_merge_support.py 會在取不到合併範圍時失敗,升級 openpyxl 前務必先跑測試。
