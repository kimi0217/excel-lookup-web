# excel-lookup-web

內網唯讀查詢網頁:讀取資料夾(含子資料夾)中的 Excel,每個檔案(.xlsx、.xlsm,副檔名不分大小寫)= 一個分類,提供搜尋。

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
- 啟動:python app.py → http://localhost:8000(開發用;給使用者的是「啟動.bat」/「啟動.command」)
- 測試:pytest
- 產生各種格式的測試檔:python tests/fixtures/generate_fixtures.py(輸出到 tests/fixtures/generated/,不進版控)

## 架構
- `app.py`:FastAPI、設定載入、背景工作(預先載入、每日備份)
- `excel_reader.py`:安全讀檔(暫存複本)、讀取失敗原因判斷、`Catalog`(預先載入與搜尋)
- `table_layout.py`:表頭偵測、合併儲存格填值、欄位名稱整理(純邏輯,不碰檔案)
- `cell_format.py`:儲存格值 → 顯示文字(日期、數字格式、貨幣),搜尋用正規化
- `backup.py`:每日備份到 backups/YYYY-MM-DD/(保留子資料夾結構),保留 backup_keep_days 天
- `deploy/`:給不懂程式的使用者的部署腳本,由根目錄的 `.bat`(Windows)/`.command`(macOS)呼叫
  - `common.py`:共用(內網 IP、連接埠檢查、資料夾選擇、config.json 產生);安裝前就會執行,只能用標準函式庫
  - `install.py`:建立 venv、安裝套件,依 pip 輸出判斷失敗原因
  - `launcher.py`:第一次設定(選資料夾產生 config.json)、啟動前檢查、顯示網址;`--background` 為自動啟動用
  - `autostart.py`:Windows 工作排程器(XML 建立)/ macOS LaunchAgent

## API 與前端
- `GET /api/categories`:檔案清單(狀態、筆數、工作表數、檔案修改時間)
- `GET /api/files/{name}`:單一檔案摘要與工作表清單(name 可含 /,路由用 `{name:path}`,`/rows` 路由先宣告;子資料夾中名為 rows 的檔案由 rows 路由轉給 file_info)
- `GET /api/files/{name}/rows?sheet=&q=&sort=&order=&page=&page_size=`:工作表分頁資料(預設每頁 100,上限 500);篩選、排序、分頁都在伺服器端,不可一次送出整份大檔。排序:數字(含貨幣、千分位)依數值,空白永遠排最後;各欄排序結果快取在 Sheet._sort_cache
- `GET /api/search?q=`:跨所有檔案搜尋(最多 500 筆)
- 首頁依資料夾分組(static/index.html 的 groupByFolder),檔名篩選比對含資料夾的完整名稱
- 前端以 hash 路由:`#/` 檔案列表、`#/file/<檔名>?sheet=&q=&sort=&order=&page=` 資料檢視、`#/search?q=` 跨檔搜尋。檔案內的搜尋、排序、換頁用 history.replaceState 更新網址(不新增瀏覽紀錄),上一頁回到檔案列表

## 部署腳本的規則
- 所有畫面訊息用繁體中文,出錯時說明原因與處理方式(common.print_error),不要只顯示英文錯誤
- `.bat` 必須 UTF-8 無 BOM、CRLF 換行,前兩行為 `@echo off`、`chcp 65001 >nul`;中文只能出現在 echo/rem/title,echo 訊息用全形括號(半形括號在 if 區塊中會提早結束區塊)。.gitattributes 強制換行格式,tests/test_deploy.py 會檢查
- `.bat` 只負責找到 Python 並呼叫 deploy/ 的腳本,邏輯一律寫在 Python
- 自動啟動以目前使用者「登入後」執行(讀得到網路共用資料夾、不需密碼);工作排程器一定要用 XML 建立,才能取消預設的 3 天執行上限與電池限制

## 設計決定
- **讀檔用 openpyxl 唯讀模式,不用 pandas.read_excel**:需要每個儲存格的數字格式(還原開頭 0、貨幣、千分位)與工作表可見狀態,pandas 拿不到。pandas 只用於搜尋(Series.str.contains)。
- **公式讀 Excel 存下的計算結果**(data_only=True):沒經過 Excel 存檔的檔案(例如程式產生的)公式欄會是空白。
- **預先載入**:啟動時與每 reload_interval_seconds 秒檢查檔案修改時間與大小,在背景只重讀有變動的檔案(小檔優先)。搜尋絕不讀檔;檔案重新載入完成前沿用舊資料;從未載入過的檔案回報為 loading。
- **檔案類型**:讀 data_dir 中的 .xlsx、.xlsm(副檔名不分大小寫,excel_reader.is_excel_file);include_subfolders(預設 true)時遞迴讀取子資料夾(excel_reader.find_excel_files),略過 . 與 ~ 開頭的資料夾、符號連結的資料夾與沒有權限的子資料夾。分類名稱為相對路徑去掉副檔名(「醫材/耗材清單」);同一資料夾中同名不同副檔名時用完整檔名。備份使用相同規則並保留資料夾結構。
- **備份資料夾在 data_dir 底下**:Catalog 與備份都以 exclude=[backup_dir] 排除,絕不可把備份當成資料;backup_dir 就是 data_dir 時略過最上層的日期資料夾。
- **壞檔隔離**:任何讀取失敗只影響該檔,回報簡短中文原因(密碼保護、舊版 .xls、損毀、無法存取),不可讓「全部分類」搜尋失敗。讀取失敗的檔案每次 refresh 都重讀(同樣的錯誤只記錄一次)。
- **資料夾斷線**:scan_folder 失敗時保留已載入的資料,/api/categories 的 folder 回報 ok 與 last_success;請求一律只用背景 refresh 快取的檔案清單,不存取資料夾(網路斷線時存取可能卡住數秒)。
- **隱藏工作表預設不讀**:隱藏與極度隱藏的工作表在 include_hidden_sheets 為 false(預設)時完全不讀取,避免使用者刻意隱藏的資料被搜尋到。
- **隱藏的列與欄**(include_hidden_rows_columns,預設 false):false 時手動隱藏的欄與列不讀(table_layout.drop_hidden,在偵測表頭之前移除);位於篩選範圍內的隱藏列是篩選造成的,照常讀取。篩選範圍包含工作表的 `<autoFilter>` 與「格式化為表格」各表格 XML 中的 `<autoFilter>`(透過工作表的 _rels 找到 xl/tables/*.xml)。唯讀模式取不到這些資訊,與合併儲存格相同從工作表 XML 解析 `<row hidden="1">`、`<col min max hidden="1">`;每個工作表的 XML 只讀一次。
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
- openpyxl 唯讀模式不提供合併範圍與隱藏列欄,程式透過內部屬性 `_worksheet_path`(excel_reader.SHEET_PATH_ATTR)找到工作表 XML,自行解析 `<mergeCell>`、`<row hidden>`、`<col hidden>`、`<autoFilter>`。
- 這是非公開 API,升級 openpyxl 後可能改名。取不到時程式不會當掉,但合併儲存格不會填值、隱藏的列與欄也會被讀取;啟動時會檢查(merge_support_ok)並在記錄中印出警告。
- tests/test_merge_support.py 會在取不到合併範圍時失敗,升級 openpyxl 前務必先跑測試。
