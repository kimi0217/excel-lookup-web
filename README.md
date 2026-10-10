# Excel 查詢網頁

內網使用的唯讀 Excel 查詢網頁。把 .xlsx 檔放在一個資料夾,每個檔案就是一個分類,同事用瀏覽器就能搜尋所有檔案的內容。

- **唯讀**:程式只讀取複本,絕不修改原始 Excel。
- **自動更新**:檔案新增、修改、刪除後,約 10 秒內自動反映,不必重新啟動。
- **每日備份**:每天把所有 Excel 複製一份到備份資料夾,保留 30 天。

> ⚠️ **沒有登入或權限控管**。能連到這台電腦的人都看得到資料夾中所有 Excel 的內容(隱藏工作表除外)。請只放可以公開給內網的資料,並只在內網使用。

---

## 安裝

需要 **Python 3.11 以上**。

### macOS

1. 安裝 Python:到 <https://www.python.org/downloads/> 下載 macOS 安裝檔並安裝。
2. 把專案資料夾放到固定位置,建議 **不要放在「桌面」「文件」「下載項目」**(原因見[開機自動啟動](#macos-開機自動啟動))。以下以 `/Users/Shared/excel-lookup-web` 為例。
3. 開啟「終端機」,執行:

   ```bash
   cd /Users/Shared/excel-lookup-web
   python3 -m venv venv
   venv/bin/pip install -r requirements.txt
   cp config.example.json config.json
   ```

4. 用文字編輯器修改 `config.json`(欄位說明見[設定](#設定-configjson)),至少要改 `data_dir`。
5. 確認安裝正確(全部通過才繼續):

   ```bash
   venv/bin/pytest
   ```

6. 啟動:

   ```bash
   venv/bin/python app.py
   ```

7. 用瀏覽器開 <http://localhost:8000>。其他電腦用 `http://這台電腦的IP:8000`(IP 在「系統設定 → 網路」查看)。

### Windows

1. 安裝 Python:到 <https://www.python.org/downloads/> 下載 Windows 安裝檔。安裝時 **勾選「Add python.exe to PATH」**。
2. 把專案資料夾放到固定位置,以下以 `C:\excel-lookup-web` 為例。
3. 開啟「命令提示字元」(cmd),執行:

   ```bat
   cd /d C:\excel-lookup-web
   py -3 -m venv venv
   venv\Scripts\python.exe -m pip install -r requirements.txt
   copy config.example.json config.json
   ```

   > 直接呼叫 `venv\Scripts\python.exe`,不需要「啟用」虛擬環境,也不會遇到 PowerShell 執行原則的限制。

4. 用記事本修改 `config.json`,至少要改 `data_dir`。路徑寫法見[Windows 路徑寫法](#windows-路徑寫法)。
5. 確認安裝正確(全部通過才繼續):

   ```bat
   venv\Scripts\python.exe -m pytest
   ```

6. 啟動:

   ```bat
   venv\Scripts\python.exe app.py
   ```

7. 用瀏覽器開 <http://localhost:8000>。其他電腦用 `http://這台電腦的IP:8000`(IP 用 `ipconfig` 查看「IPv4 位址」)。

---

## 設定(config.json)

`config.json` 含真實路徑,不會進版控。找不到 `config.json` 時會改用 `config.example.json`。修改後需重新啟動程式。

```json
{
  "data_dir": "./data",
  "backup_dir": "./backups",
  "backup_keep_days": 30,
  "host": "0.0.0.0",
  "port": 8000,
  "max_results": 500,
  "include_hidden_sheets": false,
  "reload_interval_seconds": 10
}
```

| 欄位 | 預設值 | 說明 |
|---|---|---|
| `data_dir` | `./data` | 放 Excel 的資料夾。只讀取這一層的 `.xlsx`,不含子資料夾。相對路徑以專案資料夾為準。程式只需要這個資料夾的 **讀取** 權限,建議只給讀取權限。 |
| `backup_dir` | `./backups` | 每日備份存放位置,每天一個 `YYYY-MM-DD` 資料夾。建議放在與 `data_dir` 不同的磁碟。備份一律啟用。 |
| `backup_keep_days` | `30` | 備份保留天數,超過的日期資料夾會自動刪除(只刪日期格式命名的資料夾)。備份空間約為「所有 Excel 總大小 × 保留天數」。 |
| `host` | `0.0.0.0` | `0.0.0.0` 讓其他電腦可以連線;`127.0.0.1` 只有本機能用。 |
| `port` | `8000` | 網址中的連接埠。若改了,防火牆規則也要跟著改。 |
| `max_results` | `500` | 每次搜尋最多回傳幾筆(所有檔案合計)。 |
| `include_hidden_sheets` | `false` | 是否讀取隱藏與極度隱藏的工作表。`false` 時完全不讀,搜尋不到。 |
| `reload_interval_seconds` | `10` | 多久檢查一次檔案是否有變動(秒)。只重新讀取有變動的檔案。 |

### Windows 路徑寫法

JSON 中的反斜線 `\` 必須寫兩次,或改用正斜線 `/`(兩種都可以):

```json
"data_dir": "D:\\共用資料\\Excel"
"data_dir": "D:/共用資料/Excel"
```

網路芳鄰的共用資料夾(UNC 路徑):

```json
"data_dir": "\\\\fileserver\\共用\\Excel"
"data_dir": "//fileserver/共用/Excel"
```

> 若要開機自動啟動,**請用 UNC 路徑,不要用對應的磁碟機代號**(例如 `Z:`)。磁碟機代號只在使用者登入後才存在,開機時啟動的程式看不到。

只寫一個反斜線(例如 `"D:\共用"`)會讓程式啟動失敗並顯示 JSON 格式錯誤。

---

## 防火牆

其他電腦要連線,需要允許這台電腦的 TCP 8000 埠(或你在 `port` 設定的埠)傳入連線。

### macOS

第一次啟動時若出現「是否允許 Python 接受傳入的網路連線?」,按 **允許**。

若沒有出現或不小心按了拒絕:「系統設定 → 網路 → 防火牆 → 選項」,找到 Python 並設為「允許傳入連線」。

### Windows

第一次啟動時若出現「Windows 安全性警訊」,勾選 **私人網路**(公司網域電腦也勾 **網域網路**),按「允許存取」。**不要勾公用網路。**

或以系統管理員身分開啟命令提示字元,直接新增規則:

```bat
netsh advfirewall firewall add rule name="Excel 查詢網頁" dir=in action=allow protocol=TCP localport=8000 profile=private,domain
```

---

## 開機自動啟動

### macOS 開機自動啟動

使用 launchd。以下設定會在開機時啟動,程式意外結束時自動重新啟動。

1. 建立 `/Library/LaunchDaemons/com.excel-lookup-web.plist`(需要管理員權限,例如 `sudo nano /Library/LaunchDaemons/com.excel-lookup-web.plist`),內容如下。把 `你的帳號` 換成要用來執行的 macOS 帳號,路徑換成實際位置:

   ```xml
   <?xml version="1.0" encoding="UTF-8"?>
   <!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
   <plist version="1.0">
   <dict>
     <key>Label</key>
     <string>com.excel-lookup-web</string>
     <key>UserName</key>
     <string>你的帳號</string>
     <key>ProgramArguments</key>
     <array>
       <string>/Users/Shared/excel-lookup-web/venv/bin/python</string>
       <string>/Users/Shared/excel-lookup-web/app.py</string>
     </array>
     <key>WorkingDirectory</key>
     <string>/Users/Shared/excel-lookup-web</string>
     <key>RunAtLoad</key>
     <true/>
     <key>KeepAlive</key>
     <true/>
     <key>StandardOutPath</key>
     <string>/Users/Shared/excel-lookup-web/app.log</string>
     <key>StandardErrorPath</key>
     <string>/Users/Shared/excel-lookup-web/app.log</string>
   </dict>
   </plist>
   ```

2. 啟用並立即啟動:

   ```bash
   sudo chown root:wheel /Library/LaunchDaemons/com.excel-lookup-web.plist
   sudo launchctl bootstrap system /Library/LaunchDaemons/com.excel-lookup-web.plist
   ```

3. 停用:`sudo launchctl bootout system/com.excel-lookup-web`

> **macOS 隱私權限制**:背景服務預設無法讀取「桌面」「文件」「下載項目」與外接磁碟。專案和 Excel 資料夾若放在這些位置,程式會讀不到檔案(分類清單是空的,或記錄中出現 `Operation not permitted`)。請把它們移到 `/Users/Shared/` 等其他位置;或到「系統設定 → 隱私權與安全性 → 完整磁碟取用權限」加入 `venv/bin/python` 實際指向的 Python 程式。

### Windows 開機自動啟動

使用「工作排程器」。

1. 在專案資料夾建立 `start.bat`,內容如下(會把記錄寫到 `app.log`):

   ```bat
   @echo off
   cd /d C:\excel-lookup-web
   venv\Scripts\python.exe app.py >> app.log 2>&1
   ```

2. 開啟「工作排程器」→「建立工作」(不是「建立基本工作」):
   - **一般**:名稱填「Excel 查詢網頁」;選「不論使用者是否登入都執行」;執行帳號選一個 **有權限讀取 Excel 資料夾** 的帳號(資料夾在網路共用時,不要用 SYSTEM)。
   - **觸發程序**:新增 →「啟動時」。
   - **動作**:新增 →「啟動程式」→ 選 `C:\excel-lookup-web\start.bat`,「開始位置」填 `C:\excel-lookup-web`。
   - **條件**:取消「只有在電腦使用 AC 電源時才啟動這個工作」。
   - **設定**:取消「如果工作執行時間超過下列時間,便停止工作」;勾選「如果工作失敗,每隔 1 分鐘重新啟動」。
3. 按確定並輸入該帳號密碼。在工作上按右鍵 →「執行」立即測試,用瀏覽器確認可以開啟。

---

## 啟動記錄

程式的記錄會顯示在終端機,開機自動啟動時則寫到 `app.log`。需要留意的訊息:

- `無法讀取 XXX.xlsx`:該檔案有密碼保護、已損毀或是舊版 .xls,網頁上也會顯示原因。其他檔案不受影響。
- `取不到合併儲存格範圍`:openpyxl 版本與 requirements.txt 不同,合併儲存格不會填入值。請重新執行 `pip install -r requirements.txt`。
- `自動備份失敗`:檢查 `backup_dir` 是否存在、是否有寫入權限、磁碟空間是否足夠。

---

## 更新程式

```bash
git pull
venv/bin/pip install -r requirements.txt     # Windows:venv\Scripts\python.exe -m pip install -r requirements.txt
venv/bin/pytest                              # Windows:venv\Scripts\python.exe -m pytest
```

測試全部通過後重新啟動程式。requirements.txt 中的套件版本都已鎖定,**請不要自行升級套件**(尤其是 openpyxl,見下方限制);要升級時先在測試環境跑過 pytest。

---

## Excel 檔案的處理方式

- 每個工作表自動找出表頭所在的列(前 10 個非空白列中,格數最多、大部分是文字、下面接著有資料的那一列)。表頭上方的標題、說明列不會顯示,也搜尋不到。
- 全空白的列與欄不顯示。
- 資料區的合併儲存格,範圍內每一格都填入合併格的值,例如合併的「部門」欄,同部門每一筆都搜尋得到。
- 重複的欄位名稱顯示為「電話」「電話(2)」,空白的欄位名稱顯示為「欄位3」。
- 日期顯示為 `2026-01-15`,有時間才顯示時間。搜尋時 `/` 與 `-` 視為相同。
- 數字依儲存格格式顯示:`00123`、`1,234`、`15%`、`NT$1,000`、`1,000 元` 等。搜尋時忽略千分位逗號,搜 `1000` 也找得到 `NT$1,000`。
- 檔案有密碼保護、已損毀或是舊版 .xls 時,下拉選單會標示原因,其他檔案照常搜尋。

## 已知限制

- **不支援多層表頭**:表頭有兩列以上(例如「人員」底下再分「姓名」「分機」)時,只有第一列當作欄位名稱,第二列會顯示成第一筆資料。
- **公式需要用 Excel 存過的檔案**:程式讀取的是 Excel 存檔時記錄的計算結果。由其他程式產生、沒有用 Excel 開啟存檔過的檔案,公式欄會是空白。遇到這種情況,用 Excel 開啟後存檔即可。
- **分數與科學記號格式**:顯示為一般數字(例如 `1 1/2` 顯示為 `1.5`,`1.23E+04` 顯示為 `12300`)。
- **只支援 .xlsx**:舊版 .xls 與 .xlsm 不會讀取,請另存為 .xlsx。
- **有密碼保護(開啟需要密碼)的檔案無法讀取**。只限制編輯的保護不影響。
- **表頭偵測可能判斷錯誤**:例如表頭有很多空白欄名、而資料列每格都有文字時,可能把第一筆資料當成表頭。遇到時可以在 Excel 中補齊欄位名稱。
- **啟動或檔案變動後的幾秒內**,背景正在讀取大檔時,網頁回應可能變慢到約 1 秒;尚未載入完成的檔案會在網頁上顯示「載入中」。
- **所有資料都載入記憶體**:Excel 總量很大時(數十萬列以上),請留意電腦記憶體。
- **合併儲存格依賴 openpyxl 內部屬性**:升級 openpyxl 後可能失效(記錄中會出現警告,測試也會失敗),因此版本已鎖定。
