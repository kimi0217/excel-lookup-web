# excel-lookup-web

內網唯讀查詢網頁:讀取資料夾中的 Excel,每個 .xlsx = 一個分類,提供搜尋。

## 原則
- 唯讀:程式絕不寫入或修改原始 Excel
- 讀檔一律先複製到暫存,讀完立即刪除
- 必須在 macOS 與 Windows 都能執行(路徑用 pathlib,檔案讀寫指定 utf-8)
- 使用者介面文字一律使用繁體中文
- 前端不用框架,保持單一 HTML 檔

## 技術
- Python 3.10+、FastAPI、uvicorn、pandas、openpyxl
- 測試:pytest

## 指令
- 啟動:python app.py → http://localhost:8000
- 測試:pytest