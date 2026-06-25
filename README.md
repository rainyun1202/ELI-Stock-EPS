# 📈 ELI Stock EPS & Valuation Tracker (美股核心估值追蹤系統)

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-App-FF4B4B.svg)](https://streamlit.io/)
[![yfinance](https://img.shields.io/badge/yfinance-Data-green.svg)](https://pypi.org/project/yfinance/)

透過 Python 與 Streamlit 打造的輕量級資料視覺化與匯出工具，專為快速批量獲取美股核心基本面數據所設計。
一鍵抓取多檔股票的最新價格、本益比 (Trailing & Forward P/E) 以及每股盈餘 (EPS TTM)，
並自動匯出 Excel 報表。

🌐 **[https://eli-stock-eps.streamlit.app/]**

---

## ✨ 核心特色 (Features)

* 🚀 **批量快速自動抓取**：內建 40+ 檔熱門科技與藍籌股代號（可自由擴充），一鍵自動依序向伺服器請求數據，並提供即時的視覺化進度條反饋。
* 📊 **精準估值指標**：告別繁瑣的網頁查詢，直接提取投資決策最需要的關鍵數據：
    * 當前股價 (Current Price)
    * 本益比 (Trailing P/E)
    * 前瞻本益比 (Forward P/E)
    * 近四季每股盈餘 (Diluted EPS TTM)
* 📥 **智慧排版 Excel 匯出**：底層整合 `openpyxl` 引擎，匯出的 Excel 檔案會**自動適應文字寬度**（精準辨識中英文長度）並全面套用「微軟正黑體」，下載後零加工即可直接用於匯報。
* 🛡️ **穩定防封鎖架構**：棄用傳統易失效的網頁爬蟲 (Web Scraping)，全面改用 `yfinance` 串接官方級別資料流，完美避開 Cookie 驗證與反爬蟲阻擋。

---

## 💻 本地端安裝與執行 (Local Installation)

如果你想在自己的電腦上運行或修改此專案，請依照以下步驟操作：

**1. 複製儲存庫 (Clone the repository)**
```bash
git clone [https://github.com/rainyun1202/ELI-Stock-EPS.git](https://github.com/rainyun1202/ELI-Stock-EPS.git)
cd ELI-Stock-EPS

```

**2. 安裝必備套件 (Install dependencies)**
建議使用虛擬環境，然後執行：

```bash
pip install -r requirements.txt

```

**3. 啟動應用程式 (Run the app)**

```bash
streamlit run app.py

```

執行後，瀏覽器將自動開啟 `http://localhost:8501` 顯示應用程式介面。

---

## 🛠️ 技術棧 (Tech Stack)

* **前端介面**：[Streamlit](https://streamlit.io/)
* **數據來源**：[yfinance](https://pypi.org/project/yfinance/) (Yahoo Finance API 封裝)
* **資料處理**：[Pandas](https://pandas.pydata.org/)
* **Excel 處理**：[openpyxl](https://openpyxl.readthedocs.io/)

---

## 📝 如何自訂股票清單？

打開 `app.py`，尋找 `tickers_list` 變數。你可以隨意新增、刪除或修改裡面的股票代號字串，程式會在下次執行時自動讀取新的清單。

```python
tickers_list = [
    'AAPL', 'MSFT', 'NVDA', 'TSLA', 'YOUR_TICKER_HERE'
]

```

---

## 📄 授權條款 (License)

此專案採用 MIT 授權條款 - 詳情請參見 [LICENSE](https://www.google.com/search?q=LICENSE) 檔案。

```

***

**💡 幾個小建議：**
1. **補上截圖**：強烈建議你在 GitHub 頁面上傳一張這個網頁應用程式運作時的截圖（或是匯出的漂亮 Excel 截圖），然後在 `[點此體驗線上版本]` 下方加入語法 `![App Screenshot](你的圖片連結)`。人們通常是視覺動物，看到精美的截圖會大大增加點擊與使用的意願。
2. **替換連結**：記得把上方草稿中 `(這裡替換成你的 Streamlit 專屬網址)` 換成你實際發佈到 Streamlit Community Cloud 的 URL。

```