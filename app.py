import streamlit as st
import yfinance as yf
import pandas as pd
import io
import unicodedata
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# --- 輔助函式：計算字串視覺寬度 ---
def get_text_width(text):
    """精準計算視覺寬度：中文字/全形算 2.2，英數字算 1.2"""
    width = 0
    for char in str(text):
        # 判斷是否為全形或東亞字元
        if unicodedata.east_asian_width(char) in ('F', 'W'):
            width += 2.2 
        else:
            width += 1.2
    return width

# --- 頁面基本設定 ---
st.set_page_config(page_title="美股估值追蹤儀表板", layout="wide")

st.title("📈 指定美股核心估值與 EPS 追蹤系統")
st.markdown("點擊下方按鈕獲取最新即時數據，抓取完成後即可預覽並下載 Excel 檔案。")

# 1. 準備你的股票代號清單
tickers_list = [
    'AAL', 'AAPL', 'AMD', 'AMZN', 'ARM', 'ASML', 'AVGO', 'BABA', 'BAC', 'BIIB',
    'BNTX', 'BP', 'CCL', 'COIN', 'DELL', 'DIS', 'ENPH', 'ENVX', 'INTC', 'JPM',
    'META', 'MRNA', 'MRVL', 'MSFT', 'MU', 'NVAX', 'NVDA', 'ON', 'ORCL', 'OXY', 'PLTR',
    'PLUG', 'QCOM', 'RGTI', 'SEDG', 'SHOP', 'SMCI', 'SOUN', 'SWKS', 'TEM', 'TSLA',
    'TSM', 'UAL', 'XYF'
]

# 點擊按鈕後開始執行
if st.button("🚀 開始抓取最新數據", type="primary"):
    stock_data = []
    total_tickers = len(tickers_list)
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # 2. 透過迴圈抓取股票
    for i, ticker in enumerate(tickers_list):
        status_text.text(f"⏳ 正在處理: {ticker} ({i+1}/{total_tickers})...")
        
        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            
            # 3. 提取核心指標
            stock_data.append({
                '代號': ticker,
                '公司名稱': info.get('longName', info.get('shortName', 'N/A')),
                '目前價格': info.get('currentPrice', 'N/A'),
                '本益比 (TTM)': info.get('trailingPE', 'N/A'),
                '前瞻本益比': info.get('forwardPE', 'N/A'),
                'EPS (TTM)': info.get('trailingEps', 'N/A')
            })
            
        except Exception as e:
            st.toast(f"❌ 抓取 {ticker} 時發生錯誤: {e}")
            
        progress_bar.progress((i + 1) / total_tickers)
        
    status_text.success("✅ 抓取完成！")
    
    # 4. 將資料轉換為 DataFrame 並顯示在網頁上
    df = pd.DataFrame(stock_data)
    st.dataframe(df, use_container_width=True)
    
    # 5. 將 DataFrame 寫入記憶體緩衝區並進行 Excel 排版格式化
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Stock Data')
        
        # 取得 openpyxl 的工作簿與工作表物件
        workbook = writer.book
        worksheet = writer.sheets['Stock Data']
        
        # 設定字型為微軟正黑體
        ms_font = Font(name='微軟正黑體')
        
        # 逐欄設定寬度與字型
        for i, col_name in enumerate(df.columns):
            col_letter = get_column_letter(i + 1) # 取得欄位字母 (例如 A, B, C)
            
            # 尋找該欄位中最長的視覺寬度 (比較標題列與所有資料列)
            max_width = get_text_width(col_name)
            for cell_value in df[col_name]:
                current_width = get_text_width(cell_value)
                if current_width > max_width:
                    max_width = current_width
            
            # 設定最佳欄寬 (額外加 2 的緩衝空間讓版面更舒服)
            worksheet.column_dimensions[col_letter].width = max_width + 2
            
            # 將整欄的每一個儲存格套用微軟正黑體
            for cell in worksheet[col_letter]:
                cell.font = ms_font
        
    # 6. 建立 Excel 下載按鈕
    st.download_button(
        label="📥 下載資料為 Excel 檔案",
        data=excel_buffer.getvalue(),
        file_name="ELI_stock_EPS.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )