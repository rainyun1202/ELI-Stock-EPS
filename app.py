import streamlit as st
import yfinance as yf
import pandas as pd
import io

# 頁面基本設定
st.set_page_config(page_title="美股估值追蹤儀表板", layout="wide")

st.title("📈 美股核心估值與 EPS 追蹤系統")
st.markdown("點擊下方按鈕獲取最新即時數據，抓取完成後即可預覽並下載為 Excel 檔案。")

# 1. 準備你的股票代號清單
tickers_list = [
    'AAL', 'AAPL', 'AMD', 'AMZN', 'ARM', 'ASML', 'AVGO', 'BABA', 'BAC', 'BIIB',
    'BNTX', 'BP', 'CCL', 'COIN', 'DELL', 'DIS', 'ENPH', 'ENVX', 'INTC', 'JPM',
    'META', 'MRNA', 'MSFT', 'MU', 'NVAX', 'NVDA', 'ON', 'ORCL', 'OXY', 'PLTR',
    'PLUG', 'QCOM', 'RGTI', 'SEDG', 'SHOP', 'SMCI', 'SOUN', 'SWKS', 'TEM', 'TSLA',
    'TSM', 'UAL', 'XYF'
]

# 點擊按鈕後開始執行
if st.button("🚀 開始抓取最新數據", type="primary"):
    stock_data = []
    total_tickers = len(tickers_list)
    
    # 建立進度條與狀態文字的 UI 佔位符
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
            # 使用 toast 顯示非阻擋式的錯誤通知
            st.toast(f"❌ 抓取 {ticker} 時發生錯誤: {e}")
            
        # 更新進度條比例
        progress_bar.progress((i + 1) / total_tickers)
        
    status_text.success("✅ 抓取完成！")
    
    # 4. 將資料轉換為 DataFrame 並顯示在網頁上
    df = pd.DataFrame(stock_data)
    st.dataframe(df, use_container_width=True)
    
    # 5. 將 DataFrame 寫入記憶體緩衝區 (取代直接存檔硬碟)
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Stock Data')
        
    # 6. 建立 Excel 下載按鈕
    st.download_button(
        label="📥 下載資料為 Excel 檔案 (ELI_stock_EPS.xlsx)",
        data=excel_buffer.getvalue(),
        file_name="ELI_stock_EPS.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )