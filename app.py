import streamlit as st
import yfinance as yf
import pandas as pd
import io
import unicodedata
from datetime import datetime
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# --- 輔助函式：計算字串視覺寬度 ---
def get_text_width(text):
    """精準計算視覺寬度：中文字/全形算 2.2，英數字算 1.2"""
    width = 0
    for char in str(text):
        if unicodedata.east_asian_width(char) in ('F', 'W'):
            width += 2.2 
        else:
            width += 1.2
    return width

# --- 輔助函式：數值格式化 (轉百分比或小數點) ---
def format_percent(val):
    if val is None or pd.isna(val):
        return "N/A"
    try:
        return f"{float(val) * 100:.2f}%"
    except:
        return "N/A"

def format_float(val):
    if val is None or pd.isna(val):
        return "N/A"
    try:
        return round(float(val), 2)
    except:
        return "N/A"

# --- 輔助函式：安全提取財報日期 ---
def get_earnings_date(stock_info, ticker_obj):
    """嘗試多種方法從 yfinance 提取下一次財報發布日"""
    try:
        # 方法 1: 從 info 中的 UNIX timestamp 提取 (最穩定)
        timestamp = stock_info.get('earningsTimestamp') or stock_info.get('earningsTimestampStart')
        if timestamp:
            return datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d')
            
        # 方法 2: 從 calendar 模組提取
        cal = ticker_obj.calendar
        if isinstance(cal, dict) and 'Earnings Date' in cal:
            dates = cal['Earnings Date']
            if dates and len(dates) > 0:
                return dates[0].strftime('%Y-%m-%d')
    except:
        pass
    return "未知 / N/A"

# --- 頁面基本設定 ---
st.set_page_config(page_title="指定美股基本面追蹤系統", layout="wide")

st.title("🚀 美股基本面與財報事件追蹤系統")
st.markdown("""
點擊下方按鈕獲取最新即時數據，抓取完成後即可預覽並下載 Excel 檔案。
""")

# 1. 準備股票代號清單
tickers_list = [
    'AAL', 'AAPL', 'AMD', 'AMAT', 'AMZN', 'ARM', 'ASML', 'AVGO', 'BABA', 'BAC', 'BIIB',
    'BNTX', 'BP', 'CCL', 'COIN', 'DELL', 'DIS', 'ENPH', 'ENVX', 'INTC', 'JPM', 'KLAC',
    'LRCX', 'META', 'MRNA', 'MSFT', 'MU', 'NVAX', 'NVDA', 'ON', 'ORCL', 'OXY', 'PLTR',
    'PLUG', 'QCOM', 'RGTI', 'SEDG', 'SHOP', 'SMCI', 'SOUN', 'SWKS', 'TEM', 'TSLA',
    'TSM', 'UAL', 'XYF'
]

if st.button("📊 開始抓取數據", type="primary"):
    stock_data = []
    total_tickers = len(tickers_list)
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, ticker in enumerate(tickers_list):
        status_text.text(f"⏳ 正在處理: {ticker} ({i+1}/{total_tickers}) - 獲取數據中...")
        
        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            
            # --- 提取核心與深度指標 ---
            stock_data.append({
                '代號': ticker,
                '公司名稱': info.get('longName', info.get('shortName', 'N/A')),
                '最新股價': format_float(info.get('currentPrice')),
                
                # 財報事件
                '下次財報發布日': get_earnings_date(info, stock),
                
                # 估值指標 (Valuation)
                '本益比 (TTM)': format_float(info.get('trailingPE')),
                '前瞻本益比': format_float(info.get('forwardPE')),
                '股價淨值比 (P/B)': format_float(info.get('priceToBook')),
                'EPS (TTM)': format_float(info.get('trailingEps')),
                
                # 成長與獲利能力 (Growth & Profitability)
                '營收年增率 (YoY)': format_percent(info.get('revenueGrowth')),
                '盈餘年增率 (YoY)': format_percent(info.get('earningsGrowth')),
                '股東權益報酬率 (ROE)': format_percent(info.get('returnOnEquity')),
                '營業利益率 (Margin)': format_percent(info.get('operatingMargins'))
            })
            
        except Exception as e:
            st.toast(f"❌ 抓取 {ticker} 時發生錯誤: {e}")
            
        progress_bar.progress((i + 1) / total_tickers)
        
    status_text.success("✅ 數據抓取完成！")
    
    # 4. 轉換為 DataFrame 並顯示在網頁上
    df = pd.DataFrame(stock_data)
    st.dataframe(df, use_container_width=True)
    
    # 5. 將 DataFrame 寫入記憶體緩衝區並進行進階 Excel 排版格式化
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Deep Fundamentals')
        
        workbook = writer.book
        worksheet = writer.sheets['Deep Fundamentals']
        
        ms_font = Font(name='微軟正黑體')
        # 標題列粗體
        header_font = Font(name='微軟正黑體', bold=True, color='FFFFFF')
        
        for i, col_name in enumerate(df.columns):
            col_letter = get_column_letter(i + 1)
            
            # 尋找最長寬度
            max_width = get_text_width(col_name)
            for cell_value in df[col_name]:
                current_width = get_text_width(cell_value)
                if current_width > max_width:
                    max_width = current_width
            
            worksheet.column_dimensions[col_letter].width = max_width + 3
            
            # 格式化儲存格
            for row_idx, cell in enumerate(worksheet[col_letter], 1):
                cell.font = ms_font
                # 將數值與百分比置中對齊，看起來更整齊
                if row_idx > 1 and cell.value != "N/A" and col_name not in ['代號', '公司名稱']:
                    cell.alignment = Alignment(horizontal='right')
                
                # 給標題列特別的樣式 (如果你想要的話可以開啟)
                # if row_idx == 1:
                #     cell.font = header_font
        
    # 6. 建立 Excel 下載按鈕
    st.download_button(
        label="📥 下載報表 (Excel)",
        data=excel_buffer.getvalue(),
        file_name="ELI_Stock.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )