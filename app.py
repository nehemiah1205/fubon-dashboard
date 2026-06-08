import streamlit as st
import pandas as pd

# 網頁基本設定
st.set_page_config(page_title="戰情儀表板", layout="centered")
st.title("🚀 竹耀通訊處戰情儀表板")

# 建立檔案上傳區塊
uploaded_file = st.file_uploader("請上傳『當期通訊處排名-FYC.csv』", type=["csv"])

if uploaded_file:
    # 讀取 CSV 檔案。富邦的報表表頭有5行，我們跳過前5行直接讀取純數據
    df = pd.read_csv(uploaded_file, skiprows=5, header=None)
    
    # 在第 3 欄 (index 2) 尋找名稱為「竹耀」的資料列
    target_row = df[df[2] == '竹耀']
    
    if not target_row.empty:
        data = target_row.iloc[0]
        
        # 根據我們對報表的解析，鎖定對應的欄位索引來抓取數字
        month_target = float(data[5])   # 第6欄：當月目標
        month_actual = float(data[17])  # 第18欄：總核實FYC
        month_rate = float(data[18])    # 第19欄：當月總核實達成率
        
        year_target = float(data[6])    # 第7欄：累計目標
        year_actual = float(data[28])   # 第29欄：累計核實FYC
        year_rate = float(data[29])     # 第30欄：累計達成率
        
        st.success(f"✅ 成功載入 {data[2]}通訊處 (主管：{data[3]}) 的即時數據！")
        
        # ====== 1. 當月 FYC 達成進度 ======
        st.markdown("### 📊 當月 FYC 達成進度")
        col1, col2, col3 = st.columns(3)
        col1.metric("當月目標", f"{month_target:,.2f} 萬")
        col2.metric("總核實 FYC", f"{month_actual:,.2f} 萬")
        col3.metric("當月達成率", f"{month_rate * 100:.2f}%")
        
        # 繪製進度條 (防呆機制：確保進度條最大值不超過 1.0)
        progress_value_month = min(month_actual / month_target, 1.0) if month_target > 0 else 0
        st.progress(progress_value_month)
        
        st.divider() # 分隔線
        
        # ====== 2. 累計 FYC 達成進度 ======
        st.markdown("### 🏆 累計 FYC 達成進度")
        col4, col5, col6 = st.columns(3)
        col4.metric("累計目標", f"{year_target:,.2f} 萬")
        col5.metric("累計核實 FYC", f"{year_actual:,.2f} 萬")
        col6.metric("累計達成率", f"{year_rate * 100:.2f}%")
        
        progress_value_year = min(year_actual / year_target, 1.0) if year_target > 0 else 0
        st.progress(progress_value_year)
        
        # 如果累計達標，觸發慶祝特效
        if year_rate >= 1:
            st.balloons()
            st.info("🎉 太棒了！累計 FYC 目標已順利達標！")
            
    else:
        st.error("找不到目標通訊處的資料，請確認上傳的 CSV 檔案是否正確。")