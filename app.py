import streamlit as st
import pandas as pd

# 網頁基本設定
st.set_page_config(page_title="竹耀戰情室", layout="wide")
st.title("🚀 竹耀通訊處戰情儀表板 3.0")

# 1. 升級為「多檔案上傳」 (注意這裡加上了 accept_multiple_files=True)
uploaded_files = st.file_uploader("請一次上傳所有相關的報表 (可框選多個檔案拖曳)", type=["xlsx", "xls", "xlsm"], accept_multiple_files=True)

if uploaded_files:
    # 建立變數來確認我們抓到了哪些資料
    has_fyc = False
    has_team = False
    has_kpi = False

    for file in uploaded_files:
        try:
            # 讀取 Excel 檔案內所有的工作表名稱
            xl = pd.ExcelFile(file, engine='openpyxl')
            sheet_names = xl.sheet_names
            
            # ==========================================
            # 模組 A：核實進度與個人排名 (從第一份報表抓取)
            # ==========================================
            if "當期通訊處排名-FYC" in sheet_names:
                df_unit = pd.read_excel(file, sheet_name="當期通訊處排名-FYC", skiprows=5, header=None, engine='openpyxl')
                target_row = df_unit[df_unit[2] == '竹耀']
                if not target_row.empty:
                    data = target_row.iloc[0]
                    month_target, month_actual, month_rate = float(data[5]), float(data[17]), float(data[18])
                    year_target, year_actual, year_rate = float(data[6]), float(data[27]), float(data[28])
                    has_fyc = True

            if "個人排名_FYC" in sheet_names:
                df_person = pd.read_excel(file, sheet_name="個人排名_FYC", skiprows=5, header=None, engine='openpyxl')
                team_data = df_person[df_person[3] == 'HC157']
                if not team_data.empty:
                    chart_data = pd.DataFrame({
                        '夥伴姓名': team_data[4].astype(str),
                        '職稱': team_data[5].astype(str),
                        '總核實FYC': pd.to_numeric(team_data[17], errors='coerce').fillna(0)
                    }).sort_values(by='總核實FYC', ascending=False)
                    has_team = True

            # ==========================================
            # 模組 B：各項活動率指標 (從新的業績報表抓取)
            # ==========================================
            if "關鍵指標 (分隊)" in sheet_names:
                # 讀取整張表，不需要跳過表頭，我們直接搜尋代號
                df_kpi = pd.read_excel(file, sheet_name="關鍵指標 (分隊)", engine='openpyxl')
                # 找出包含 'HC157' 的那一行
                mask = df_kpi.apply(lambda row: row.astype(str).str.contains('HC157').any(), axis=1)
                kpi_row = df_kpi[mask]
                
                if not kpi_row.empty:
                    kdata = kpi_row.iloc[0] # 抓取那行數字
                    # 依據 AI 解析的欄位位置萃取數字
                    fyc_rate = float(kdata.iloc[5]) if pd.notnull(kdata.iloc[5]) else 0.0
                    ju_rate = float(kdata.iloc[13]) if pd.notnull(kdata.iloc[13]) else 0.0
                    shi_rate = float(kdata.iloc[21]) if pd.notnull(kdata.iloc[21]) else 0.0
                    zhuang_rate = float(kdata.iloc[29]) if pd.notnull(kdata.iloc[29]) else 0.0
                    has_kpi = True
                    
        except Exception as e:
            st.error(f"❌ 讀取 {file.name} 時發生錯誤：{e}")

    # ==========================================
    # 開始繪製網頁畫面
    # ==========================================
    if has_fyc or has_team or has_kpi:
        st.success("✅ 資料載入成功！")
        
        # 顯示 3.0 新增的活動率 KPI 區塊
        if has_kpi:
            st.markdown("### 🎯 單位活動率與關鍵指標")
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("FYC 達成率", f"{fyc_rate * 100:.2f}%")
            k2.metric("舉績率", f"{ju_rate * 100:.2f}%")
            k3.metric("實動率", f"{shi_rate * 100:.2f}%")
            k4.metric("壯實人力率", f"{zhuang_rate * 100:.2f}%")
            st.divider()

        # 顯示原本的 FYC 達成進度區塊
        if has_fyc:
            col_m, col_y = st.columns(2)
            with col_m:
                st.markdown("### 📊 當月 FYC 達成進度")
                c1, c2, c3 = st.columns(3)
                c1.metric("當月目標", f"{month_target:,.2f} 萬")
                c2.metric("總核實 FYC", f"{month_actual:,.2f} 萬")
                c3.metric("核實達成率", f"{month_rate * 100:.2f}%")
                st.progress(min(month_actual / month_target, 1.0) if month_target > 0 else 0)

            with col_y:
                st.markdown("### 🏆 累計 FYC 達成進度")
                c4, c5, c6 = st.columns(3)
                c4.metric("累計目標", f"{year_target:,.2f} 萬")
                c5.metric("累計核實 FYC", f"{year_actual:,.2f} 萬")
                c6.metric("累計達成率", f"{year_rate * 100:.2f}%")
                st.progress(min(year_actual / year_target, 1.0) if year_target > 0 else 0)
            
            if year_rate >= 1:
                st.balloons()
            st.divider()

        # 顯示個人排行榜
        if has_team:
            st.markdown("### 👥 團隊夥伴 FYC 貢獻排行榜")
            col_chart, col_table = st.columns([2, 1])
            with col_chart:
                st.bar_chart(chart_data.set_index('夥伴姓名')['總核實FYC'])
            with col_table:
                st.dataframe(chart_data, hide_index=True, use_container_width=True)

    else:
        st.warning("⚠️ 上傳的檔案中，找不到任何與竹耀通訊處相關的報表資料。")