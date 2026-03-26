import streamlit as st
import pandas as pd
import os

# ==========================================
# 網頁基本設定
# ==========================================
st.set_page_config(page_title="竹耀戰情室", layout="wide")
st.title("🚀 竹耀戰情儀表板")

# 設定要自動讀取的檔案名稱
file_fyc = "data_fyc.xlsx"
file_kpi = "data_kpi.xlsm"

has_fyc, has_team, has_kpi, has_daily = False, False, False, False
hero_daily_list, hero_accum_list = [], []

# ==========================================
# 模組 A：自動讀取 FYC 核實報表 (最終戰果)
# ==========================================
if os.path.exists(file_fyc):
    try:
        # 讀取單位核實進度
        df_unit = pd.read_excel(file_fyc, sheet_name="當期通訊處排名-FYC", skiprows=5, header=None, engine='openpyxl')
        target_row = df_unit[df_unit[2] == '竹耀']
        if not target_row.empty:
            data = target_row.iloc[0]
            month_target, month_actual, month_rate = float(data[5]), float(data[17]), float(data[18])
            year_target, year_actual, year_rate = float(data[6]), float(data[27]), float(data[28])
            has_fyc = True

        # 讀取個人核實排名 (用於貢獻長條圖)
        df_person = pd.read_excel(file_fyc, sheet_name="個人排名_FYC", skiprows=5, header=None, engine='openpyxl')
        team_data = df_person[df_person[3] == 'HC157'].copy()
        if not team_data.empty:
            chart_data = pd.DataFrame({
                '夥伴姓名': team_data[4].astype(str),
                '職稱': team_data[5].astype(str),
                '總核實FYC': pd.to_numeric(team_data[17], errors='coerce').fillna(0)
            }).sort_values(by='總核實FYC', ascending=False)
            has_team = True

    except Exception as e:
        st.error(f"讀取 {file_fyc} 發生錯誤：{e}")

# ==========================================
# 模組 B：自動讀取 KPI 與 受理業績報表 (每日動能)
# ==========================================
if os.path.exists(file_kpi):
    # 1. 抓取關鍵指標
    try:
        df_kpi = pd.read_excel(file_kpi, sheet_name="關鍵指標 (分隊)", engine='openpyxl')
        mask = df_kpi.iloc[:, 1].astype(str).str.contains('HC157')
        kpi_row = df_kpi[mask]
        if not kpi_row.empty:
            kdata = kpi_row.iloc[0]
            fyc_rate, ju_rate, shi_rate, zhuang_rate = float(kdata.iloc[5]), float(kdata.iloc[13]), float(kdata.iloc[21]), float(kdata.iloc[29])
            has_kpi = True
    except Exception as e:
        st.error(f"讀取關鍵指標發生錯誤：{e}")

    # 2. 抓取每日/累計受理排名 (TEAM 分隊)
    try:
        df_daily = pd.read_excel(file_kpi, sheet_name="TEAM (分隊)", engine='openpyxl')
        # 篩選 HC157
        team_mask = df_daily.iloc[:, 1].astype(str).str.contains('HC157')
        df_hc157 = df_daily[team_mask].copy()
        
        if not df_hc157.empty:
            # 💡 核心過濾機制：排除職稱是數字(代表團隊總和)的欄位，只留下個人
            valid_title_mask = pd.to_numeric(df_hc157.iloc[:, 3], errors='coerce').isna() & df_hc157.iloc[:, 3].notna()
            individuals = df_hc157[valid_title_mask].copy()
            
            # 清理名字中的全形空白
            individuals.iloc[:, 2] = individuals.iloc[:, 2].astype(str).str.replace('　', '').str.strip()
            # 確保數字格式
            individuals.iloc[:, 5] = pd.to_numeric(individuals.iloc[:, 5], errors='coerce').fillna(0) # 當日受理
            individuals.iloc[:, 7] = pd.to_numeric(individuals.iloc[:, 7], errors='coerce').fillna(0) # 累計受理
            
            # 抓取前三名
            daily_top3 = individuals.sort_values(by=individuals.columns[5], ascending=False).head(3)
            accum_top3 = individuals.sort_values(by=individuals.columns[7], ascending=False).head(3)
            
            # 建立英雄榜資料函數
            def build_hero_list(df_top):
                medal_colors = ["🥇 金牌", "🥈 銀牌", "🥉 銅牌"]
                result = []
                for i, (_, row) in enumerate(df_top.iterrows()):
                    name = str(row.iloc[2])
                    photo_path = f"{name}.png" 
                    if not os.path.exists(photo_path):
                        photo_path = f"{name}.jpg"
                        if not os.path.exists(photo_path):
                            photo_path = "https://w7.pngwing.com/pngs/129/292/png-transparent-computer-icons-user-profile-male-avatar-avatar-heroes-human-male.png"
                    result.append({
                        'rank': medal_colors[i], 'name': name, 'title': str(row.iloc[3]),
                        'photo': photo_path, 'value': row.iloc[5] if df_top.equals(daily_top3) else row.iloc[7]
                    })
                return result
            
            hero_daily_list = build_hero_list(daily_top3)
            hero_accum_list = build_hero_list(accum_top3)
            has_daily = True

    except Exception as e:
        st.error(f"讀取 TEAM(分隊) 發生錯誤：{e}")

# ==========================================
# 繪製網頁畫面
# ==========================================
if has_fyc or has_team or has_kpi or has_daily:
    st.success("✅ 戰情資料已自動更新至最新版！")
    
    # 🎯 模組 1: 單位活動率
    if has_kpi:
        st.markdown("### 🎯 竹耀關鍵指標")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("FYC 達成率", f"{fyc_rate * 100:.2f}%")
        k2.metric("舉績率", f"{ju_rate * 100:.2f}%")
        k3.metric("實動率", f"{shi_rate * 100:.2f}%")
        k4.metric("壯實人力率", f"{zhuang_rate * 100:.2f}%")
        st.divider()

    # 🥇 模組 2: 每日與累計受理英雄榜
    if has_daily:
        st.markdown("<h2 style='text-align: center; color: #ffcc00;'>🏆 本月受理動能英雄榜</h2>", unsafe_allow_html=True)
        
        tab1, tab2 = st.tabs(["🔥 今日受理 Top 3", "📈 當月累計受理 Top 3"])
        
        def render_heroes(hero_list, label):
            h_cols = st.columns(3)
            for i, col in enumerate(h_cols):
                if i < len(hero_list):
                    hero = hero_list[i]
                    with col:
                        st.markdown(f"""
                        <div style="text-align: center; border: 2px solid #ddd; border-radius: 10px; padding: 15px; background-color: #f9f9f9;">
                            <h3 style="color: #333;">{hero['rank']}</h3>
                            <img src="{hero['photo']}" width="150" style="border-radius: 50%; aspect-ratio: 1/1; object-fit: cover;">
                            <h2 style="margin-top: 10px; color: #1a73e8;">{hero['name']}</h2>
                            <p style="color: #666; margin-top: -10px;">({hero['title']})</p>
                            <hr>
                            <p style="font-size: 1.2em; color: #333;">{label}</p>
                            <h1 style="color: #d93025; font-size: 2.5em; margin-top: -15px;">{hero['value']:,.0f}</h1>
                        </div>
                        """, unsafe_allow_html=True)
                        
        with tab1:
            render_heroes(hero_daily_list, "今日受理 (FYC)")
        with tab2:
            render_heroes(hero_accum_list, "累計受理 (FYC)")
            
        st.divider()

    # 📊 模組 3: FYC 核實達成進度
    if has_fyc:
        st.markdown("### 📊 年度核實進度總覽 (最終戰果)")
        col_m, col_y = st.columns(2)
        with col_m:
            c1, c2, c3 = st.columns(3)
            c1.metric("當月目標", f"{month_target:,.2f} 萬")
            c2.metric("總核實 FYC", f"{month_actual:,.2f} 萬")
            c3.metric("核實達成率", f"{month_rate * 100:.2f}%")
            st.progress(min(month_actual / month_target, 1.0) if month_target > 0 else 0)

        with col_y:
            c4, c5, c6 = st.columns(3)
            c4.metric("累計目標", f"{year_target:,.2f} 萬")
            c5.metric("累計核實 FYC", f"{year_actual:,.2f} 萬")
            c6.metric("累計達成率", f"{year_rate * 100:.2f}%")
            st.progress(min(year_actual / year_target, 1.0) if year_target > 0 else 0)
        
        if year_rate >= 1:
            st.balloons()
        st.divider()

    # 👥 模組 4: 個人核實業績排行榜
    if has_team:
        st.markdown("### 👥 團隊夥伴年度核實貢獻排行榜")
        col_chart, col_table = st.columns([2, 1])
        with col_chart:
            st.bar_chart(chart_data.set_index('夥伴姓名')['總核實FYC'])
        with col_table:
            st.dataframe(chart_data, hide_index=True, use_container_width=True)
else:
    st.info("🔄 系統正在等待最新的報表資料...")