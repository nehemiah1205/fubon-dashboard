import streamlit as st
import pandas as pd
import os
import base64
import altair as alt

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
unit_daily_fyc = 0.0
unit_accum_fyc = 0.0

# 🛠️ 圖片轉碼器
def get_image_base64(image_path):
    try:
        with open(image_path, "rb") as img_file:
            encoded_string = base64.b64encode(img_file.read()).decode()
            ext = "jpeg" if image_path.lower().endswith(".jpg") else "png"
            return f"data:image/{ext};base64,{encoded_string}"
    except Exception:
        return None

# 🛠️ 數字清洗濾網
def clean_pct(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, str):
        v = val.replace('%', '').replace(',', '').strip()
        try:
            return float(v) / 100.0
        except:
            return 0.0
    try:
        return float(val)
    except:
        return 0.0

# ==========================================
# 模組 A：自動讀取 FYC 核實報表 (最終戰果)
# ==========================================
if os.path.exists(file_fyc):
    try:
        df_unit = pd.read_excel(file_fyc, sheet_name="當期通訊處排名-FYC", skiprows=5, header=None, engine='openpyxl')
        # FYC 報表可能包含單位名稱或主管名稱，這裡我們用 HC157 作為主要防線
        target_row = df_unit[df_unit.apply(lambda r: r.astype(str).str.contains('HC157').any(), axis=1)]
        if not target_row.empty:
            data = target_row.iloc[0]
            month_target, month_actual, month_rate = float(data[5]), float(data[17]), float(data[18])
            year_target, year_actual, year_rate = float(data[6]), float(data[27]), float(data[28])
            has_fyc = True

        df_person = pd.read_excel(file_fyc, sheet_name="個人排名_FYC", skiprows=7, header=None, engine='openpyxl')
        team_data = df_person[df_person[3].astype(str).str.contains('HC157', na=False)].copy()
        
        if not team_data.empty:
            chart_data = pd.DataFrame({
                '夥伴姓名': team_data.iloc[:, 2].astype(str),     
                '職稱': team_data.iloc[:, 3].astype(str),        
                '總核實FYC': pd.to_numeric(team_data.iloc[:, 15], errors='coerce').fillna(0) 
            }).sort_values(by='總核實FYC', ascending=False)
            
            chart_data = chart_data[chart_data['總核實FYC'] > 0]
            has_team = True
            
    except Exception as e:
        st.error(f"❌ 讀取 {file_fyc} 發生錯誤：{e}")
else:
    st.warning(f"⚠️ 雲端找不到檔案：{file_fyc} (請確認是否已上傳)")

# ==========================================
# 模組 B：自動讀取 KPI 與 受理業績報表 (每日動能)
# ==========================================
if os.path.exists(file_kpi):
    try:
        df_kpi = pd.read_excel(file_kpi, sheet_name="關鍵指標 (分隊)", engine='openpyxl')
        
        # 🎯 終極精準雷達：雙重身分核對 (必須是 HC157，且整列包含主管「王新智」，完美避開子分隊)
        mask_hc157 = df_kpi.iloc[:, 1].astype(str).str.contains('HC157', na=False)
        mask_manager = df_kpi.apply(lambda r: r.astype(str).str.contains('王新智').any(), axis=1)
        
        kpi_row = df_kpi[mask_hc157 & mask_manager]
        
        if kpi_row.empty:
            # 若發生例外，退一步找「精確等於」HC157 的那一列
            kpi_row = df_kpi[df_kpi.iloc[:, 1].astype(str).str.strip() == 'HC157']
            
        if not kpi_row.empty:
            kdata = kpi_row.iloc[0]
            try:
                fyc_rank = int(float(kdata.iloc[0]))
            except:
                fyc_rank = "-"
                
            unit_daily_fyc = float(kdata.iloc[3]) if pd.notnull(kdata.iloc[3]) else 0.0
            unit_accum_fyc = float(kdata.iloc[4]) if pd.notnull(kdata.iloc[4]) else 0.0
            
            fyc_rate = clean_pct(kdata.iloc[5])
            ju_rate = clean_pct(kdata.iloc[13])
            shi_rate = clean_pct(kdata.iloc[21])
            zhuang_rate = clean_pct(kdata.iloc[29])
            
            has_kpi = True
    except Exception as e:
        st.error(f"❌ 讀取 KPI 指標時發生錯誤：{e}") 

    try:
        df_daily = pd.read_excel(file_kpi, sheet_name="TEAM (分隊)", engine='openpyxl')
        
        # 這裡的 contains 是對的，因為我們要抓底下所有子單位的夥伴
        team_mask = df_daily.iloc[:, 1].astype(str).str.contains('HC157', na=False)
        df_hc157 = df_daily[team_mask].copy()
        
        if not df_hc157.empty:
            valid_title_mask = pd.to_numeric(df_hc157.iloc[:, 3], errors='coerce').isna() & df_hc157.iloc[:, 3].notna()
            individuals = df_hc157[valid_title_mask].copy()
            
            individuals.iloc[:, 2] = individuals.iloc[:, 2].astype(str).str.replace(' ', '').str.strip()
            individuals.iloc[:, 5] = pd.to_numeric(individuals.iloc[:, 5], errors='coerce').fillna(0)
            individuals.iloc[:, 7] = pd.to_numeric(individuals.iloc[:, 7], errors='coerce').fillna(0)
            
            daily_active = individuals[individuals.iloc[:, 5] > 0]
            daily_top3 = daily_active.sort_values(by=individuals.columns[5], ascending=False).head(3)
            
            accum_active = individuals[individuals.iloc[:, 7] > 0]
            accum_top3 = accum_active.sort_values(by=individuals.columns[7], ascending=False).head(3)
            
            def build_hero_list(df_top):
                medal_colors = ["🥇 金牌", "🥈 銀牌", "🥉 銅牌"]
                result = []
                for i, (_, row) in enumerate(df_top.iterrows()):
                    name = str(row.iloc[2])
                    img_src = "https://w7.pngwing.com/pngs/129/292/png-transparent-computer-icons-user-profile-male-avatar-avatar-heroes-human-male.png"
                    if os.path.exists(f"{name}.png"):
                        img_src = get_image_base64(f"{name}.png")
                    elif os.path.exists(f"{name}.jpg"):
                        img_src = get_image_base64(f"{name}.jpg")
                        
                    result.append({
                        'rank': medal_colors[i], 'name': name, 'title': str(row.iloc[3]),
                        'photo_src': img_src, 'value': row.iloc[5] if df_top.equals(daily_top3) else row.iloc[7]
                    })
                return result
            
            hero_daily_list = build_hero_list(daily_top3)
            hero_accum_list = build_hero_list(accum_top3)
            has_daily = True
    except Exception as e:
        st.error(f"❌ 讀取 TEAM 英雄榜時發生錯誤：{e}")
else:
    st.warning(f"⚠️ 雲端找不到檔案：{file_kpi} (請確認是否已上傳)")

# ==========================================
# 繪製網頁畫面
# ==========================================
if has_fyc or has_team or has_kpi or has_daily:
    st.success("✅ 戰情資料已自動更新至最新版！")
    
    st.markdown("### 🎯 本週業績動能預報")
    st.link_button("📝 點此回報本週預估新增 FYC", "https://forms.gle/7vL5Xw9RQJepSwVP8", use_container_width=True)
    st.divider()
    
    if has_kpi:
        st.markdown("### 🎯 單位戰力與關鍵指標")
        
        def big_metric_card(title, value, color):
            return f"""
            <div style="text-align: center; border: 2px solid #eee; border-radius: 10px; padding: 20px; background-color: #fff; box-shadow: 0 4px 10px rgba(0,0,0,0.08);">
                <p style="font-size: 1.2em; color: #555; margin-bottom: 5px; font-weight: bold;">{title}</p>
                <h1 style="color: {color}; font-size: 2.8em; margin: 0; font-weight: 900; letter-spacing: 1px;">{value}</h1>
            </div>
            """

        r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
        with r1_col1:
            st.markdown(big_metric_card("🏆 通訊處排名", f"第 {fyc_rank} 名", "#ffaa00"), unsafe_allow_html=True)
        with r1_col2:
            st.markdown(big_metric_card("🔥 單日受理 FYC", f"{unit_daily_fyc:,.0f}", "#1a73e8"), unsafe_allow_html=True)
        with r1_col3:
            st.markdown(big_metric_card("📈 累計受理 FYC", f"{unit_accum_fyc:,.0f}", "#d93025"), unsafe_allow_html=True)
        with r1_col4:
            st.markdown(big_metric_card("🎯 FYC 達成率", f"{fyc_rate * 100:.1f}%", "#34a853"), unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        r2_col1, r2_col2, r2_col3 = st.columns(3)
        r2_col1.metric("舉績率", f"{ju_rate * 100:.1f}%")
        r2_col2.metric("實動率", f"{shi_rate * 100:.1f}%")
        r2_col3.metric("壯實人力率", f"{zhuang_rate * 100:.1f}%")
        st.divider()

    if has_daily:
        st.markdown("<h2 style='text-align: center; color: #ffcc00;'>🏆 本日受理英雄榜</h2>", unsafe_allow_html=True)
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
                            <img src="{hero['photo_src']}" width="150" height="150" style="border-radius: 50%; object-fit: cover; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
                            <h2 style="margin-top: 15px; color: #1a73e8;">{hero['name']}</h2>
                            <p style="color: #666; margin-top: -10px;">({hero['title']})</p>
                            <hr>
                            <p style="font-size: 1.2em; color: #333;">{label}</p>
                            <h1 style="color: #d93025; font-size: 2.5em; margin-top: -15px;">{hero['value']:,.0f}</h1>
                        </div>
                        """, unsafe_allow_html=True)
                        
        with tab1:
            if not hero_daily_list:
                st.markdown("""
                <div style="text-align: center; padding: 50px; background-color: #f9f9f9; border-radius: 10px; border: 2px dashed #ccc;">
                    <h2 style="color: #666;">⏳ 今日尚未有夥伴報件</h2>
                    <p style="color: #999; font-size: 1.2em;">全體準備中，等待首件捷報！💪</p>
                </div>
                """, unsafe_allow_html=True)
            else:
                render_heroes(hero_daily_list, "今日受理 (FYC)")
                
        with tab2:
            if not hero_accum_list:
                st.markdown("""
                <div style="text-align: center; padding: 50px; background-color: #f9f9f9; border-radius: 10px; border: 2px dashed #ccc;">
                    <h2 style="color: #666;">⏳ 本月尚未有夥伴報件</h2>
                    <p style="color: #999; font-size: 1.2em;">大家繼續努力，創造佳績！💪</p>
                </div>
                """, unsafe_allow_html=True)
            else:
                render_heroes(hero_accum_list, "累計受理 (FYC)")
        st.divider()

    if has_fyc:
        st.markdown("### 📊 上月核實進度總覽 (最終戰果)")
        col_m, col_y = st.columns(2)
        with col_m:
            c1, c2, c3 = st.columns(3)
            c1.metric("當月目標", f"{month_target:,.2f} 萬")
            c2.metric("總核實 FYC", f"{month_actual:,.2f} 萬")
            c3.metric("核實達成率", f"{month_rate * 100:.1f}%")
            st.progress(min(month_actual / month_target, 1.0) if month_target > 0 else 0)

        with col_y:
            c4, c5, c6 = st.columns(3)
            c4.metric("累計目標", f"{year_target:,.2f} 萬")
            c5.metric("累計核實 FYC", f"{year_actual:,.2f} 萬")
            c6.metric("累計達成率", f"{year_rate * 100:.1f}%")
            st.progress(min(year_actual / year_target, 1.0) if year_target > 0 else 0)
        st.divider()

    if has_team:
        st.markdown("### 👥 上月核實貢獻排行榜")
        col_chart, col_table = st.columns([2, 1])
        with col_chart:
            chart = alt.Chart(chart_data).mark_bar(color='#1a73e8').encode(
                x=alt.X('夥伴姓名', sort='-y', axis=alt.Axis(labelAngle=0)), 
                y=alt.Y('總核實FYC', title='總核實FYC'),
                tooltip=['夥伴姓名', '職稱', '總核實FYC']
            ).properties(height=400)
            st.altair_chart(chart, use_container_width=True)
            
        with col_table:
            st.dataframe(chart_data, hide_index=True, use_container_width=True)