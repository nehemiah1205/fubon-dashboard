import streamlit as st
import pandas as pd
import os
import base64
import altair as alt

# ==========================================
# 網頁基本設定
# ==========================================
st.set_page_config(page_title="竹耀戰情室", layout="wide")

# ==========================================
# 🎨 莫蘭迪藍 淺色主題
# ==========================================
MORANDI_CSS = """
<style>
:root {
    --morandi-bg-top: #F1F5F6;
    --morandi-bg-bottom: #E4EBED;
    --morandi-surface: #FFFFFF;
    --morandi-border: #DCE6E8;
    --morandi-blue: #7C97A3;
    --morandi-blue-deep: #51707D;
    --morandi-blue-soft: #A9BFC8;
    --morandi-gold: #C4A576;
    --morandi-rose: #B98072;
    --morandi-sage: #8FA88C;
    --morandi-text: #3E4A50;
    --morandi-text-soft: #83949A;
}

.stApp {
    background: linear-gradient(180deg, var(--morandi-bg-top) 0%, var(--morandi-bg-bottom) 100%);
}

[data-testid="stHeader"] {
    background: transparent;
}

h1, h2, h3 {
    color: var(--morandi-blue-deep) !important;
    font-weight: 700 !important;
}

.morandi-hero-title {
    text-align: center;
    letter-spacing: 2px;
}

hr {
    border-color: var(--morandi-border) !important;
}

/* 連結按鈕 / 按鈕 */
.stLinkButton a, .stButton button {
    background-color: var(--morandi-blue) !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    box-shadow: 0 3px 8px rgba(81, 112, 125, 0.18) !important;
}
.stLinkButton a:hover, .stButton button:hover {
    background-color: var(--morandi-blue-deep) !important;
}

/* 提示框 (success / warning / error) */
[data-testid="stAlert"] {
    background-color: var(--morandi-surface) !important;
    border: 1px solid var(--morandi-border) !important;
    border-radius: 10px !important;
    color: var(--morandi-text) !important;
}

/* st.metric */
[data-testid="stMetric"] {
    background-color: var(--morandi-surface);
    border: 1px solid var(--morandi-border);
    border-radius: 12px;
    padding: 14px 10px;
    box-shadow: 0 2px 6px rgba(81, 112, 125, 0.06);
}
[data-testid="stMetricLabel"] {
    color: var(--morandi-text-soft) !important;
}
[data-testid="stMetricValue"] {
    color: var(--morandi-blue-deep) !important;
}

/* 分頁 tabs */
.stTabs [data-baseweb="tab-list"] {
    gap: 6px;
}
.stTabs [data-baseweb="tab"] {
    background-color: var(--morandi-surface);
    border: 1px solid var(--morandi-border);
    border-radius: 8px 8px 0 0;
    color: var(--morandi-text-soft);
    font-weight: 600;
}
.stTabs [aria-selected="true"] {
    background-color: var(--morandi-blue) !important;
    color: #ffffff !important;
    border-color: var(--morandi-blue) !important;
}

/* 進度條 */
[data-testid="stProgress"] > div > div > div {
    background-color: var(--morandi-blue) !important;
}
[data-testid="stProgress"] > div > div {
    background-color: var(--morandi-border) !important;
}

/* 表格 */
[data-testid="stDataFrame"] {
    border: 1px solid var(--morandi-border);
    border-radius: 12px;
    overflow: hidden;
}
</style>
"""
st.markdown(MORANDI_CSS, unsafe_allow_html=True)

st.title("🚀 竹耀戰情儀表板")

# 設定要自動讀取的檔案名稱
file_fyc = "data_fyc.xlsx"
file_kpi = "data_kpi.xlsm"

has_fyc, has_team, has_kpi, has_daily = False, False, False, False
has_monthly_rank = False
hero_daily_list, hero_accum_list = [], []
unit_daily_fyc = 0.0
unit_accum_fyc = 0.0
monthly_rank_data = pd.DataFrame()

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
        target_row = df_unit[df_unit.apply(lambda r: r.astype(str).str.contains('HC157').any(), axis=1)]
        if not target_row.empty:
            data = target_row.iloc[0]
            month_target, month_actual, month_rate = float(data[5]), float(data[17]), float(data[18])
            year_target, year_actual, year_rate = float(data[6]), float(data[27]), float(data[28])
            has_fyc = True

        df_person = pd.read_excel(file_fyc, sheet_name="個人排名_FYC", skiprows=7, header=None, engine='openpyxl')
        team_data = df_person[df_person[1].astype(str).str.contains('HC157', na=False)].copy()
        
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
        
        # 📡 雷達 A：專抓「上方 FYC 指標」 (鎖定 HC157 單位總計列)
        mask_hc157_exact = df_kpi.iloc[:, 1].astype(str).str.strip() == 'HC157'
        kpi_row_fyc = df_kpi[mask_hc157_exact]
        if kpi_row_fyc.empty: # 備用方案
            kpi_row_fyc = df_kpi[df_kpi.iloc[:, 1].astype(str).str.contains('HC157', na=False)]
            
        if not kpi_row_fyc.empty:
            kdata_fyc = kpi_row_fyc.iloc[0]
            try:
                fyc_rank = int(float(kdata_fyc.iloc[0]))
            except:
                fyc_rank = "-"
            unit_daily_fyc = float(kdata_fyc.iloc[3]) if pd.notnull(kdata_fyc.iloc[3]) else 0.0
            unit_accum_fyc = float(kdata_fyc.iloc[4]) if pd.notnull(kdata_fyc.iloc[4]) else 0.0
            fyc_rate = clean_pct(kdata_fyc.iloc[5])
            has_kpi = True

        # 📡 雷達 B：專抓「下方 人力機率指標」 (鎖定 王新智 主管列)
        kpi_row_manpower = df_kpi[df_kpi.apply(lambda r: r.astype(str).str.contains('王新智').any(), axis=1)]
        if not kpi_row_manpower.empty:
            kdata_manpower = kpi_row_manpower.iloc[0]
            ju_rate = clean_pct(kdata_manpower.iloc[13])
            shi_rate = clean_pct(kdata_manpower.iloc[21])
            zhuang_rate = clean_pct(kdata_manpower.iloc[29])
        else:
            # 如果萬一找不到主管名字，備用抓取原本的 HC157 列
            ju_rate = clean_pct(kdata_fyc.iloc[13]) if not kpi_row_fyc.empty else 0.0
            shi_rate = clean_pct(kdata_fyc.iloc[21]) if not kpi_row_fyc.empty else 0.0
            zhuang_rate = clean_pct(kdata_fyc.iloc[29]) if not kpi_row_fyc.empty else 0.0
            
    except Exception as e:
        st.error(f"❌ 讀取 KPI 指標時發生錯誤：{e}") 

    try:
        df_daily = pd.read_excel(file_kpi, sheet_name="TEAM (分隊)", engine='openpyxl')
        team_mask = df_daily.iloc[:, 1].astype(str).str.contains('HC157', na=False)
        df_hc157 = df_daily[team_mask].copy()
        
        if not df_hc157.empty:
            valid_title_mask = pd.to_numeric(df_hc157.iloc[:, 3], errors='coerce').isna() & df_hc157.iloc[:, 3].notna()
            individuals = df_hc157[valid_title_mask].copy()
            
            individuals.iloc[:, 2] = individuals.iloc[:, 2].astype(str).str.replace(' ', '').str.strip()
            individuals.iloc[:, 5] = pd.to_numeric(individuals.iloc[:, 5], errors='coerce').fillna(0)  # 當日FYC
            individuals.iloc[:, 6] = pd.to_numeric(individuals.iloc[:, 6], errors='coerce').fillna(0)  # 累計受理件數
            individuals.iloc[:, 7] = pd.to_numeric(individuals.iloc[:, 7], errors='coerce').fillna(0)  # 累計FYC
            
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

            # 📊 本月受理排行榜（全員，含掛零，資料來源同 TEAM 分頁）
            monthly_rank_data = pd.DataFrame({
                '夥伴姓名': individuals.iloc[:, 2].astype(str),
                '職稱': individuals.iloc[:, 3].astype(str),
                '累計受理件數': individuals.iloc[:, 6],
                '累計受理FYC': individuals.iloc[:, 7],
            }).sort_values(by='累計受理FYC', ascending=False)
            has_monthly_rank = True
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
            <div style="text-align: center; border: 1px solid #DCE6E8; border-radius: 14px; padding: 20px; background-color: #FFFFFF; box-shadow: 0 4px 14px rgba(81,112,125,0.10);">
                <p style="font-size: 1.1em; color: #83949A; margin-bottom: 5px; font-weight: 600; letter-spacing: 0.5px;">{title}</p>
                <h1 style="color: {color}; font-size: 2.6em; margin: 0; font-weight: 800; letter-spacing: 1px;">{value}</h1>
            </div>
            """

        r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
        with r1_col1:
            st.markdown(big_metric_card("🏆 通訊處排名", f"第 {fyc_rank} 名", "#C4A576"), unsafe_allow_html=True)
        with r1_col2:
            st.markdown(big_metric_card("🔥 單日受理 FYC", f"{unit_daily_fyc:,.0f}", "#7C97A3"), unsafe_allow_html=True)
        with r1_col3:
            st.markdown(big_metric_card("📈 累計受理 FYC", f"{unit_accum_fyc:,.0f}", "#B98072"), unsafe_allow_html=True)
        with r1_col4:
            st.markdown(big_metric_card("🎯 FYC 達成率", f"{fyc_rate * 100:.1f}%", "#8FA88C"), unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        r2_col1, r2_col2, r2_col3 = st.columns(3)
        r2_col1.metric("舉績率", f"{ju_rate * 100:.1f}%")
        r2_col2.metric("實動率", f"{shi_rate * 100:.1f}%")
        r2_col3.metric("壯實人力率", f"{zhuang_rate * 100:.1f}%")
        st.divider()

    if has_daily:
        st.markdown("<h2 style='text-align: center; color: #C4A576;'>🏆 本日受理英雄榜</h2>", unsafe_allow_html=True)
        tab1, tab2 = st.tabs(["🔥 今日受理 Top 3", "📈 當月累計受理 Top 3"])
        
        def render_heroes(hero_list, label):
            h_cols = st.columns(3)
            for i, col in enumerate(h_cols):
                if i < len(hero_list):
                    hero = hero_list[i]
                    with col:
                        st.markdown(f"""
                        <div style="text-align: center; border: 1px solid #DCE6E8; border-radius: 14px; padding: 15px; background-color: #FFFFFF; box-shadow: 0 4px 14px rgba(81,112,125,0.10);">
                            <h3 style="color: #C4A576;">{hero['rank']}</h3>
                            <img src="{hero['photo_src']}" width="150" height="150" style="border-radius: 50%; object-fit: cover; box-shadow: 0 4px 8px rgba(81,112,125,0.15); border: 3px solid #E4EBED;">
                            <h2 style="margin-top: 15px; color: #51707D;">{hero['name']}</h2>
                            <p style="color: #83949A; margin-top: -10px;">({hero['title']})</p>
                            <hr style="border-color: #DCE6E8;">
                            <p style="font-size: 1.2em; color: #3E4A50;">{label}</p>
                            <h1 style="color: #B98072; font-size: 2.5em; margin-top: -15px;">{hero['value']:,.0f}</h1>
                        </div>
                        """, unsafe_allow_html=True)
                        
        with tab1:
            if not hero_daily_list:
                st.markdown("""
                <div style="text-align: center; padding: 50px; background-color: #FFFFFF; border-radius: 14px; border: 2px dashed #DCE6E8;">
                    <h2 style="color: #83949A;">⏳ 今日尚未有夥伴報件</h2>
                    <p style="color: #A9BFC8; font-size: 1.2em;">全體準備中，等待首件捷報！💪</p>
                </div>
                """, unsafe_allow_html=True)
            else:
                render_heroes(hero_daily_list, "今日受理 (FYC)")
                
        with tab2:
            if not hero_accum_list:
                st.markdown("""
                <div style="text-align: center; padding: 50px; background-color: #FFFFFF; border-radius: 14px; border: 2px dashed #DCE6E8;">
                    <h2 style="color: #83949A;">⏳ 本月尚未有夥伴報件</h2>
                    <p style="color: #A9BFC8; font-size: 1.2em;">大家繼續努力，創造佳績！💪</p>
                </div>
                """, unsafe_allow_html=True)
            else:
                render_heroes(hero_accum_list, "累計受理 (FYC)")
        st.divider()

    if has_monthly_rank:
        st.markdown("### 📊 本月受理排行榜")
        col_chart_m, col_table_m = st.columns([2, 1])
        with col_chart_m:
            chart_monthly = alt.Chart(monthly_rank_data).mark_bar(color='#8FA88C').encode(
                x=alt.X('夥伴姓名', sort='-y', axis=alt.Axis(labelAngle=0)),
                y=alt.Y('累計受理FYC', title='累計受理FYC'),
                tooltip=['夥伴姓名', '職稱', '累計受理件數', '累計受理FYC']
            ).properties(height=400)
            st.altair_chart(chart_monthly, use_container_width=True)

        with col_table_m:
            st.dataframe(monthly_rank_data, hide_index=True, use_container_width=True)
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
            chart = alt.Chart(chart_data).mark_bar(color='#7C97A3').encode(
                x=alt.X('夥伴姓名', sort='-y', axis=alt.Axis(labelAngle=0)), 
                y=alt.Y('總核實FYC', title='總核實FYC'),
                tooltip=['夥伴姓名', '職稱', '總核實FYC']
            ).properties(height=400)
            st.altair_chart(chart, use_container_width=True)
            
        with col_table:
            st.dataframe(chart_data, hide_index=True, use_container_width=True)