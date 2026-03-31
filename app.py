import streamlit as st
import pandas as pd
import os
import base64
import altair as alt
import msoffcrypto
import io

# ==========================================
# 網頁基本設定
# ==========================================
st.set_page_config(page_title="竹耀戰情室", layout="wide")
st.title("🚀 竹耀戰情儀表板")

# 🔐 固定密碼（直接改這裡）
FYC_PASSWORD = "請填入FYC密碼"
KPI_PASSWORD = "請填入KPI密碼"

has_fyc, has_team, has_kpi, has_daily = False, False, False, False
hero_daily_list, hero_accum_list = [], []

# ==========================================
# 🛠️ 工具函式
# ==========================================
def get_image_base64(image_path):
    """圖片轉 Base64"""
    try:
        with open(image_path, "rb") as img_file:
            encoded_string = base64.b64encode(img_file.read()).decode()
            ext = "jpeg" if image_path.lower().endswith(".jpg") else "png"
            return f"data:image/{ext};base64,{encoded_string}"
    except Exception:
        return None

def decrypt_excel(file_bytes, password):
    """將加密的 Excel bytes 解密後回傳 BytesIO"""
    try:
        buf = io.BytesIO(file_bytes)
        office_file = msoffcrypto.OfficeFile(buf)
        office_file.load_key(password=password)
        decrypted = io.BytesIO()
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        return decrypted
    except Exception as e:
        st.error(f"❌ 解密失敗：{e}，請確認密碼是否正確")
        return None

def open_excel(file_bytes, password):
    """
    嘗試直接開啟，若失敗則先解密。
    每次呼叫都回傳一個全新的 BytesIO，可安全地 seek(0) 重複使用。
    """
    # 先試未加密
    try:
        buf = io.BytesIO(file_bytes)
        pd.read_excel(buf, sheet_name=0, nrows=1, engine='openpyxl')
        return io.BytesIO(file_bytes)  # 回傳全新 buffer
    except Exception:
        pass
    # 有加密，解密
    return decrypt_excel(file_bytes, password)

# ==========================================
# 📂 側邊欄：上傳檔案
# ==========================================
with st.sidebar:
    st.header("📂 上傳報表檔案")
    st.caption("收到 Email 附件後，直接在這裡上傳即可更新儀表板")

    uploaded_fyc = st.file_uploader(
        "上傳 FYC 核實報表 (.xlsx)",
        type=["xlsx", "xls"],
        key="fyc"
    )
    uploaded_kpi = st.file_uploader(
        "上傳 KPI 受理報表 (.xlsm)",
        type=["xlsm", "xlsx", "xls"],
        key="kpi"
    )

    st.markdown("---")
    st.success(f"✅ FYC：{uploaded_fyc.name}") if uploaded_fyc else st.warning("⏳ 尚未上傳 FYC 報表")
    st.success(f"✅ KPI：{uploaded_kpi.name}") if uploaded_kpi else st.warning("⏳ 尚未上傳 KPI 報表")

# ==========================================
# 模組 A：讀取 FYC 核實報表
# ==========================================
if uploaded_fyc:
    try:
        fyc_bytes = uploaded_fyc.read()

        buf1 = open_excel(fyc_bytes, FYC_PASSWORD)
        if buf1:
            df_unit = pd.read_excel(buf1, sheet_name="當期通訊處排名-FYC", skiprows=5, header=None, engine='openpyxl')
            target_row = df_unit[df_unit[2] == '竹耀']
            if not target_row.empty:
                data = target_row.iloc[0]
                month_target, month_actual, month_rate = float(data[5]), float(data[17]), float(data[18])
                year_target, year_actual, year_rate = float(data[6]), float(data[27]), float(data[28])
                has_fyc = True

        buf2 = open_excel(fyc_bytes, FYC_PASSWORD)
        if buf2:
            df_person = pd.read_excel(buf2, sheet_name="個人排名_FYC", skiprows=5, header=None, engine='openpyxl')
            team_data = df_person[df_person[3] == 'HC157'].copy()
            if not team_data.empty:
                chart_data = pd.DataFrame({
                    '夥伴姓名': team_data[4].astype(str),
                    '職稱': team_data[5].astype(str),
                    '總核實FYC': pd.to_numeric(team_data[17], errors='coerce').fillna(0)
                }).sort_values(by='總核實FYC', ascending=False)
                has_team = True

    except Exception as e:
        st.error(f"讀取 FYC 報表發生錯誤：{e}")

# ==========================================
# 模組 B：讀取 KPI 與受理業績報表
# ==========================================
if uploaded_kpi:
    try:
        kpi_bytes = uploaded_kpi.read()

        buf3 = open_excel(kpi_bytes, KPI_PASSWORD)
        if buf3:
            df_kpi = pd.read_excel(buf3, sheet_name="關鍵指標 (分隊)", engine='openpyxl')
            mask = df_kpi.iloc[:, 1].astype(str).str.contains('HC157')
            kpi_row = df_kpi[mask]
            if not kpi_row.empty:
                kdata = kpi_row.iloc[0]
                try:
                    fyc_rank = int(float(kdata.iloc[0]))
                except Exception:
                    fyc_rank = "-"
                unit_daily_fyc  = float(kdata.iloc[3])  if pd.notnull(kdata.iloc[3])  else 0.0
                unit_accum_fyc  = float(kdata.iloc[4])  if pd.notnull(kdata.iloc[4])  else 0.0
                fyc_rate        = float(kdata.iloc[5])  if pd.notnull(kdata.iloc[5])  else 0.0
                ju_rate         = float(kdata.iloc[13]) if pd.notnull(kdata.iloc[13]) else 0.0
                shi_rate        = float(kdata.iloc[21]) if pd.notnull(kdata.iloc[21]) else 0.0
                zhuang_rate     = float(kdata.iloc[29]) if pd.notnull(kdata.iloc[29]) else 0.0
                has_kpi = True
    except Exception:
        pass

    try:
        buf4 = open_excel(kpi_bytes, KPI_PASSWORD)
        if buf4:
            df_daily = pd.read_excel(buf4, sheet_name="TEAM (分隊)", engine='openpyxl')
            team_mask = df_daily.iloc[:, 1].astype(str).str.contains('HC157')
            df_hc157 = df_daily[team_mask].copy()
            if not df_hc157.empty:
                valid_title_mask = (
                    pd.to_numeric(df_hc157.iloc[:, 3], errors='coerce').isna()
                    & df_hc157.iloc[:, 3].notna()
                )
                individuals = df_hc157[valid_title_mask].copy()
                individuals.iloc[:, 2] = individuals.iloc[:, 2].astype(str).str.replace('　', '').str.strip()
                individuals.iloc[:, 5] = pd.to_numeric(individuals.iloc[:, 5], errors='coerce').fillna(0)
                individuals.iloc[:, 7] = pd.to_numeric(individuals.iloc[:, 7], errors='coerce').fillna(0)
                daily_top3 = individuals.sort_values(by=individuals.columns[5], ascending=False).head(3)
                accum_top3 = individuals.sort_values(by=individuals.columns[7], ascending=False).head(3)

                def build_hero_list(df_top, is_daily):
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
                            'photo_src': img_src,
                            'value': row.iloc[5] if is_daily else row.iloc[7]
                        })
                    return result

                hero_daily_list = build_hero_list(daily_top3, is_daily=True)
                hero_accum_list = build_hero_list(accum_top3, is_daily=False)
                has_daily = True
    except Exception:
        pass

# ==========================================
# 繪製網頁畫面
# ==========================================
if has_fyc or has_team or has_kpi or has_daily:
    st.success("✅ 戰情資料已載入最新版！")

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
            st.markdown(big_metric_card("🎯 FYC 達成率", f"{fyc_rate * 100:.2f}%", "#34a853"), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        r2_col1, r2_col2, r2_col3 = st.columns(3)
        r2_col1.metric("舉績率", f"{ju_rate * 100:.2f}%")
        r2_col2.metric("實動率", f"{shi_rate * 100:.2f}%")
        r2_col3.metric("壯實人力率", f"{zhuang_rate * 100:.2f}%")
        st.divider()

    if has_daily:
        st.markdown("<h2 style='text-align: center; color: #ffcc00;'>🏆 本月受理英雄榜</h2>", unsafe_allow_html=True)
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
            render_heroes(hero_daily_list, "今日受理 (FYC)")
        with tab2:
            render_heroes(hero_accum_list, "累計受理 (FYC)")
        st.divider()

    if has_fyc:
        st.markdown("### 📊 上月核實進度總覽")
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

else:
    st.info("👈 請先在左側上傳 FYC 與 KPI 報表，儀表板將自動顯示最新資料")
