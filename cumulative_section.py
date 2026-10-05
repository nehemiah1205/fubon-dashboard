"""
竹耀戰情室：本月累積業績走勢區塊

資料來源：auto_rename.py 每天累積的三個 CSV
  history_units.csv     區部各單位每日累計（含目標、舉績／實動／壯實）
  history_people.csv    本單位每位同仁每日件數與 FYC（含所屬小組）
  history_policies.csv  本單位每件保單的年期險種（來自競賽檔）

在 app.py 中：
    from cumulative_section import load_history, render_cumulative_section, today_codes
"""
from __future__ import annotations

import os
import re

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

MY_UNIT = "HC157"
WEEKDAY = "一二三四五六日"

# 沿用戰情室莫蘭迪藍配色
BLUE, BLUE_DEEP, BLUE_SOFT = "#7C97A3", "#51707D", "#A9BFC8"
GOLD, ROSE, SAGE = "#C4A576", "#B98072", "#8FA88C"
TEXT, TEXT_SOFT, BORDER = "#3E4A50", "#83949A", "#DCE6E8"

_V = tuple(int(x) for x in st.__version__.split(".")[:2])
WIDE = {"width": "stretch"} if _V >= (1, 46) else {"use_container_width": True}

SECTION_CSS = """
<style>
.cum-note{color:#83949A;font-size:.92em;margin:-8px 0 10px}
.cum-note b{color:#51707D}
.cum-chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:6px}
.cum-chips span{background:#FFFFFF;border:1px solid #DCE6E8;border-radius:999px;padding:4px 12px;font-size:.92em;color:#3E4A50}
.cum-chips span em{font-style:normal;color:#83949A;font-size:.88em;margin-left:4px}
.cum-code{display:inline-block;font-family:ui-monospace,monospace;font-size:.85em;background:#F7EFE1;color:#3E4A50;border-radius:5px;padding:1px 7px;margin:2px 3px 0 0}
</style>
"""


def _nice(x: float) -> float:
    """把軸的上限取整，刻度才會是 10 萬、20 萬這種整數。"""
    import math
    step = 10 ** math.floor(math.log10(x)) if x > 0 else 1
    return math.ceil(x / step) * step


def _md(d: pd.Timestamp) -> str:
    return f"{d.month}/{d.day}"


def _norm(s) -> str:
    return re.sub(r"[\s　-]+", "", str(s))


def _section(icon: str, title: str):
    st.markdown(f'<div class="morandi-section-title"><i class="ti {icon}"></i><span>{title}</span></div>',
                unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def load_history(folder: str = ".", _mtime: float = 0):
    def read(name):
        p = os.path.join(folder, name)
        if not os.path.exists(p):
            return pd.DataFrame()
        df = pd.read_csv(p)
        for c in ("date", "range_start", "range_end"):
            if c in df:
                df[c] = pd.to_datetime(df[c])
        return df
    return read("history_units.csv"), read("history_people.csv"), read("history_policies.csv")


def history_mtime(folder: str = ".") -> float:
    """讓快取在 CSV 更新時自動失效。"""
    return sum(os.path.getmtime(os.path.join(folder, f)) for f in
               ("history_units.csv", "history_people.csv", "history_policies.csv")
               if os.path.exists(os.path.join(folder, f)))


def today_codes(policies: pd.DataFrame, name: str, date=None) -> list[str]:
    """英雄榜用：某位同仁在最新一天（或指定日）賣的年期險種。"""
    if policies.empty:
        return []
    d = date if date is not None else policies["date"].max()
    m = policies[(policies["date"] == d) & (policies["name"].map(_norm) == _norm(name))]
    return m["code"].tolist()


# ------------------------------------------------------------------ 主區塊
def render_cumulative_section(units: pd.DataFrame, people: pd.DataFrame, policies: pd.DataFrame,
                              unit: str = MY_UNIT, key: str = "cum"):
    if units.empty:
        st.info("還沒有累積資料：請先用 auto_rename.py 處理每日受理業績報表並上傳 history_*.csv")
        return
    st.markdown(SECTION_CSS, unsafe_allow_html=True)

    # 工作月：以最新一天報表上的期間為準
    last_day = units["date"].max()
    lu = units[units["date"] == last_day]
    rs, re_ = lu["range_start"].iloc[0], lu["range_end"].iloc[0]
    u_hist = units[(units["unit"] == unit) & units["date"].between(rs, re_)].sort_values("date")
    p_hist = people[(people["unit"] == unit) & people["date"].between(rs, re_)] if not people.empty else people
    pol = policies[(policies["unit"] == unit) & policies["date"].between(rs, re_)] if not policies.empty else policies
    if u_hist.empty:
        st.info(f"本工作月還沒有 {unit} 的資料")
        return
    U = u_hist.iloc[-1]

    # 日期軸：工作月內的平日，加上有資料的日子
    have = set(u_hist["date"])
    axis = [d for d in pd.date_range(rs, re_) if d.weekday() < 5 or d in have]
    weekdays = len(axis)
    elapsed = sum(1 for d in axis if d <= last_day)
    pace = elapsed / weekdays
    rate = U["cf"] / U["target"] if U["target"] else 0

    by_day = u_hist.set_index("date")
    cum = [by_day["cf"].get(d) if d in have else None for d in axis]
    daily = [by_day["df"].get(d, 0) if d in have else 0 for d in axis]

    _section("ti-chart-line", "本月累積業績走勢")
    gap = (rate - pace) * 100
    st.markdown(
        f'<div class="cum-note">{rs.year - 1911} 年 {re_.month} 月工作月（{_md(rs)}–{_md(re_)}）·'
        f' 達成率 <b>{rate * 100:.1f}%</b>，時間進度 {pace * 100:.1f}%（第 {elapsed}/{weekdays} 個工作天），'
        f'{"超前" if gap >= 0 else "落後"} <b>{abs(gap):.1f}</b> 個百分點 · 游標移到任一天看報件同仁與險種，點一下看明細</div>',
        unsafe_allow_html=True)

    # hover 內容
    def day_people(d):
        if p_hist.empty:
            return p_hist
        m = p_hist[(p_hist["date"] == d) & ((p_hist["dc"] != 0) | (p_hist["df"] != 0))]
        return m.sort_values(["df", "dc"], ascending=False)

    def codes_of(d, name):
        if pol.empty or d not in set(pol["date"]):
            return None
        return pol[(pol["date"] == d) & (pol["name"].map(_norm) == _norm(name))]["code"].tolist()

    hover = []
    for d, c in zip(axis, cum):
        head = f"<b>{_md(d)}（{WEEKDAY[d.weekday()]}）</b>"
        if d not in have:
            hover.append(head + "<br>尚無這天的日報" if d > last_day else head + "<br>沒有這天的日報")
            continue
        r = by_day.loc[d]
        lines = [head, f"當日 {r['dc']:.0f} 件 · FYC {r['df']:,.0f}　累計 {r['cc']:.0f} 件 · FYC {r['cf']:,.0f}", "────────────"]
        dp = day_people(d)
        if dp.empty:
            lines.append("<i>當天沒有報件</i>")
        for row in dp.head(10).itertuples():
            codes = codes_of(d, row.name)
            ctxt = "、".join(codes) if codes else ("競賽檔無此件" if codes == [] else "")
            lines.append(f"<b>{row.name}</b>（{row.team or '—'}） {row.dc:.0f} 件 · {row.df:,.0f}"
                         + (f"<br>　　{ctxt}" if ctxt else ""))
        if len(dp) > 10:
            lines.append(f"<i>另有 {len(dp) - 10} 位</i>")
        hover.append("<br>".join(lines))

    labels = [_md(d) for d in axis]
    last_idx = max((i for i, v in enumerate(cum) if v is not None), default=-1)
    fig = go.Figure()
    fig.add_bar(x=labels, y=daily, name="當日 FYC", yaxis="y2", marker_color=GOLD, opacity=0.6, hoverinfo="skip")
    fig.add_scatter(x=labels, y=[U["target"] * (i + 1) / weekdays for i in range(weekdays)], name="目標進度",
                    mode="lines", line=dict(color=ROSE, dash="dash", width=1.5), hoverinfo="skip")
    fig.add_scatter(x=labels, y=cum, name="累計 FYC", mode="lines+markers", connectgaps=True,
                    line=dict(color=BLUE_DEEP, width=3), fill="tozeroy", fillcolor="rgba(124,151,163,0.15)",
                    marker=dict(size=[10 if i == last_idx else 6 for i in range(len(axis))], color=BLUE_DEEP),
                    customdata=hover, hovertemplate="%{customdata}<extra></extra>")
    fig.update_layout(
        height=430, margin=dict(l=10, r=10, t=30, b=10),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(255,255,255,0.0)",
        font=dict(family="Noto Sans TC, sans-serif", color=TEXT),
        hovermode="x", hoverlabel=dict(bgcolor="white", bordercolor=BORDER, font_size=13, align="left"),
        legend=dict(orientation="h", y=1.1, x=0, font=dict(color=TEXT_SOFT)),
        xaxis=dict(showgrid=False, tickfont=dict(color=TEXT_SOFT), linecolor=BORDER),
        yaxis=dict(gridcolor=BORDER, zeroline=False, tickfont=dict(color=TEXT_SOFT), rangemode="tozero", tickformat=",.0f"),
        yaxis2=dict(overlaying="y", side="right", showgrid=False, tickfont=dict(color=TEXT_SOFT),
                    range=[0, _nice(max(max(daily), 1) * 3)], tickformat=",.0f"),
        dragmode=False,
    )
    event = st.plotly_chart(fig, **WIDE, on_select="rerun", selection_mode="points",
                            key=f"{key}_chart", config={"displayModeBar": False})

    # 點選的那天；沒點就顯示最近一個有報件的日子
    picked = None
    try:
        pts = event.selection.points if event else []
        if pts:
            picked = axis[pts[0]["point_index"]]
    except Exception:
        pass
    if picked is None or picked not in have:
        active = [d for d in sorted(have, reverse=True) if not day_people(d).empty]
        picked = active[0] if active else last_day

    c1, c2 = st.columns([1.3, 1])
    with c1:
        dp = day_people(picked)
        st.markdown(f"**{_md(picked)}（{WEEKDAY[picked.weekday()]}）報件明細**　"
                    f"{len(dp)} 位同仁 · {dp['dc'].sum():.0f} 件 · FYC {dp['df'].sum():,.0f}" if not dp.empty
                    else f"**{_md(picked)} 報件明細**　當天沒有報件")
        if not dp.empty:
            tbl = pd.DataFrame({
                "同仁": dp["name"], "小組": dp["team"].replace("", "—"),
                "商品": [("、".join(c) if c else ("競賽檔無此件" if c == [] else "未匯入競賽檔"))
                         for c in (codes_of(picked, n) for n in dp["name"])],
                "件數": dp["dc"].astype(int), "受理 FYC": dp["df"],
            })
            st.dataframe(tbl, hide_index=True, **WIDE,
                         column_config={"受理 FYC": st.column_config.NumberColumn(format="%,d")})
    with c2:
        st.markdown("**小組進度**　舉績＝壽險件數大於 0")
        latest_p = p_hist[p_hist["date"] == p_hist["date"].max()] if not p_hist.empty else p_hist
        if not latest_p.empty:
            g = (latest_p.assign(team=latest_p["team"].replace("", "—"))
                 .groupby("team").agg(人數=("name", "size"), 舉績=("ju", "sum"), 件數=("cc", "sum"), 累計FYC=("cf", "sum"))
                 .sort_values("累計FYC", ascending=False).reset_index().rename(columns={"team": "小組"}))
            g[["舉績", "件數"]] = g[["舉績", "件數"]].astype(int)
            st.dataframe(g, hide_index=True, **WIDE,
                         column_config={"累計FYC": st.column_config.ProgressColumn(
                             "累計 FYC", format="%,d", min_value=0, max_value=float(g["累計FYC"].max() or 1))})

    # 商品分布
    st.markdown("<br>", unsafe_allow_html=True)
    _section("ti-packages", "本月商品分布")
    if pol.empty:
        st.caption("還沒有競賽檔的保單資料")
    else:
        pol = pol.assign(險種=pol["code"].str.replace(r"^\d+", "", regex=True).replace("", pd.NA).fillna(pol["code"]))
        dist = (pol.groupby("險種").agg(件數=("n", "sum"), 競賽FYC=("fyc", "sum"),
                                        年期=("code", lambda s: "、".join(f"{c}×{k}" for c, k in s.value_counts().items())))
                .sort_values(["件數", "競賽FYC"], ascending=False).reset_index())
        dist["件數"] = dist["件數"].astype(int)
        st.markdown(f'<div class="cum-note">{"、".join(_md(d) for d in sorted(pol["date"].unique()))} 競賽檔 · '
                    f'{int(pol["n"].sum())} 件、{len(dist)} 種險種 · 競賽 FYC 與受理 FYC 計算口徑不同，僅供商品比較</div>',
                    unsafe_allow_html=True)
        d1, d2 = st.columns([1, 1.1])
        with d1:
            st.dataframe(dist, hide_index=True, **WIDE, column_config={
                "件數": st.column_config.ProgressColumn("件數", format="%d 件", min_value=0, max_value=float(dist["件數"].max())),
                "競賽FYC": st.column_config.NumberColumn("競賽 FYC", format="%,d"),
                "年期": st.column_config.TextColumn("年期險種")})
        with d2:
            pick = st.selectbox("看誰賣了這個險種", dist["險種"].tolist(), key=f"{key}_prod")
            who = pol[pol["險種"] == pick].sort_values(["date", "fyc"], ascending=[False, False])
            team_map = dict(zip(p_hist["name"].map(_norm), p_hist["team"])) if not p_hist.empty else {}
            st.dataframe(pd.DataFrame({
                "日期": who["date"].map(_md), "同仁": who["name"],
                "小組": who["name"].map(lambda n: team_map.get(_norm(n)) or "—"),
                "年期險種": who["code"], "競賽 FYC": who["fyc"],
            }), hide_index=True, **WIDE, column_config={"競賽 FYC": st.column_config.NumberColumn(format="%,d")})

    # 尚未舉績
    if not p_hist.empty:
        latest_p = p_hist[p_hist["date"] == p_hist["date"].max()]
        zero = latest_p[latest_p["ju"] == 0]
        st.markdown("<br>", unsafe_allow_html=True)
        _section("ti-user-exclamation", f"本月尚未舉績（{len(zero)} 位）")
        if zero.empty:
            st.success("全員已舉績！")
        else:
            st.markdown('<div class="cum-chips">' + "".join(
                f'<span>{r.name}<em>{r.team or ""}</em></span>' for r in zero.itertuples()) + "</div>",
                unsafe_allow_html=True)

    # 區部單位達成率
    st.markdown("<br>", unsafe_allow_html=True)
    _section("ti-building-community", "竹苗區部 單位達成率")
    board = lu[lu["rank_in_grp"].notna() & (lu["target"] > 0)].copy()
    board["rate"] = board["cf"] / board["target"]
    board = board.sort_values("rate")
    pos = list(board.sort_values("rate", ascending=False)["unit"]).index(unit) + 1 if unit in set(board["unit"]) else None
    if pos:
        st.markdown(f'<div class="cum-note">{unit} 目前第 <b>{pos}</b>/{len(board)} 名 · 紅色虛線為時間進度 {pace * 100:.1f}%</div>',
                    unsafe_allow_html=True)
    bf = go.Figure(go.Bar(
        x=board["rate"], y=board["unit"], orientation="h",
        marker_color=[GOLD if u == unit else BLUE_SOFT for u in board["unit"]],
        text=[f"{r * 100:.1f}%" for r in board["rate"]], textposition="outside",
        customdata=board[["mgr", "cf", "target"]].values,
        hovertemplate="<b>%{y}</b> %{customdata[0]}<br>累計 FYC %{customdata[1]:,.0f} / 目標 %{customdata[2]:,.0f}<extra></extra>"))
    bf.add_vline(x=pace, line=dict(color=ROSE, dash="dash", width=1.5))
    bf.update_layout(height=max(320, 26 * len(board) + 40), margin=dict(l=10, r=40, t=10, b=10),
                     plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                     font=dict(family="Noto Sans TC, sans-serif", color=TEXT),
                     xaxis=dict(tickformat=".0%", gridcolor=BORDER, range=[0, max(board["rate"].max(), pace) * 1.15]),
                     yaxis=dict(tickfont=dict(color=TEXT)), dragmode=False)
    st.plotly_chart(bf, **WIDE, config={"displayModeBar": False}, key=f"{key}_units")
