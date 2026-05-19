import os
import math
import streamlit as st
import pandas as pd
import MetaTrader5 as mt5
import plotly.express as px
import plotly.graph_objects as go
import requests
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from io import BytesIO

# ═══════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════
DD_ALERT_DEFAULT = float(os.environ.get("DD_ALERT_DEFAULT", 10.0))
EXPORT_PATH      = os.environ.get("EXPORT_PATH", r"C:\Thomson\data_export.csv")

# EA Profiles — tambah EA baru di sini
EA_PROFILES = {
    "🧠 Neuro Looping":        {"magic": 8880999, "gv_prefix": "NEURO_",   "label": "Neuro"},
    "⚡ Thomson Liquidity":     {"magic": 888999,  "gv_prefix": "THOMSON_", "label": "Thomson"},
    "🔧 Custom":               {"magic": 0,       "gv_prefix": "",         "label": "Custom"},
}

st.set_page_config(page_title="Multi-EA Dashboard", layout="wide")

# ═══════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════
def classify_session(hour_utc):
    if   12 <= hour_utc < 16: return "New York"
    elif  7 <= hour_utc < 12: return "London"
    elif 16 <= hour_utc < 21: return "New York"
    elif  0 <= hour_utc <  7: return "Asian"
    else:                      return "Off Session"

def parse_comment(comment: str):
    """
    Supports:
      "Bayes+Neuro Buy|LON+BULL"  → (Bayesian, LON, BULL)
      "Grid Buy|ASIA+NEUT"        → (Grid, ASIA, NEUT)
      "Liq Sell|BOS|Bear"         → (Liquidity, BOS, Bear)  [legacy]
    """
    parts    = str(comment).split("|")
    raw_type = parts[0].strip()
    ctx      = parts[1].strip() if len(parts) > 1 else "—"

    if   "Grid"  in raw_type:                          entry_type = "Grid"
    elif "Liq"   in raw_type:                          entry_type = "Liquidity"
    elif "Neuro" in raw_type or "Bayes" in raw_type:   entry_type = "Bayesian"
    else:                                               entry_type = "Other"

    if "+" in ctx:
        p2       = ctx.split("+")
        ms_event = p2[0] if len(p2) > 0 else "—"
        ms_dir   = p2[1] if len(p2) > 1 else "—"
    else:
        ms_event = ctx
        ms_dir   = parts[2].strip() if len(parts) > 2 else "—"

    return entry_type, ms_event, ms_dir

def calc_metrics(d):
    t = len(d)
    if t == 0: return {}
    wins = d[d['net_profit'] > 0]['net_profit']
    loss = d[d['net_profit'] <= 0]['net_profit']
    wr   = len(wins) / t * 100
    gp   = wins.sum()
    gl   = abs(loss.sum())
    pf   = gp / gl if gl > 0 else float('inf')
    aw   = wins.mean() if len(wins) > 0 else 0
    al   = abs(loss.mean()) if len(loss) > 0 else 0
    exp  = (wr/100 * aw) - ((1 - wr/100) * al)
    return {
        'total_trades': t, 'win_rate': wr, 'gross_profit': gp,
        'gross_loss': gl, 'profit_factor': pf, 'expectancy': exp,
        'avg_win': aw, 'avg_loss': al, 'net_profit': d['net_profit'].sum(),
        'max_dd': d['Drawdown'].min(), 'max_dd_pct': d['Drawdown_Pct'].min(),
    }

def duration_str(seconds):
    if seconds is None or (isinstance(seconds, float) and math.isnan(seconds)):
        return "N/A"
    try:
        h, r = divmod(int(seconds), 3600)
        m, s = divmod(r, 60)
        if h > 0:   return f"{h}h {m}m"
        elif m > 0: return f"{m}m {s}s"
        else:       return f"{s}s"
    except Exception:
        return "N/A"

def safe_sec(v):
    return 0.0 if (v is None or (isinstance(v, float) and math.isnan(v))) else v * 60

def to_excel(d):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        d.to_excel(w, index=False, sheet_name='Trades')
    return buf.getvalue()

def add_dd_columns(d):
    d = d.sort_values('time').reset_index(drop=True).copy()
    d['Cumulative_Balance'] = d['net_profit'].cumsum()
    d['Peak']               = d['Cumulative_Balance'].cummax()
    d['Drawdown']           = d['Cumulative_Balance'] - d['Peak']
    d['Drawdown_Pct']       = (d['Drawdown'] / d['Peak'].replace(0, float('nan'))) * 100
    return d

@st.cache_data(ttl=1800)
def get_ff_news():
    try:
        url  = "https://nfs.faireconomy.media/ff_calendar_thisweek.xml"
        resp = requests.get(url, timeout=10)
        root = ET.fromstring(resp.content)
        news = []
        for item in root.findall('.//event'):
            date_str = item.findtext('date', '')
            time_str = item.findtext('time', '')
            try:
                dt_utc = datetime.strptime(f"{date_str} {time_str}", "%m-%d-%Y %I:%M%p")
                dt_wib = dt_utc + timedelta(hours=7)
            except Exception:
                dt_wib = None
            news.append({
                'Time (WIB)': dt_wib,
                'Country':    item.findtext('country', ''),
                'Event':      item.findtext('title', ''),
                'Impact':     item.findtext('impact', ''),
                'Forecast':   item.findtext('forecast', ''),
                'Previous':   item.findtext('previous', ''),
                'Actual':     item.findtext('actual', ''),
            })
        return pd.DataFrame(news)
    except Exception:
        return pd.DataFrame()

# ═══════════════════════════════════════════════
# MT5 DATA
# ═══════════════════════════════════════════════
def get_mt5_data(magic_number: int = 8880999):
    if not mt5.initialize():
        st.error("Failed to connect to MT5. Make sure the MT5 Terminal is open.")
        return pd.DataFrame()
    deals = mt5.history_deals_get(datetime(2024, 1, 1), datetime.now() + timedelta(days=1))
    if not deals:
        mt5.shutdown()
        return pd.DataFrame()
    df = pd.DataFrame(list(deals), columns=deals[0]._asdict().keys())
    df = df[(df['magic'] == magic_number) & (df['entry'] == 1)]
    if not df.empty:
        df['time']       = pd.to_datetime(df['time'], unit='s')
        df['net_profit'] = df['profit'] + df['commission'] + df['swap']
        df = df.sort_values('time').reset_index(drop=True)
        df = add_dd_columns(df)
        df['hour_utc']   = df['time'].dt.hour
        df['Session']    = df['hour_utc'].apply(classify_session)
        df['time_wib']   = df['time'] + timedelta(hours=7)
        df['hour_wib']   = df['time_wib'].dt.hour
        parsed           = df['comment'].apply(parse_comment)
        df['Entry_Type'] = [p[0] for p in parsed]
        df['MS_Event']   = [p[1] for p in parsed]
        df['MS_Dir']     = [p[2] for p in parsed]
        df['duration_sec'] = 0
        try:
            export_p = EXPORT_PATH.replace(".csv", f"_{magic_number}.csv")
            df[['time','symbol','volume','net_profit',
                'Cumulative_Balance','Peak','Drawdown','Drawdown_Pct',
                'Session','time_wib','Entry_Type','MS_Event','MS_Dir']].to_csv(export_p, index=False)
        except Exception:
            pass
    mt5.shutdown()
    return df

def get_trade_duration(df_raw, magic_number: int = 8880999):
    if not mt5.initialize(): return df_raw
    all_deals = mt5.history_deals_get(datetime(2024, 1, 1), datetime.now() + timedelta(days=1))
    mt5.shutdown()
    if not all_deals: return df_raw
    all_df = pd.DataFrame(list(all_deals), columns=all_deals[0]._asdict().keys())
    all_df['time'] = pd.to_datetime(all_df['time'], unit='s')
    opens  = all_df[(all_df['magic'] == magic_number) & (all_df['entry'] == 0)][['position_id','time']].rename(columns={'time':'open_time'})
    if 'position_id' not in df_raw.columns: return df_raw
    closes = df_raw[['position_id','time']].rename(columns={'time':'close_time'})
    merged = closes.merge(opens, on='position_id', how='left')
    merged['duration_sec'] = (merged['close_time'] - merged['open_time']).dt.total_seconds().fillna(0)
    df_raw['duration_sec'] = merged['duration_sec'].values
    return df_raw

# ═══════════════════════════════════════════════
# LOAD DATA
# ═══════════════════════════════════════════════
st.title("📊 Multi-EA Quantitative Journal")
st.markdown("**Performance · Drawdown · Session · Market Structure · News · Remote EA**")

# ═══════════════════════════════════════════════
# SIDEBAR — EA SELECTOR
# ═══════════════════════════════════════════════
st.sidebar.header("⚙️ Dashboard Filter")

# ── EA Profile Selector ──
st.sidebar.subheader("🤖 Select EA")
selected_profile = st.sidebar.selectbox("EA Profile:", list(EA_PROFILES.keys()))
profile = EA_PROFILES[selected_profile]

if profile["magic"] == 0:
    # Custom mode — user input magic number manually
    MAGIC_NUMBER = st.sidebar.number_input("Magic Number:", min_value=1,
                                           max_value=99999999, value=8880999, step=1)
    GV_PREFIX = st.sidebar.text_input("GV Prefix (e.g. NEURO_):", value="NEURO_")
else:
    MAGIC_NUMBER = profile["magic"]
    GV_PREFIX    = profile["gv_prefix"]

st.sidebar.caption(f"Magic: `{MAGIC_NUMBER}` · GV: `{GV_PREFIX or '—'}`")
st.sidebar.divider()

# ── Load data for selected EA ──
@st.cache_data(ttl=30, show_spinner="Loading MT5 data...")
def load_ea_data(magic):
    return get_mt5_data(magic)

df = load_ea_data(MAGIC_NUMBER)
if not df.empty and 'position_id' in df.columns:
    df = get_trade_duration(df, MAGIC_NUMBER)

if df.empty:
    st.warning(f"No data found for **{selected_profile}** (Magic: `{MAGIC_NUMBER}`). "
               f"Make sure the EA has closed at least one trade.")
    st.stop()

selected_period = st.sidebar.selectbox("Chart Aggregation:", ["Daily", "Weekly", "Monthly"])

st.sidebar.divider()
st.sidebar.subheader("📆 Date Filter")
min_date = df['time_wib'].dt.date.min()
max_date = df['time_wib'].dt.date.max()
preset   = st.sidebar.selectbox("Preset:", [
    "Custom", "Today", "Last 7 Days", "Last 30 Days",
    "This Month", "Last Month", "This Year", "All Data"
])
today = datetime.now().date()
if   preset == "Today":         date_start, date_end = today, today
elif preset == "Last 7 Days":   date_start, date_end = today - timedelta(days=6), today
elif preset == "Last 30 Days":  date_start, date_end = today - timedelta(days=29), today
elif preset == "This Month":    date_start, date_end = today.replace(day=1), today
elif preset == "Last Month":
    fe = today.replace(day=1); date_end = fe - timedelta(days=1); date_start = date_end.replace(day=1)
elif preset == "This Year":     date_start, date_end = today.replace(month=1, day=1), today
elif preset == "All Data":      date_start, date_end = min_date, max_date
else:
    _ds = st.sidebar.date_input("From:", value=min_date, min_value=min_date, max_value=max_date)
    _de = st.sidebar.date_input("To:",   value=max_date, min_value=min_date, max_value=max_date)
    date_start = _ds.date() if isinstance(_ds, datetime) else _ds
    date_end   = _de.date() if isinstance(_de, datetime) else _de

if preset != "Custom":
    date_start = max(date_start, min_date)
    date_end   = min(date_end,   max_date)
    st.sidebar.info(f"📅 {date_start.strftime('%d %b %Y')} → {date_end.strftime('%d %b %Y')}")

if date_start > date_end:
    st.sidebar.error("Start date cannot be greater than end date.")
    st.stop()

st.sidebar.divider()
st.sidebar.subheader("🚨 Drawdown Alert")
dd_threshold = st.sidebar.slider("Threshold DD (%)", 1.0, 50.0, DD_ALERT_DEFAULT, 0.5)

# News sidebar
st.sidebar.divider()
st.sidebar.subheader("📰 Today's News")
impact_filter_sb = st.sidebar.multiselect("Filter Impact:", ["High","Medium","Low"], default=["High"])
news_df  = get_ff_news()
now_wib  = datetime.now(timezone.utc).replace(tzinfo=None) + timedelta(hours=7)
today_wib = now_wib.date()
if not news_df.empty and 'Time (WIB)' in news_df.columns:
    news_today = news_df[
        news_df['Time (WIB)'].notna() &
        news_df['Time (WIB)'].apply(lambda x: x.date() if x else None).eq(today_wib) &
        news_df['Impact'].isin(impact_filter_sb)
    ].sort_values('Time (WIB)')
    if news_today.empty:
        st.sidebar.caption("No news events today.")
    else:
        icon = {"High":"🔴","Medium":"🟡","Low":"🟢"}
        for _, row in news_today.iterrows():
            wt = row['Time (WIB)'].strftime('%H:%M') if row['Time (WIB)'] else "-"
            st.sidebar.markdown(f"{icon.get(row['Impact'],'⚪')} **{wt}** [{row['Country']}] {row['Event']}")
else:
    st.sidebar.caption("Failed to load news. Check your internet connection.")

# ═══════════════════════════════════════════════
# FILTER & METRICS
# ═══════════════════════════════════════════════
mask = (df['time_wib'].dt.date >= date_start) & (df['time_wib'].dt.date <= date_end)
dff  = add_dd_columns(df[mask].copy())

if dff.empty:
    st.warning("No data available for the selected period.")
    st.stop()

m = calc_metrics(dff)

current_dd_pct = abs(dff['Drawdown_Pct'].iloc[-1]) if len(dff) > 0 else 0
if current_dd_pct >= dd_threshold:
    st.error(f"🚨 **DRAWDOWN ALERT!** Current DD: **{current_dd_pct:.1f}%** — exceeded threshold {dd_threshold:.1f}%")
elif current_dd_pct >= dd_threshold * 0.8:
    st.warning(f"⚠️ **Approaching DD threshold!** Current DD: **{current_dd_pct:.1f}%** (threshold: {dd_threshold:.1f}%)")

st.caption(f"📆 **{date_start.strftime('%d %b %Y')} – {date_end.strftime('%d %b %Y')}** · {m['total_trades']} trades · EA: **{selected_profile}** (Magic: `{MAGIC_NUMBER}`)")

mc1,mc2,mc3,mc4,mc5,mc6 = st.columns(6)
mc1.metric("Net Profit",    f"${m['net_profit']:.2f}")
mc2.metric("Total Trades",  m['total_trades'])
mc3.metric("Win Rate",      f"{m['win_rate']:.1f}%")
mc4.metric("Profit Factor", f"{m['profit_factor']:.2f}" if m['profit_factor'] != float('inf') else "∞")
mc5.metric("Expectancy",    f"${m['expectancy']:.2f}")
mc6.metric("Max Drawdown",  f"${m['max_dd']:.2f}", f"{m['max_dd_pct']:.1f}%", delta_color="inverse")

st.divider()

# ═══════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════
tab1,tab2,tab3,tab4,tab5,tab6,tab7 = st.tabs([
    "📅 Performance","📉 Drawdown","🌏 Session",
    "🏗️ Market Structure","⏱️ Trade Duration","📰 News","🎮 Remote EA"
])

# ════════════════════════════════════════════════
# TAB 1 — PERFORMANCE
# ════════════════════════════════════════════════
with tab1:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader(f"Profit per {selected_period}")
        if selected_period == "Daily":
            g = dff.groupby(dff['time'].dt.date)['net_profit'].sum().reset_index()
        elif selected_period == "Weekly":
            g = dff.groupby(dff['time'].dt.to_period('W').astype(str))['net_profit'].sum().reset_index()
        else:
            g = dff.groupby(dff['time'].dt.to_period('M').astype(str))['net_profit'].sum().reset_index()
        st.plotly_chart(px.bar(g, x=g.columns[0], y='net_profit', color='net_profit',
                               color_continuous_scale='RdYlGn', labels={'net_profit':'Net Profit ($)'}),
                        width='stretch')
    with c2:
        st.subheader("Balance Growth (Cumulative)")
        fig2 = px.line(dff, x='time', y='Cumulative_Balance',
                       labels={'Cumulative_Balance':'Balance ($)','time':'Time'})
        fig2.update_traces(line_color='#2ecc71', fill='tozeroy')
        st.plotly_chart(fig2, width='stretch')

    st.divider()
    cp, cd = st.columns([1,2])
    with cp:
        st.subheader("📋 Pair Profit")
        pd2 = dff.groupby('symbol')['net_profit'].sum().reset_index().sort_values('net_profit', ascending=False)
        st.plotly_chart(px.bar(pd2, x='symbol', y='net_profit', color='net_profit',
                               color_continuous_scale='RdYlGn',
                               labels={'net_profit':'Net Profit ($)','symbol':'Pair'}), width='stretch')
    with cd:
        st.subheader("📝 Last History")
        display_cols = ['time_wib','symbol','volume','net_profit','Session','MS_Event','MS_Dir']
        if dff.get('duration_sec', pd.Series([0])).sum() > 0:
            dff['Duration'] = dff['duration_sec'].apply(duration_str)
            display_cols.append('Duration')
        st.dataframe(dff[display_cols].sort_values('time_wib', ascending=False),
                     width='stretch', height=300)

    st.divider()
    st.subheader("📥 Export Data")
    ex1, ex2 = st.columns(2)
    export_df = dff[display_cols].sort_values('time_wib', ascending=False)
    with ex1:
        st.download_button("⬇️ Download CSV",
                           data=export_df.to_csv(index=False).encode('utf-8'),
                           file_name=f"neuro_ea_{date_start}_{date_end}.csv",
                           mime='text/csv')
    with ex2:
        try:
            st.download_button("⬇️ Download Excel", data=to_excel(export_df),
                               file_name=f"neuro_ea_{date_start}_{date_end}.xlsx",
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        except Exception:
            st.info("Run `pip install openpyxl` to enable Excel export.")

    st.divider()
    st.subheader("🔄 Comparison Mode — Compare 2 Periods")

    def to_date(v):
        return v.date() if hasattr(v, 'date') and callable(v.date) and not isinstance(v, type) else v

    mid_date = min_date + (max_date - min_date) // 2
    cmp1, cmp2 = st.columns(2)
    with cmp1:
        cs1 = to_date(st.date_input("Period A — From:", value=min_date, min_value=min_date, max_value=max_date, key='cs1'))
        ce1 = to_date(st.date_input("Period A — To:",   value=mid_date, min_value=min_date, max_value=max_date, key='ce1'))
    with cmp2:
        cs2 = to_date(st.date_input("Period B — From:", value=mid_date, min_value=min_date, max_value=max_date, key='cs2'))
        ce2 = to_date(st.date_input("Period B — To:",   value=max_date, min_value=min_date, max_value=max_date, key='ce2'))

    dA = add_dd_columns(df[(df['time_wib'].dt.date >= cs1) & (df['time_wib'].dt.date <= ce1)].copy())
    dB = add_dd_columns(df[(df['time_wib'].dt.date >= cs2) & (df['time_wib'].dt.date <= ce2)].copy())
    mA, mB = calc_metrics(dA), calc_metrics(dB)

    if mA and mB:
        def fmt(mx):
            return [f"${mx['net_profit']:.2f}", mx['total_trades'], f"{mx['win_rate']:.1f}%",
                    f"{mx['profit_factor']:.2f}" if mx['profit_factor'] != float('inf') else "∞",
                    f"${mx['expectancy']:.2f}", f"${mx['max_dd']:.2f}", f"{mx['max_dd_pct']:.1f}%"]
        cmp_df = pd.DataFrame({
            'Metric': ['Net Profit ($)','Total Trades','Win Rate (%)','Profit Factor','Expectancy ($)','Max DD ($)','Max DD (%)'],
            f"A ({cs1.strftime('%d %b')}–{ce1.strftime('%d %b')})": fmt(mA),
            f"B ({cs2.strftime('%d %b')}–{ce2.strftime('%d %b')})": fmt(mB),
        })
        st.dataframe(cmp_df.astype(str), width='stretch', hide_index=True)
        fig_cmp = go.Figure()
        for d, lbl, clr in [(dA,"Period A","#2ecc71"),(dB,"Period B","#3498db")]:
            if len(d) == 0: continue
            d2 = d.sort_values('time').copy(); d2['cum'] = d2['net_profit'].cumsum()
            fig_cmp.add_trace(go.Scatter(x=list(range(len(d2))), y=d2['cum'], name=lbl, line=dict(color=clr, width=2)))
        fig_cmp.update_layout(xaxis_title='Trade #', yaxis_title='Cumulative Profit ($)',
                              hovermode='x unified', margin=dict(t=20))
        st.plotly_chart(fig_cmp, width='stretch')

# ════════════════════════════════════════════════
# TAB 2 — DRAWDOWN
# ════════════════════════════════════════════════
with tab2:
    st.subheader("📉 Drawdown Chart (Peak-to-Trough)")
    fig_dd = go.Figure()
    fig_dd.add_trace(go.Scatter(x=dff['time'], y=dff['Drawdown'],
                                fill='tozeroy', fillcolor='rgba(231,76,60,0.2)',
                                line=dict(color='#e74c3c', width=1.5), name='Drawdown ($)',
                                hovertemplate='%{x}<br>DD: $%{y:.2f}<extra></extra>'))
    idx = dff['Drawdown'].idxmin()
    fig_dd.add_trace(go.Scatter(x=[dff.loc[idx,'time']], y=[dff.loc[idx,'Drawdown']],
                                mode='markers+text', marker=dict(color='#c0392b', size=10),
                                text=[f"  Max DD: ${dff.loc[idx,'Drawdown']:.2f} ({dff.loc[idx,'Drawdown_Pct']:.1f}%)"],
                                textposition='top right', textfont=dict(size=11, color='#c0392b'),
                                name='Max Drawdown'))
    if dff['Peak'].max() > 0:
        fig_dd.add_hline(y=-dd_threshold/100*dff['Peak'].max(), line_dash='dash',
                         line_color='orange', line_width=1.5,
                         annotation_text=f"Alert {dd_threshold:.0f}%", annotation_position="right")
    fig_dd.add_hline(y=0, line_dash='dot', line_color='gray', line_width=1)
    fig_dd.update_layout(yaxis_title='Drawdown ($)', xaxis_title='Time',
                         hovermode='x unified', margin=dict(t=30))
    st.plotly_chart(fig_dd, width='stretch')

    st.subheader("📊 Drawdown Summary")
    st.dataframe(pd.DataFrame({
        'Metric': ['Max DD ($)','Max DD (%)','Current DD ($)','Current DD (%)','Alert Threshold (%)'],
        'Value':  [f"${dff['Drawdown'].min():.2f}", f"{dff['Drawdown_Pct'].min():.2f}%",
                   f"${dff['Drawdown'].iloc[-1]:.2f}", f"{dff['Drawdown_Pct'].iloc[-1]:.2f}%",
                   f"{dd_threshold:.1f}%"]
    }), width='stretch', hide_index=True)

# ════════════════════════════════════════════════
# TAB 3 — SESSION
# ════════════════════════════════════════════════
with tab3:
    st.subheader("🌏 Win Rate per Trading Session")
    st.caption("WIB (UTC+7) · Asian 07–16 · London 14–23 · New York 19–04")
    s_colors = {"Asian":"#f39c12","London":"#2980b9","New York":"#27ae60","Off Session":"#7f8c8d"}
    sess_stats = []
    for s in ["Asian","London","New York","Off Session"]:
        sd = dff[dff['Session']==s]
        if not len(sd): continue
        t=len(sd); w=len(sd[sd['net_profit']>0])
        sess_stats.append({'Session':s,'Total Trades':t,'Win':w,'Loss':t-w,
                           'Win Rate (%)':round(w/t*100,1),
                           'Net Profit ($)':round(sd['net_profit'].sum(),2),
                           'Avg Profit ($)':round(sd['net_profit'].mean(),2)})
    sdf2 = pd.DataFrame(sess_stats)
    cols = st.columns(len(sdf2))
    for i, row in sdf2.iterrows():
        clr = s_colors.get(row['Session'],'#888')
        cols[i].markdown(f"""<div style="background:#1e1e1e;border-left:4px solid {clr};
            border-radius:8px;padding:12px 16px;margin-bottom:8px;">
            <div style="font-size:13px;color:{clr};font-weight:600;">{row['Session']}</div>
            <div style="font-size:28px;font-weight:700;color:#fff;margin:4px 0;">{row['Win Rate (%)']}%</div>
            <div style="font-size:12px;color:#aaa;">{row['Win']}W / {row['Loss']}L · {row['Total Trades']} trades</div>
            <div style="font-size:12px;color:#aaa;margin-top:2px;">Net: <b style="color:#fff;">${row['Net Profit ($)']}</b></div>
        </div>""", unsafe_allow_html=True)
    st.divider()
    sc1, sc2 = st.columns(2)
    with sc1:
        fw = px.bar(sdf2, x='Session', y='Win Rate (%)', color='Session',
                    color_discrete_map=s_colors, text='Win Rate (%)')
        fw.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
        fw.update_layout(yaxis=dict(range=[0,115]), showlegend=False, margin=dict(t=20))
        st.plotly_chart(fw, width='stretch')
    with sc2:
        fn = px.bar(sdf2, x='Session', y='Net Profit ($)', color='Net Profit ($)',
                    color_continuous_scale='RdYlGn', text='Net Profit ($)')
        fn.update_traces(texttemplate='$%{text:.2f}', textposition='outside')
        fn.update_layout(showlegend=False, margin=dict(t=20))
        st.plotly_chart(fn, width='stretch')
    st.divider()
    st.subheader("📋 Session Detail Table")
    st.dataframe(sdf2, width='stretch', hide_index=True)
    st.divider()
    st.subheader("⏰ Trade Distribution by Hour (WIB)")
    hourly = dff.groupby('hour_wib').agg(
        Trades=('net_profit','count'),
        WinRate=('net_profit', lambda x: (x>0).mean()*100),
        NetProfit=('net_profit','sum')
    ).reset_index()
    fh = px.bar(hourly, x='hour_wib', y='WinRate', color='WinRate',
                color_continuous_scale='RdYlGn',
                hover_data={'Trades':True,'NetProfit':':.2f'},
                labels={'hour_wib':'Hour (WIB)','WinRate':'Win Rate (%)'}, text='Trades')
    fh.update_traces(texttemplate='%{text} trades', textposition='outside')
    fh.update_layout(xaxis=dict(tickmode='linear',dtick=1), yaxis=dict(range=[0,115]), margin=dict(t=20))
    for x0,x1,c,lbl in [(6.5,15.5,"#f39c12","Asian"),(13.5,22.5,"#2980b9","London"),(18.5,23.5,"#27ae60","New York")]:
        fh.add_vrect(x0=x0, x1=x1, fillcolor=c, opacity=0.07,
                     annotation_text=lbl, annotation_position="top left")
    st.plotly_chart(fh, width='stretch')

# ════════════════════════════════════════════════
# TAB 4 — MARKET STRUCTURE
# ════════════════════════════════════════════════
with tab4:
    st.subheader("🏗️ Market Structure Analysis (BOS & CHoCH / Neuro Context)")
    st.caption("Comment format: EntryType|SESSION+BIAS or EntryType|BOS|Dir (legacy)")
    ms_colors  = {"BOS":"#2980b9","CHoCH":"#e67e22","NONE":"#7f8c8d",
                  "LON":"#2980b9","ASIA":"#f39c12","USA":"#27ae60","—":"#7f8c8d"}
    dir_colors = {"Bull":"#27ae60","BULL":"#27ae60","Bear":"#e74c3c","BEAR":"#e74c3c",
                  "None":"#7f8c8d","NEUT":"#7f8c8d","—":"#7f8c8d"}
    type_colors= {"Liquidity":"#9b59b6","Bayesian":"#2ecc71","Grid":"#e67e22","Other":"#95a5a6"}

    bos_df   = dff[dff['MS_Event']=='BOS']
    choch_df = dff[dff['MS_Event']=='CHoCH']
    none_df  = dff[~dff['MS_Event'].isin(['BOS','CHoCH'])]

    mc1,mc2,mc3 = st.columns(3)
    for col, sub, clr, label in [
        (mc1,bos_df,"#2980b9","BOS"),
        (mc2,choch_df,"#e67e22","CHoCH"),
        (mc3,none_df,"#7f8c8d","Other / Neuro Ctx")
    ]:
        t=len(sub); w=len(sub[sub['net_profit']>0]) if t>0 else 0
        wr=w/t*100 if t>0 else 0; net=sub['net_profit'].sum() if t>0 else 0
        avg=sub['net_profit'].mean() if t>0 else 0
        col.markdown(f"""<div style="background:#1e1e1e;border-left:4px solid {clr};
            border-radius:8px;padding:14px 18px;margin-bottom:10px;">
            <div style="font-size:14px;color:{clr};font-weight:700;">{label}</div>
            <div style="font-size:30px;font-weight:800;color:#fff;margin:6px 0;">{wr:.1f}%</div>
            <div style="font-size:12px;color:#aaa;">{w}W / {t-w}L · {t} trades</div>
            <div style="font-size:12px;color:#aaa;margin-top:2px;">
                Net: <b style="color:#fff;">${net:.2f}</b> · Avg: <b style="color:#fff;">${avg:.2f}</b>
            </div></div>""", unsafe_allow_html=True)

    st.divider()
    r1c1,r1c2 = st.columns(2)
    with r1c1:
        st.subheader("Win Rate by MS Event / Context")
        ms_stats=[]
        for evt in dff['MS_Event'].unique():
            sub=dff[dff['MS_Event']==evt]
            t=len(sub); w=len(sub[sub['net_profit']>0])
            ms_stats.append({'MS Event':evt,'Total':t,'Win':w,'Loss':t-w,
                             'Win Rate (%)':round(w/t*100,1),'Net Profit ($)':round(sub['net_profit'].sum(),2)})
        ms_df=pd.DataFrame(ms_stats).sort_values('Win Rate (%)',ascending=False)
        fms=px.bar(ms_df,x='MS Event',y='Win Rate (%)',color='MS Event',text='Win Rate (%)')
        fms.update_traces(texttemplate='%{text:.1f}%',textposition='outside')
        fms.update_layout(yaxis=dict(range=[0,115]),showlegend=False,margin=dict(t=20))
        st.plotly_chart(fms,width='stretch')
    with r1c2:
        st.subheader("Net Profit by MS Event / Context")
        fmsp=px.bar(ms_df,x='MS Event',y='Net Profit ($)',color='Net Profit ($)',
                    color_continuous_scale='RdYlGn',text='Net Profit ($)')
        fmsp.update_traces(texttemplate='$%{text:.2f}',textposition='outside')
        fmsp.update_layout(showlegend=False,margin=dict(t=20))
        st.plotly_chart(fmsp,width='stretch')

    st.divider()
    r2c1,r2c2 = st.columns(2)
    with r2c1:
        st.subheader("Win Rate by Direction / Bias")
        dir_stats=[]
        for d in dff['MS_Dir'].unique():
            sub=dff[dff['MS_Dir']==d]
            t=len(sub); w=len(sub[sub['net_profit']>0])
            dir_stats.append({'Direction':d,'Win Rate (%)':round(w/t*100,1),
                              'Net Profit ($)':round(sub['net_profit'].sum(),2),'Trades':t})
        if dir_stats:
            fd=px.bar(pd.DataFrame(dir_stats),x='Direction',y='Win Rate (%)',
                      color='Direction',text='Win Rate (%)',hover_data=['Trades','Net Profit ($)'])
            fd.update_traces(texttemplate='%{text:.1f}%',textposition='outside')
            fd.update_layout(yaxis=dict(range=[0,115]),showlegend=False,margin=dict(t=20))
            st.plotly_chart(fd,width='stretch')
    with r2c2:
        st.subheader("Entry Type Distribution")
        etype=dff.groupby('Entry_Type').agg(
            Trades=('net_profit','count'),
            WinRate=('net_profit',lambda x:(x>0).mean()*100),
            NetProfit=('net_profit','sum')
        ).reset_index()
        fe=px.bar(etype,x='Entry_Type',y='WinRate',color='Entry_Type',
                  color_discrete_map=type_colors,text='Trades',hover_data={'NetProfit':':.2f'},
                  labels={'Entry_Type':'Entry Type','WinRate':'Win Rate (%)'})
        fe.update_traces(texttemplate='%{text} trades',textposition='outside')
        fe.update_layout(yaxis=dict(range=[0,115]),showlegend=False,margin=dict(t=20))
        st.plotly_chart(fe,width='stretch')

    st.divider()
    r3c1,r3c2 = st.columns(2)
    with r3c1:
        st.subheader("MS Event per Symbol")
        sym_ms=dff.groupby(['symbol','MS_Event']).size().reset_index(name='Trades')
        st.plotly_chart(px.bar(sym_ms,x='symbol',y='Trades',color='MS_Event',
                               barmode='group',labels={'symbol':'Symbol','MS_Event':'Context'}),
                        width='stretch')
    with r3c2:
        st.subheader("📈 Cumulative Profit by Entry Type")
        dff_s=dff.sort_values('time').copy()
        fig_cum=go.Figure()
        for et,clr in [("Bayesian","#2ecc71"),("Liquidity","#9b59b6"),("Grid","#f39c12"),("Other","#95a5a6")]:
            sub=dff_s[dff_s['Entry_Type']==et].copy()
            if not len(sub): continue
            sub['cum']=sub['net_profit'].cumsum()
            fig_cum.add_trace(go.Scatter(x=sub['time'],y=sub['cum'],name=et,line=dict(color=clr,width=2)))
        fig_cum.add_hline(y=0,line_dash='dot',line_color='gray',line_width=1)
        fig_cum.update_layout(yaxis_title='Cumulative Profit ($)',xaxis_title='Time',
                              hovermode='x unified',legend=dict(orientation='h',y=1.05),margin=dict(t=40))
        st.plotly_chart(fig_cum,width='stretch')

    st.divider()
    st.subheader("📋 MS Detail Table")
    st.dataframe(dff[['time_wib','symbol','volume','net_profit','Entry_Type','MS_Event','MS_Dir','Session']
                    ].sort_values('time_wib',ascending=False), width='stretch', height=350)

# ════════════════════════════════════════════════
# TAB 5 — TRADE DURATION
# ════════════════════════════════════════════════
with tab5:
    st.subheader("⏱️ Trade Duration Analysis")
    has_dur = 'duration_sec' in dff.columns and dff['duration_sec'].sum() > 0
    if not has_dur:
        st.info("Duration data not available. Make sure `position_id` exists in MT5 history and the EA has been recompiled.")
    else:
        dff['Duration_Min'] = dff['duration_sec'] / 60
        dff['Duration_Str'] = dff['duration_sec'].apply(duration_str)
        avg_dur      = dff['Duration_Min'].mean()
        avg_dur_win  = dff[dff['net_profit']>0]['Duration_Min'].mean()
        avg_dur_loss = dff[dff['net_profit']<=0]['Duration_Min'].mean()

        d1,d2,d3 = st.columns(3)
        d1.metric("Avg Duration (All)",  duration_str(safe_sec(avg_dur)))
        d2.metric("Avg Duration (Win)",  duration_str(safe_sec(avg_dur_win)))
        d3.metric("Avg Duration (Loss)",
                  duration_str(safe_sec(avg_dur_loss))
                  if not (avg_dur_loss is None or (isinstance(avg_dur_loss,float) and math.isnan(avg_dur_loss)))
                  else "N/A (no losses yet)")

        st.divider()
        dc1,dc2 = st.columns(2)
        with dc1:
            st.subheader("Trade Duration Distribution")
            st.plotly_chart(px.histogram(dff, x='Duration_Min', nbins=30,
                                         color_discrete_sequence=['#3498db'],
                                         labels={'Duration_Min':'Duration (min)'}), width='stretch')
        with dc2:
            st.subheader("Duration vs Profit")
            fig_sc = px.scatter(dff, x='Duration_Min', y='net_profit',
                                color='net_profit', color_continuous_scale='RdYlGn',
                                hover_data=['symbol','Duration_Str'],
                                labels={'Duration_Min':'Duration (min)','net_profit':'Net Profit ($)'})
            fig_sc.add_hline(y=0, line_dash='dot', line_color='gray')
            st.plotly_chart(fig_sc, width='stretch')

        st.divider()
        st.subheader("Avg Duration per Session")
        dur_sess = dff.groupby('Session').agg(
            Avg_Min=('Duration_Min','mean'), Trades=('net_profit','count')
        ).reset_index()
        dur_sess['Avg_Dur_Str'] = (dur_sess['Avg_Min']*60).apply(duration_str)
        fig_ds = px.bar(dur_sess, x='Session', y='Avg_Min', color='Session',
                        color_discrete_map={"Asian":"#f39c12","London":"#2980b9",
                                            "New York":"#27ae60","Off Session":"#7f8c8d"},
                        text='Avg_Dur_Str', labels={'Avg_Min':'Avg Duration (min)'})
        fig_ds.update_traces(textposition='outside')
        fig_ds.update_layout(showlegend=False, margin=dict(t=20))
        st.plotly_chart(fig_ds, width='stretch')

# ════════════════════════════════════════════════
# TAB 6 — NEWS
# ════════════════════════════════════════════════
with tab6:
    st.subheader("📰 Economic Calendar — ForexFactory")
    impact_filter = st.multiselect("Filter Impact:", ["High","Medium","Low"],
                                   default=["High","Medium"], key='news_tab')
    country_list  = sorted(news_df['Country'].dropna().unique().tolist()) if not news_df.empty else []
    country_filter = st.multiselect("Filter Country:", country_list,
                                    default=[c for c in ['USD','EUR','GBP','JPY','XAU'] if c in country_list])

    if news_df.empty:
        st.warning("Failed to load news from ForexFactory. Check the VPS internet connection.")
    else:
        fn = news_df.copy()
        if impact_filter:  fn = fn[fn['Impact'].isin(impact_filter)]
        if country_filter: fn = fn[fn['Country'].isin(country_filter)]
        fn = fn.sort_values('Time (WIB)')

        st.subheader("🔜 Upcoming News")
        upcoming = fn[fn['Time (WIB)'] >= now_wib].head(10)
        if upcoming.empty:
            st.info("No upcoming news for the selected filters.")
        else:
            ic = {"High":"#e74c3c","Medium":"#f39c12","Low":"#2ecc71"}
            for _, row in upcoming.iterrows():
                clr = ic.get(row['Impact'],"#888")
                wt  = row['Time (WIB)'].strftime('%a %d %b %H:%M WIB') if row['Time (WIB)'] else "-"
                delta = row['Time (WIB)'] - now_wib
                hrs = int(delta.total_seconds()//3600); mins = int((delta.total_seconds()%3600)//60)
                countdown = f"{hrs}h {mins}m away" if hrs > 0 else f"{mins}m away"
                st.markdown(f"""<div style="background:#1e1e1e;border-left:4px solid {clr};
                    border-radius:8px;padding:10px 14px;margin-bottom:6px;
                    display:flex;justify-content:space-between;align-items:center;">
                    <div>
                        <span style="color:{clr};font-weight:700;font-size:12px;">● {row['Impact'].upper()}</span>
                        &nbsp;<span style="color:#aaa;font-size:12px;">[{row['Country']}]</span>
                        &nbsp;<span style="color:#fff;font-size:13px;font-weight:500;">{row['Event']}</span>
                    </div>
                    <div style="text-align:right;min-width:160px;">
                        <div style="color:#aaa;font-size:11px;">{wt}</div>
                        <div style="color:{clr};font-size:11px;font-weight:600;">{countdown}</div>
                    </div></div>""", unsafe_allow_html=True)

        st.divider()
        st.subheader("✅ News Result (Released)")
        released = fn[fn['Time (WIB)'] < now_wib].sort_values('Time (WIB)', ascending=False).head(20).copy()
        if released.empty:
            st.info("No news released yet.")
        else:
            released['Time (WIB)'] = released['Time (WIB)'].apply(
                lambda x: x.strftime('%a %d %b %H:%M') if x else "-")
            st.dataframe(released[['Time (WIB)','Country','Impact','Event','Forecast','Previous','Actual']],
                         width='stretch', hide_index=True)
        st.divider()
        st.caption("Data from ForexFactory RSS · Auto-refresh every 30 minutes · Time in WIB (UTC+7)")
        if st.button("🔄 Refresh News"):
            st.cache_data.clear()
            st.rerun()

# ════════════════════════════════════════════════
# TAB 7 — REMOTE EA
# ════════════════════════════════════════════════
with tab7:
    st.subheader(f"🎮 Remote EA Control — {selected_profile}")
    st.caption(f"Commands sent via MT5 GlobalVariables · Magic: `{MAGIC_NUMBER}` · GV Prefix: `{GV_PREFIX}`")

    # GV keys built dynamically from selected EA profile prefix
    p = GV_PREFIX  # e.g. "NEURO_" or "THOMSON_" or custom
    GV_CMD          = p + "CMD"
    GV_LOT          = p + "LOT"
    GV_TP           = p + "TP"
    GV_DD           = p + "DD"
    GV_TIME_EN      = p + "TIME_EN"
    GV_START_H      = p + "START_H"
    GV_START_M      = p + "START_M"
    GV_END_H        = p + "END_H"
    GV_END_M        = p + "END_M"
    GV_STATUS       = p + "STATUS"
    GV_POSITIONS    = p + "POSITIONS"
    GV_EQUITY       = p + "EQUITY"
    GV_BALANCE      = p + "BALANCE"
    GV_BAYESPROB    = p + "BAYESPROB"
    GV_MTFSENTIMENT = p + "SENTIMENT"
    GV_HTFBIAS      = p + "HTFBIAS"
    GV_CTX_THR      = p + "CTX_THR"
    GV_CTX_EXITMULT = p + "CTX_EXITMULT"
    GV_CTX_STREAK   = p + "CTX_STREAK"
    GV_CTX_PAUSED   = p + "CTX_PAUSED"
    GV_HEARTBEAT    = p + "HEARTBEAT"

    if not GV_PREFIX:
        st.info("ℹ️ Remote EA control requires a GV Prefix. Set it in the sidebar under EA Profile > Custom.")
        st.stop()

    def gv_get(key, default=0.0):
        try:
            val = mt5.global_variable_get(key)
            return val if val is not None else default
        except Exception:
            return default

    def gv_set(key, value):
        try:
            if not mt5.initialize(): return False
            result = mt5.global_variable_set(key, float(value))
            mt5.shutdown()
            return result
        except Exception:
            return False

    if not mt5.initialize():
        st.error("MT5 not connected. Make sure MT5 Terminal is running on this machine.")
    else:
        heartbeat       = gv_get(GV_HEARTBEAT)
        ea_status       = int(gv_get(GV_STATUS))
        ea_positions    = int(gv_get(GV_POSITIONS))
        ea_equity       = gv_get(GV_EQUITY)
        ea_balance      = gv_get(GV_BALANCE)
        ea_prob         = gv_get(GV_BAYESPROB)
        ea_sentiment    = int(gv_get(GV_MTFSENTIMENT))
        ea_htfbias      = int(gv_get(GV_HTFBIAS))
        ea_ctx_thr      = gv_get(GV_CTX_THR)
        ea_ctx_exitmult = gv_get(GV_CTX_EXITMULT)
        ea_ctx_streak   = int(gv_get(GV_CTX_STREAK))
        ea_ctx_paused   = int(gv_get(GV_CTX_PAUSED))
        mt5.shutdown()

        now_ts   = datetime.now(timezone.utc).timestamp()
        hb_age   = now_ts - heartbeat if heartbeat > 0 else 999
        ea_online = hb_age < 30

        status_map = {0:("🟢 Running","#27ae60"),1:("🟡 Paused","#f39c12"),2:("🔴 Stopped","#e74c3c")}
        status_label, status_color = status_map.get(ea_status, ("❓ Unknown","#888"))

        st.markdown("### 📡 EA Live Status")
        if not ea_online:
            st.warning(f"⚠️ EA heartbeat not detected for {hb_age:.0f}s — EA may be offline or not compiled yet.")

        s1,s2,s3,s4,s5,s6 = st.columns(6)
        s1.markdown(f"""<div style="background:#1e1e1e;border-left:4px solid {status_color};
            border-radius:8px;padding:10px 14px;">
            <div style="font-size:11px;color:#aaa;">Status</div>
            <div style="font-size:18px;font-weight:700;color:#fff;">{status_label}</div>
        </div>""", unsafe_allow_html=True)
        s2.metric("Open Positions", ea_positions)
        s3.metric("Equity",         f"${ea_equity:.2f}")
        s4.metric("Balance",        f"${ea_balance:.2f}")
        s5.metric("Bayes Prob",     f"{ea_prob:.3f}")
        sentiment_map = {2:"🚀 Risk-On",1:"🟢 Bullish",0:"⚪ Neutral",-1:"🔴 Bearish",-2:"🔻 Risk-Off"}
        htfbias_map   = {1:"HTF Bull",0:"HTF Neut",-1:"HTF Bear"}
        s6.markdown(f"""<div style="background:#1e1e1e;border-left:4px solid #888;
            border-radius:8px;padding:10px 14px;">
            <div style="font-size:11px;color:#aaa;">MTF Sentiment</div>
            <div style="font-size:13px;font-weight:600;color:#fff;">{sentiment_map.get(ea_sentiment,'—')}</div>
            <div style="font-size:11px;color:#aaa;">{htfbias_map.get(ea_htfbias,'—')}</div>
        </div>""", unsafe_allow_html=True)

        n1,n2,n3,n4 = st.columns(4)
        n1.metric("Ctx Threshold",  f"{ea_ctx_thr:.3f}",   help="dynThreshold for active context")
        n2.metric("Ctx Exit Mult",  f"{ea_ctx_exitmult:.2f}", help="dynExitMult for active context")
        n3.metric("Loss Streak",    ea_ctx_streak,          help="Consecutive losses in active context")
        n4.metric("Ctx Paused",     "YES ⚠️" if ea_ctx_paused else "NO ✅", help="Context paused by Neuro")

        st.divider()
        st.markdown("### ⚡ Quick Commands")
        st.caption("Commands are picked up by EA on the next tick.")

        cb1,cb2,cb3 = st.columns(3)
        with cb1:
            if st.button("▶️ Resume EA", use_container_width=True, type="primary"):
                st.success("✅ RESUME command sent.") if gv_set(GV_CMD,2) else st.error("❌ Failed.")
        with cb2:
            if st.button("⏸️ Pause EA", use_container_width=True):
                st.success("✅ PAUSE command sent.") if gv_set(GV_CMD,1) else st.error("❌ Failed.")
        with cb3:
            confirm = st.checkbox("I confirm — Close ALL open trades")
            if st.button("🛑 Close ALL Trades", use_container_width=True,
                         type="secondary", disabled=not confirm):
                st.success("✅ CLOSE ALL command sent.") if gv_set(GV_CMD,3) else st.error("❌ Failed.")

        st.divider()
        st.markdown("### ⚙️ Parameter Override")
        st.caption("Set to 0 to revert to EA input values.")
        po1, po2 = st.columns(2)

        with po1:
            st.markdown("**Lot Size Override**")
            lot_val = st.number_input("Lot Size (0 = EA default)", 0.0, 100.0,
                                      float(gv_get(GV_LOT)), 0.01, "%.2f")
            if st.button("Apply Lot Size", use_container_width=True):
                st.success(f"✅ Lot → {lot_val:.2f}" if lot_val>0 else "✅ Lot reset to EA default.") \
                    if gv_set(GV_LOT, lot_val) else st.error("❌ Failed.")

            st.markdown("**Global TP Override ($)**")
            tp_val = st.number_input("Global TP (0 = EA default)", 0.0, 100000.0,
                                     float(gv_get(GV_TP)), 10.0)
            if st.button("Apply Global TP", use_container_width=True):
                st.success(f"✅ TP → ${tp_val:.2f}" if tp_val>0 else "✅ TP reset to EA default.") \
                    if gv_set(GV_TP, tp_val) else st.error("❌ Failed.")

        with po2:
            st.markdown("**Max Drawdown Override (%)**")
            dd_val = st.number_input("Max DD % (0 = EA default)", 0.0, 100.0,
                                     float(gv_get(GV_DD)), 0.5)
            if st.button("Apply Max DD", use_container_width=True):
                st.success(f"✅ Max DD → {dd_val:.1f}%" if dd_val>0 else "✅ Max DD reset to EA default.") \
                    if gv_set(GV_DD, dd_val) else st.error("❌ Failed.")

            st.markdown("**Time Filter Override**")
            time_en = st.toggle("Enable Time Filter", value=bool(gv_get(GV_TIME_EN,1)))
            tc1,tc2 = st.columns(2)
            with tc1:
                start_h = st.number_input("Start Hour (UTC)", 0, 23, int(gv_get(GV_START_H,8)))
                start_m = st.number_input("Start Min",        0, 59, int(gv_get(GV_START_M,0)))
            with tc2:
                end_h   = st.number_input("End Hour (UTC)",   0, 23, int(gv_get(GV_END_H,22)))
                end_m   = st.number_input("End Min",          0, 59, int(gv_get(GV_END_M,0)))
            if st.button("Apply Time Filter", use_container_width=True):
                ok = all([gv_set(GV_TIME_EN, 1 if time_en else 0),
                          gv_set(GV_START_H, start_h), gv_set(GV_START_M, start_m),
                          gv_set(GV_END_H,   end_h),   gv_set(GV_END_M,   end_m)])
                st.success(f"✅ Time filter {'enabled' if time_en else 'disabled'}: {start_h:02d}:{start_m:02d}–{end_h:02d}:{end_m:02d} UTC") \
                    if ok else st.error("❌ Failed.")

        st.divider()
        st.markdown("### 📋 Current GV State")
        gv_table = pd.DataFrame([
            {"GlobalVariable": k, "Value": v, "Description": d} for k,v,d in [
                (GV_CMD,          str(int(gv_get(GV_CMD))),              "Pending command (0=none 1=pause 2=resume 3=close)"),
                (GV_STATUS,       str(int(gv_get(GV_STATUS))),           "EA status (0=run 1=paused)"),
                (GV_LOT,          f"{gv_get(GV_LOT):.2f}",              "Lot override (0=use input)"),
                (GV_TP,           f"{gv_get(GV_TP):.2f}",               "Global TP override"),
                (GV_DD,           f"{gv_get(GV_DD):.2f}",               "Max DD% override"),
                (GV_POSITIONS,    str(int(gv_get(GV_POSITIONS))),        "Open positions"),
                (GV_EQUITY,       f"${gv_get(GV_EQUITY):.2f}",          "Account equity"),
                (GV_BALANCE,      f"${gv_get(GV_BALANCE):.2f}",         "Account balance"),
                (GV_BAYESPROB,    f"{gv_get(GV_BAYESPROB):.4f}",        "Bayesian prob (active context)"),
                (GV_MTFSENTIMENT, str(int(gv_get(GV_MTFSENTIMENT))),    "MTF sentiment (-2 to 2)"),
                (GV_HTFBIAS,      str(int(gv_get(GV_HTFBIAS))),         "HTF bias (-1=Bear 0=Neut 1=Bull)"),
                (GV_CTX_THR,      f"{gv_get(GV_CTX_THR):.4f}",         "dynThreshold active context"),
                (GV_CTX_EXITMULT, f"{gv_get(GV_CTX_EXITMULT):.3f}",    "dynExitMult active context"),
                (GV_CTX_STREAK,   str(int(gv_get(GV_CTX_STREAK))),      "Loss streak active context"),
                (GV_CTX_PAUSED,   str(int(gv_get(GV_CTX_PAUSED))),      "Context paused (1=yes)"),
                (GV_HEARTBEAT,    f"{hb_age:.0f}s ago",                 "Last EA heartbeat"),
            ]
        ])
        st.dataframe(gv_table, width='stretch', hide_index=True)
        st.caption("💡 EA reads GV_CMD on every tick. After executing, GV_CMD is reset to 0 automatically.")