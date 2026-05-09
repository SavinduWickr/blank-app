import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
import warnings
import re
import io

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

# ─── Page Config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Gold Trade Analyzer",
    layout="wide",
    page_icon="🥇",
    initial_sidebar_state="expanded",
)

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Syne:wght@400;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Syne', sans-serif;
}

.stApp {
    background: #0d0d0d;
    color: #f0ead6;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: #111111;
    border-right: 1px solid #2a2a2a;
}

/* Metric cards */
[data-testid="stMetric"] {
    background: #161616;
    border: 1px solid #2a2a2a;
    border-radius: 12px;
    padding: 18px 20px;
}
[data-testid="stMetricLabel"] { color: #888; font-size: 0.78rem; letter-spacing: 0.08em; text-transform: uppercase; }
[data-testid="stMetricValue"] { color: #f0ead6; font-family: 'DM Mono', monospace; font-size: 1.6rem; }
[data-testid="stMetricDelta"] { font-family: 'DM Mono', monospace; font-size: 0.85rem; }

/* Headers */
h1 { color: #d4a847 !important; font-weight: 800; letter-spacing: -0.02em; }
h2, h3 { color: #f0ead6 !important; font-weight: 600; }

/* Divider */
hr { border-color: #2a2a2a; }

/* File uploader */
[data-testid="stFileUploader"] {
    border: 1px dashed #2a2a2a;
    border-radius: 12px;
    padding: 10px;
}

/* Info / success / error boxes */
.stAlert { border-radius: 10px; }

/* DataFrames */
[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

/* Plotly charts background */
.js-plotly-plot .plotly .bg { fill: #161616 !important; }

/* Tabs */
[data-testid="stTabs"] button {
    font-family: 'Syne', sans-serif;
    font-weight: 600;
    color: #888;
}
[data-testid="stTabs"] button[aria-selected="true"] {
    color: #d4a847;
    border-bottom-color: #d4a847;
}
</style>
""", unsafe_allow_html=True)

GOLD_SYMBOLS = ["XAUUSD", "GOLD", "XAU/USD", "XAUUSD.i", "XAUUSD+", "XAUUSDm",
                "XAUUSD.", "GOLDm", "XAUUSD_i", "XAUm"]

PLOTLY_DARK = dict(
    paper_bgcolor="#161616",
    plot_bgcolor="#161616",
    font=dict(color="#f0ead6", family="DM Mono"),
    # gridcolor="#222222",  <-- REMOVE THIS LINE
)

# ─── XML Parser ───────────────────────────────────────────────────────────────
def _find_tags(soup, tag_name: str):
    """Case-insensitive tag search — handles both lxml-xml (preserves case) and html.parser (lowercases)."""
    return soup.find_all(lambda t: t.name and t.name.lower() == tag_name.lower())


def parse_mt5_xml(file_bytes: bytes) -> pd.DataFrame | None:
    """
    Robust parser for MetaTrader 5 'Open XML (MS Office Excel)' reports.
    MT5 exports SpreadsheetML — XML with Workbook/Worksheet/Table/Row/Cell/Data tags.
    
    KEY INSIGHT: lxml-xml preserves tag case (Row, Cell, Data) but find_all('row') 
    is case-sensitive. We use a lambda for case-insensitive search across all parsers.
    """
    errors = []

    # ── Strategy 1: lxml-xml (best for well-formed XML) ───────────────────────
    try:
        soup = BeautifulSoup(file_bytes, "lxml-xml")
        df = _extract_deals_from_spreadsheetml(soup)
        if df is not None and not df.empty:
            return df
        errors.append("Strategy 1 (lxml-xml): Parsed but no deals found.")
    except Exception as e:
        errors.append(f"Strategy 1 (lxml-xml): {e}")

    # ── Strategy 2: lxml HTML parser (handles malformed XML) ──────────────────
    try:
        soup = BeautifulSoup(file_bytes, "lxml")
        df = _extract_deals_from_spreadsheetml(soup)
        if df is not None and not df.empty:
            return df
        errors.append("Strategy 2 (lxml): Parsed but no deals found.")
    except Exception as e:
        errors.append(f"Strategy 2 (lxml): {e}")

    # ── Strategy 3: html.parser (pure Python, no C dependency) ────────────────
    try:
        soup = BeautifulSoup(file_bytes, "html.parser")
        df = _extract_deals_from_spreadsheetml(soup)
        if df is not None and not df.empty:
            return df
        errors.append("Strategy 3 (html.parser): Parsed but no deals found.")
    except Exception as e:
        errors.append(f"Strategy 3 (html.parser): {e}")

    # ── Strategy 4: pandas read_html (for HTML-style reports) ─────────────────
    try:
        text = file_bytes.decode("utf-8", errors="replace")
        tables = pd.read_html(io.StringIO(text))
        for tbl in tables:
            df = _try_align_columns(tbl)
            if df is not None and not df.empty:
                return df
        errors.append("Strategy 4 (read_html): No usable table found.")
    except Exception as e:
        errors.append(f"Strategy 4 (read_html): {e}")

    st.error("❌ Could not parse the XML file. Attempted 4 parsing strategies.")
    with st.expander("Show debug details"):
        for e in errors:
            st.caption(e)
        try:
            preview = file_bytes[:800].decode("utf-8", errors="replace")
            st.code(preview, language="xml")
        except Exception:
            pass
    st.info("💡 Export from MT5: History tab → Right-click → Report → **Open XML (MS Office Excel)**")
    return None


def _extract_deals_from_spreadsheetml(soup) -> pd.DataFrame | None:
    """
    SpreadsheetML structure:
        Workbook > Worksheet > Table > Row > Cell > Data
    
    The file has a header section (account info) before the actual trade table.
    We find the Row that contains trade column headers (Symbol, Profit, Time…)
    and treat everything after it as data rows.
    """
    rows = _find_tags(soup, "row")
    if not rows:
        return None

    all_rows = []
    for row in rows:
        # Each Row > Cell > Data. We want the text of each Data element.
        cell_texts = []
        for cell in _find_tags(row, "cell"):
            data = _find_tags(cell, "data")
            cell_texts.append(data[0].get_text(strip=True) if data else "")
        if any(t != "" for t in cell_texts):   # skip fully-empty rows
            all_rows.append(cell_texts)

    if not all_rows:
        return None

    # Locate the header row — needs ≥3 recognised column names
    HEADER_KEYWORDS = {
        "symbol", "profit", "time", "type", "volume", "price",
        "commission", "swap", "comment", "order", "deal", "direction", "entry",
    }
    header_idx = None
    for i, row in enumerate(all_rows):
        row_lower = {c.lower().strip() for c in row}
        if len(row_lower & HEADER_KEYWORDS) >= 3:
            header_idx = i
            break

    if header_idx is None:
        return None

    headers = [c.strip() for c in all_rows[header_idx]]
    data_rows = all_rows[header_idx + 1:]

    n = len(headers)
    cleaned = [(row + [""] * n)[:n] for row in data_rows]

    df = pd.DataFrame(cleaned, columns=headers)
    return _try_align_columns(df)


def _try_align_columns(df: pd.DataFrame) -> pd.DataFrame | None:
    """
    Normalise column names, parse types, drop non-deal rows.
    Returns None if the DataFrame doesn't look like trade data.
    """
    # Normalise column names
    df.columns = [str(c).strip().title() for c in df.columns]

    # Map common alternative column names
    rename_map = {
        "Deal": "Deal",
        "Order": "Order",
        "Time": "Time",
        "Type": "Type",
        "Direction": "Direction",
        "Volume": "Volume",
        "Symbol": "Symbol",
        "Price": "Price",
        "Commission": "Commission",
        "Swap": "Swap",
        "Profit": "Profit",
        "Balance": "Balance",
        "Comment": "Comment",
        "Entry": "Entry",
        # Variations
        "S/L": "SL",
        "T/P": "TP",
        "Sl": "SL",
        "Tp": "TP",
        "Net Profit": "Profit",
    }
    df.rename(columns=rename_map, inplace=True)

    # Need at minimum: Time and Profit (or Symbol)
    required = {"Time", "Profit"}
    if not required.issubset(df.columns):
        return None

    # Drop rows where Time or Profit is empty/NaN
    df = df[df["Time"].astype(str).str.strip() != ""]
    df = df[df["Profit"].astype(str).str.strip() != ""]

    # Parse Profit
    df["Profit"] = pd.to_numeric(
        df["Profit"].astype(str).str.replace(r"[^\d.\-]", "", regex=True),
        errors="coerce"
    )
    df = df.dropna(subset=["Profit"])

    # Parse Time — MT5 uses "YYYY.MM.DD HH:MM:SS" format
    df["Time"] = df["Time"].astype(str).str.strip()
    df["Time"] = pd.to_datetime(
        df["Time"],
        format="%Y.%m.%d %H:%M:%S",
        errors="coerce"
    )
    # Fallback: let pandas infer
    mask = df["Time"].isna()
    if mask.any():
        df.loc[mask, "Time"] = pd.to_datetime(
            df.loc[mask, "Time_orig"] if "Time_orig" in df.columns else df.loc[mask, "Time"],
            errors="coerce",
            infer_datetime_format=True,
        )

    df = df.dropna(subset=["Time"])

    # Parse Volume
    if "Volume" in df.columns:
        df["Volume"] = pd.to_numeric(
            df["Volume"].astype(str).str.replace(r"[^\d.]", "", regex=True),
            errors="coerce"
        )

    # Filter: keep only "out" (closing) deals — they carry the realised P&L.
    # MT5 marks closing deals as type "out" or direction "out"
    if "Direction" in df.columns:
        dirs = df["Direction"].astype(str).str.lower()
        closing = df[dirs.str.contains("out|close", na=False)]
        if not closing.empty:
            df = closing

    # Remove balance/credit/deposit rows (profit usually ≠ 0 for real trades,
    # but balance entries have a Type of "balance" or very large round numbers)
    if "Type" in df.columns:
        bad_types = df["Type"].astype(str).str.lower().str.contains(
            "balance|credit|deposit|withdraw|correction", na=False
        )
        df = df[~bad_types]

    if df.empty:
        return None

    return df.reset_index(drop=True)


# ─── Helpers ──────────────────────────────────────────────────────────────────
def detect_gold(df: pd.DataFrame) -> pd.DataFrame:
    if "Symbol" not in df.columns:
        return df
    gold_mask = df["Symbol"].astype(str).str.upper().str.replace(
        r"[\s\._\-]", "", regex=True
    ).str.contains("|".join(["XAUUSD", "GOLD", "XAU"]), na=False)
    return df[gold_mask].reset_index(drop=True)


def apply_timezone(df: pd.DataFrame, offset_hours: int) -> pd.DataFrame:
    df = df.copy()
    df["Time"] = df["Time"] + pd.to_timedelta(offset_hours, unit="h")
    return df


def filter_session(df: pd.DataFrame, start_h: int, end_h: int) -> pd.DataFrame:
    mask = (df["Time"].dt.hour >= start_h) & (df["Time"].dt.hour < end_h)
    return df[mask].reset_index(drop=True)


def make_chart(fig, height=340):
    fig.update_layout(
        **PLOTLY_DARK,
        height=height,
        margin=dict(l=10, r=10, t=30, b=10),
        showlegend=False,
    )
    fig.update_xaxes(gridcolor="#222", linecolor="#333", showline=True)
    fig.update_yaxes(gridcolor="#222", linecolor="#333", showline=True)
    return fig


GOLD = "#d4a847"
RED  = "#e05252"
GREEN = "#4caf7d"

DAY_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

# ─── App ──────────────────────────────────────────────────────────────────────
st.markdown("## 🥇 Gold Trade Analyzer")
st.caption("Upload your MT5 XML report · Eightcap · XAUUSD")

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Settings")
    st.markdown("**Timezone**")
    tz_offset = st.slider(
        "Server → Local offset (hours)",
        min_value=-12, max_value=14, value=3,
        help=(
            "Eightcap MT5 server runs on EET (UTC+2 or UTC+3 depending on DST). "
            "Melbourne (AEST) is UTC+10, so set +8 in winter, +7 in summer. "
            "Default +3 adjusts EET summer → Melbourne AEST roughly."
        )
    )
    st.caption(f"Server time + {tz_offset}h = your local time")

    st.markdown("---")
    st.markdown("**Session Filter**")
    session_filter = st.checkbox("Only show 8AM–5PM trades", value=False)
    session_start = 8
    session_end = 17

    st.markdown("---")
    st.markdown("**Gold Symbol**")
    st.caption("Auto-detected. Common variants: XAUUSD, GOLD, XAUUSDm")
    show_all_symbols = st.checkbox("Show all symbols (not just Gold)", value=False)

    st.markdown("---")
    st.markdown("**Debug**")
    show_raw = st.checkbox("Show raw parsed data", value=False)

# ── File Upload ───────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload MT5 XML Report",
    type=["xml", "htm", "html"],
    help="Export from MT5: Toolbox → History tab → Right-click → Report → Open XML (MS Office Excel)"
)

if not uploaded_file:
    st.markdown("""
    <div style="background:#161616; border:1px solid #2a2a2a; border-radius:14px; padding:32px 36px; margin-top:16px;">
        <h4 style="color:#d4a847; margin-top:0">How to export from MetaTrader 5</h4>
        <ol style="color:#aaa; line-height:2.2; font-size:0.95rem;">
            <li>Open MetaTrader 5</li>
            <li>Press <code style="background:#222; padding:2px 6px; border-radius:4px">Ctrl + T</code> to open the Toolbox</li>
            <li>Click the <strong style="color:#f0ead6">History</strong> tab</li>
            <li>Right-click anywhere in the list → <strong style="color:#f0ead6">Report</strong></li>
            <li>Choose <strong style="color:#d4a847">Open XML (MS Office Excel)</strong></li>
            <li>Save the file and upload it above</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── Parse ─────────────────────────────────────────────────────────────────────
file_bytes = uploaded_file.read()
with st.spinner("Parsing MT5 report..."):
    df_raw = parse_mt5_xml(file_bytes)

if df_raw is None or df_raw.empty:
    st.stop()

if show_raw:
    st.subheader("Raw Parsed Data")
    st.dataframe(df_raw.head(30))
    st.caption(f"Columns: {list(df_raw.columns)}")

# ── Apply Timezone ─────────────────────────────────────────────────────────────
df = apply_timezone(df_raw, tz_offset)

# ── Gold Filter ───────────────────────────────────────────────────────────────
if not show_all_symbols:
    df_gold = detect_gold(df)
    if df_gold.empty:
        st.warning(
            f"⚠️ No Gold (XAUUSD) trades found after filtering. "
            f"Symbols in file: {df['Symbol'].unique().tolist() if 'Symbol' in df.columns else 'unknown'}. "
            f"Enable 'Show all symbols' in the sidebar to view everything."
        )
        st.stop()
    df = df_gold
    gold_label = df["Symbol"].iloc[0] if "Symbol" in df.columns else "XAUUSD"
    st.success(f"✅ {len(df)} Gold trades found ({gold_label})")
else:
    st.info(f"Showing all {len(df)} trades across all symbols.")

# ── Session Filter ─────────────────────────────────────────────────────────────
df_all_hours = df.copy()
if session_filter:
    df = filter_session(df, session_start, session_end)
    if df.empty:
        st.warning("⚠️ No trades found within 8AM–5PM session. Try disabling the session filter or adjusting the timezone offset.")
        st.stop()

df = df.sort_values("Time").reset_index(drop=True)
df["CumProfit"] = df["Profit"].cumsum()
df["Hour"] = df["Time"].dt.hour
df["Day"] = df["Time"].dt.day_name()
df["Month"] = df["Time"].dt.to_period("M").astype(str)
df["Week"] = df["Time"].dt.isocalendar().week.astype(str)
df["Date"] = df["Time"].dt.date
df["IsWin"] = df["Profit"] > 0

# ─── Summary Metrics ──────────────────────────────────────────────────────────
st.markdown("---")
total_profit  = df["Profit"].sum()
wins          = df[df["Profit"] > 0]
losses        = df[df["Profit"] < 0]
win_rate      = len(wins) / len(df) * 100 if len(df) else 0
avg_win       = wins["Profit"].mean() if len(wins) else 0
avg_loss      = losses["Profit"].mean() if len(losses) else 0
profit_factor = (wins["Profit"].sum() / abs(losses["Profit"].sum())
                 if len(losses) and losses["Profit"].sum() != 0 else float("inf"))
expectancy    = (win_rate / 100 * avg_win) + ((1 - win_rate / 100) * avg_loss)
max_dd        = (df["CumProfit"] - df["CumProfit"].cummax()).min()

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Net Profit", f"${total_profit:,.2f}",
          delta="Realized P&L")
c2.metric("Win Rate", f"{win_rate:.1f}%",
          delta=f"{len(wins)}W / {len(losses)}L")
c3.metric("Profit Factor", f"{profit_factor:.2f}",
          delta="≥ 1.5 is good")
c4.metric("Expectancy / Trade", f"${expectancy:,.2f}")
c5.metric("Max Drawdown", f"${max_dd:,.2f}",
          delta_color="inverse", delta="Peak-to-trough")
c6.metric("Total Trades", f"{len(df)}",
          delta=f"{'Session filtered' if session_filter else 'All hours'}")

# ─── Tabs ─────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📈 Equity Curve",
    "🕒 Seasonality",
    "🗓️ Day & Hour",
    "📊 Trade Stats",
    "📋 Trade Log",
])

# ── Tab 1: Equity Curve ───────────────────────────────────────────────────────
with tab1:
    st.subheader("Equity Curve (Cumulative Realized P&L)")

    fig_eq = go.Figure()
    fig_eq.add_trace(go.Scatter(
        x=df["Time"], y=df["CumProfit"],
        mode="lines",
        line=dict(color=GOLD, width=2),
        fill="tozeroy",
        fillcolor="rgba(212,168,71,0.08)",
        name="Cumulative P&L",
        hovertemplate="%{x|%Y-%m-%d %H:%M}<br>$%{y:,.2f}<extra></extra>",
    ))
    fig_eq = make_chart(fig_eq, height=380)
    fig_eq.update_layout(yaxis_title="Profit ($)", xaxis_title="")
    st.plotly_chart(fig_eq, use_container_width=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("Daily P&L")
        daily = df.groupby("Date")["Profit"].sum().reset_index()
        daily["Color"] = daily["Profit"].apply(lambda x: GREEN if x >= 0 else RED)
        fig_daily = go.Figure(go.Bar(
            x=daily["Date"], y=daily["Profit"],
            marker_color=daily["Color"],
            hovertemplate="%{x}<br>$%{y:,.2f}<extra></extra>",
        ))
        fig_daily = make_chart(fig_daily, height=300)
        st.plotly_chart(fig_daily, use_container_width=True)

    with col_b:
        st.subheader("Monthly P&L")
        monthly = df.groupby("Month")["Profit"].sum().reset_index()
        monthly["Color"] = monthly["Profit"].apply(lambda x: GREEN if x >= 0 else RED)
        fig_month = go.Figure(go.Bar(
            x=monthly["Month"], y=monthly["Profit"],
            marker_color=monthly["Color"],
            hovertemplate="%{x}<br>$%{y:,.2f}<extra></extra>",
        ))
        fig_month = make_chart(fig_month, height=300)
        st.plotly_chart(fig_month, use_container_width=True)

# ── Tab 2: Seasonality ────────────────────────────────────────────────────────
with tab2:
    st.subheader("Seasonality Heatmap — Hour of Day vs Day of Week")
    st.caption("Each cell = total profit for that day/hour combination. Green = profitable zone, Red = avoid.")

    available_days = [d for d in DAY_ORDER if d in df["Day"].unique()]
    pivot = df.pivot_table(
        index="Day", columns="Hour", values="Profit", aggfunc="sum"
    ).reindex(available_days)

    fig_heat = px.imshow(
        pivot,
        color_continuous_scale=[[0, RED], [0.5, "#1a1a1a"], [1, GREEN]],
        color_continuous_midpoint=0,
        aspect="auto",
        text_auto=".0f",
        labels=dict(x="Hour (Local Time)", y="", color="P&L ($)"),
    )
    fig_heat.update_layout(
        **PLOTLY_DARK,
        height=320,
        margin=dict(l=10, r=10, t=10, b=10),
        coloraxis_colorbar=dict(
            tickfont=dict(color="#aaa"),
            title=dict(text="P&L", font=dict(color="#aaa")),
        ),
    )
    fig_heat.update_traces(textfont=dict(size=10, color="#f0ead6"))
    st.plotly_chart(fig_heat, use_container_width=True)

    # Win rate heatmap
    st.subheader("Win Rate Heatmap — Hour of Day vs Day of Week")
    pivot_wr = df.pivot_table(
        index="Day", columns="Hour", values="IsWin", aggfunc="mean"
    ).reindex(available_days) * 100

    fig_wr = px.imshow(
        pivot_wr,
        color_continuous_scale=[[0, RED], [0.5, "#1a1a1a"], [1, GREEN]],
        range_color=[0, 100],
        color_continuous_midpoint=50,
        aspect="auto",
        text_auto=".0f",
        labels=dict(x="Hour (Local Time)", y="", color="Win %"),
    )
    fig_wr.update_layout(
        **PLOTLY_DARK,
        height=320,
        margin=dict(l=10, r=10, t=10, b=10),
        coloraxis_colorbar=dict(
            tickfont=dict(color="#aaa"),
            title=dict(text="Win%", font=dict(color="#aaa")),
        ),
    )
    fig_wr.update_traces(textfont=dict(size=10, color="#f0ead6"))
    st.plotly_chart(fig_wr, use_container_width=True)

# ── Tab 3: Day & Hour Charts ──────────────────────────────────────────────────
with tab3:
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("P&L by Day of Week")
        day_pnl = df.groupby("Day")["Profit"].sum().reindex(DAY_ORDER).dropna()
        colors = [GREEN if v >= 0 else RED for v in day_pnl]
        fig_day = go.Figure(go.Bar(
            x=day_pnl.index, y=day_pnl.values,
            marker_color=colors,
            text=[f"${v:,.0f}" for v in day_pnl.values],
            textposition="outside",
            textfont=dict(color="#f0ead6", size=11),
            hovertemplate="%{x}<br>$%{y:,.2f}<extra></extra>",
        ))
        fig_day = make_chart(fig_day, height=320)
        st.plotly_chart(fig_day, use_container_width=True)

        st.subheader("Win Rate by Day of Week")
        day_wr = df.groupby("Day")["IsWin"].mean().reindex(DAY_ORDER).dropna() * 100
        fig_wr_day = go.Figure(go.Bar(
            x=day_wr.index, y=day_wr.values,
            marker_color=[GREEN if v >= 50 else RED for v in day_wr.values],
            text=[f"{v:.0f}%" for v in day_wr.values],
            textposition="outside",
            textfont=dict(color="#f0ead6", size=11),
        ))
        fig_wr_day = make_chart(fig_wr_day, height=300)
        fig_wr_day.add_hline(y=50, line_dash="dash", line_color="#555", annotation_text="50%")
        st.plotly_chart(fig_wr_day, use_container_width=True)

    with col2:
        st.subheader("P&L by Hour (Local Time)")
        hour_pnl = df.groupby("Hour")["Profit"].sum()
        colors_h = [GREEN if v >= 0 else RED for v in hour_pnl.values]
        fig_hour = go.Figure(go.Bar(
            x=hour_pnl.index, y=hour_pnl.values,
            marker_color=colors_h,
            text=[f"${v:,.0f}" for v in hour_pnl.values],
            textposition="outside",
            textfont=dict(color="#f0ead6", size=10),
            hovertemplate="Hour %{x}:00<br>$%{y:,.2f}<extra></extra>",
        ))
        fig_hour = make_chart(fig_hour, height=320)
        fig_hour.update_layout(xaxis=dict(dtick=1))
        if session_filter:
            fig_hour.add_vrect(x0=session_start - 0.5, x1=session_end - 0.5,
                               fillcolor=GOLD, opacity=0.04,
                               annotation_text="8AM–5PM", annotation_position="top left",
                               annotation_font_color=GOLD)
        st.plotly_chart(fig_hour, use_container_width=True)

        st.subheader("Win Rate by Hour")
        hour_wr = df.groupby("Hour")["IsWin"].mean() * 100
        fig_wr_h = go.Figure(go.Bar(
            x=hour_wr.index, y=hour_wr.values,
            marker_color=[GREEN if v >= 50 else RED for v in hour_wr.values],
            text=[f"{v:.0f}%" for v in hour_wr.values],
            textposition="outside",
            textfont=dict(color="#f0ead6", size=10),
        ))
        fig_wr_h = make_chart(fig_wr_h, height=300)
        fig_wr_h.add_hline(y=50, line_dash="dash", line_color="#555", annotation_text="50%")
        fig_wr_h.update_layout(xaxis=dict(dtick=1))
        st.plotly_chart(fig_wr_h, use_container_width=True)

# ── Tab 4: Trade Stats ────────────────────────────────────────────────────────
with tab4:
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Profit Distribution")
        fig_hist = px.histogram(
            df, x="Profit", nbins=40,
            color_discrete_sequence=[GOLD],
            labels={"Profit": "Trade Profit ($)"},
        )
        fig_hist.add_vline(x=0, line_dash="dash", line_color="#555")
        fig_hist.add_vline(x=avg_win, line_dash="dot", line_color=GREEN,
                           annotation_text=f"Avg Win ${avg_win:,.0f}", annotation_font_color=GREEN)
        fig_hist.add_vline(x=avg_loss, line_dash="dot", line_color=RED,
                           annotation_text=f"Avg Loss ${avg_loss:,.0f}", annotation_font_color=RED)
        fig_hist = make_chart(fig_hist, height=320)
        st.plotly_chart(fig_hist, use_container_width=True)

    with col2:
        st.subheader("Win vs Loss — Average Size")
        fig_wl = go.Figure(go.Bar(
            x=["Avg Win", "Avg Loss"],
            y=[avg_win, abs(avg_loss)],
            marker_color=[GREEN, RED],
            text=[f"${avg_win:,.2f}", f"${abs(avg_loss):,.2f}"],
            textposition="outside",
            textfont=dict(color="#f0ead6", size=13),
        ))
        fig_wl = make_chart(fig_wl, height=320)
        st.plotly_chart(fig_wl, use_container_width=True)

    # Consecutive wins/losses
    st.subheader("Consecutive Wins & Losses (Streak Analysis)")
    streaks = []
    current_type = None
    count = 0
    for is_win in df["IsWin"]:
        if is_win == current_type:
            count += 1
        else:
            if current_type is not None:
                streaks.append({"Type": "Win" if current_type else "Loss", "Length": count})
            current_type = is_win
            count = 1
    if current_type is not None:
        streaks.append({"Type": "Win" if current_type else "Loss", "Length": count})

    if streaks:
        streak_df = pd.DataFrame(streaks)
        max_win_streak  = streak_df[streak_df["Type"] == "Win"]["Length"].max() if "Win" in streak_df["Type"].values else 0
        max_loss_streak = streak_df[streak_df["Type"] == "Loss"]["Length"].max() if "Loss" in streak_df["Type"].values else 0
        s1, s2 = st.columns(2)
        s1.metric("Max Win Streak", f"{max_win_streak} trades in a row")
        s2.metric("Max Loss Streak", f"{max_loss_streak} trades in a row",
                  delta_color="inverse", delta="Watch your psychology here")

    # Volume analysis if available
    if "Volume" in df.columns and df["Volume"].notna().any():
        st.subheader("Lot Size vs Profitability")
        fig_vol = px.scatter(
            df, x="Volume", y="Profit",
            color="IsWin",
            color_discrete_map={True: GREEN, False: RED},
            labels={"Volume": "Lot Size", "Profit": "Profit ($)", "IsWin": "Win"},
            hover_data={"Time": True},
        )
        fig_vol.add_hline(y=0, line_dash="dash", line_color="#555")
        fig_vol = make_chart(fig_vol, height=320)
        fig_vol.update_layout(showlegend=True)
        st.plotly_chart(fig_vol, use_container_width=True)

# ── Tab 5: Trade Log ──────────────────────────────────────────────────────────
with tab5:
    st.subheader("Full Trade Log")

    display_cols = [c for c in ["Time", "Symbol", "Type", "Direction", "Volume",
                                "Price", "Profit", "Commission", "Swap", "Comment"]
                   if c in df.columns]

    log_df = df[display_cols].copy()
    log_df["Profit"] = log_df["Profit"].map(lambda x: f"${x:,.2f}")

    st.dataframe(
        log_df.style.apply(
            lambda row: [
                f"color: {GREEN}" if "$" in str(v) and float(str(v).replace("$", "").replace(",", "")) > 0
                else f"color: {RED}" if "$" in str(v) and float(str(v).replace("$", "").replace(",", "")) < 0
                else ""
                for v in row
            ],
            axis=1
        ),
        use_container_width=True,
        height=500,
    )

    # Download button
    csv = df.to_csv(index=False)
    st.download_button(
        label="⬇️ Download Cleaned CSV",
        data=csv,
        file_name="gold_trades_cleaned.csv",
        mime="text/csv",
    )
