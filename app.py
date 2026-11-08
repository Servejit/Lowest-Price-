import io
import re
import time
from datetime import datetime

import pandas as pd
import streamlit as st
import yfinance as yf

st.set_page_config(page_title="Lowest Price Stock Scanner", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .main {background: #f7f9fc;}
    .block-container {padding-top: 1.4rem; padding-bottom: 2rem;}
    [data-testid="stMetric"] {background: white; border: 1px solid #e7ebf2; padding: 14px; border-radius: 12px;}
    div[data-testid="stDataFrame"] {border: 1px solid #e7ebf2; border-radius: 10px;}
</style>
""", unsafe_allow_html=True)

st.title("📊 Lowest Price Stock Scanner")
st.caption("Switch between the 20-session low setup and its opposite 20-session high setup.")

with st.sidebar:
    st.header("Scanner settings")
    near_pct = st.slider("Near-low distance (%)", 0.0, 10.0, 1.0, 0.25,
                         help="A stock is near its 20-day low when its LTP is no more than this percentage above that low.")
    only_matches = st.checkbox("Show only stocks meeting all conditions", value=False)
    st.caption("Market quotes are obtained from Yahoo Finance via yfinance. They may be delayed, unavailable, or rate-limited.")

uploaded = st.file_uploader("Upload today's Excel workbook (.xlsx)", type=["xlsx", "xls"])

def numeric_from_bracket(value):
    if pd.isna(value):
        return None
    text = str(value)
    matches = re.findall(r"\(([-+]?\d+(?:\.\d+)?)\)", text)
    if matches:
        return float(matches[-1])
    try:
        return float(value)
    except (TypeError, ValueError):
        return None

def read_workbook(file_obj):
    book = pd.ExcelFile(file_obj)
    records = []
    for sheet in book.sheet_names:
        frame = pd.read_excel(book, sheet_name=sheet)
        if frame.empty:
            continue
        symbol_col = next((c for c in frame.columns if str(c).strip().lower().startswith("symbol")), None)
        if symbol_col is None:
            continue
        low_avg_col = next((c for c in frame.columns if "16<" in str(c) and "avg.4" in str(c).lower()), None)
        high_avg_col = next((c for c in frame.columns if "16>" in str(c) and "avg.4" in str(c).lower()), None)
        for _, row in frame.iterrows():
            symbol = str(row.get(symbol_col, "")).strip().upper()
            if not symbol or symbol in {"NAN", "NONE"}:
                continue
            records.append({
                "Symbol": symbol,
                "O2L Average (%)": numeric_from_bracket(row.get(low_avg_col)) if low_avg_col else None,
                "O2H Average (%)": numeric_from_bracket(row.get(high_avg_col)) if high_avg_col else None,
                "Source sheet": sheet,
            })
    if not records:
        raise ValueError("No stock Symbol column was found in the workbook.")
    data = pd.DataFrame(records)
    # Keep one row per stock; prefer the first sheet with usable average values.
    data["_score"] = data[["O2L Average (%)", "O2H Average (%)"]].notna().sum(axis=1)
    data = data.sort_values("_score", ascending=False).drop_duplicates("Symbol", keep="first")
    return data.drop(columns="_score").sort_values("Symbol").reset_index(drop=True), book.sheet_names

def get_market_data(symbols, progress=None):
    results = []
    # Download in batches to reduce requests and avoid one request per stock.
    batch_size = 35
    for start in range(0, len(symbols), batch_size):
        batch = symbols[start:start + batch_size]
        tickers = [s + ".NS" for s in batch]
        daily = yf.download(tickers=tickers, period="2mo", interval="1d",
                            group_by="ticker", auto_adjust=False, progress=False,
                            threads=True, timeout=25)
        intraday = yf.download(tickers=tickers, period="1d", interval="1m",
                               group_by="ticker", auto_adjust=False, progress=False,
                               threads=True, timeout=25)
        for symbol, ticker in zip(batch, tickers):
            try:
                if isinstance(daily.columns, pd.MultiIndex):
                    d = daily[ticker].dropna(how="all") if ticker in daily.columns.get_level_values(0) else pd.DataFrame()
                else:
                    d = daily.dropna(how="all") if len(batch) == 1 else pd.DataFrame()
                if d.empty or not {"Low", "Open", "High", "Close"}.issubset(d.columns):
                    results.append({"Symbol": symbol, "Quote status": "No daily data"})
                    continue
                last20 = d.tail(20)
                lowest20 = float(last20["Low"].min())
                highest20 = float(last20["High"].max())
                day_open = float(d["Open"].iloc[-1])
                day_low = float(d["Low"].iloc[-1])
                day_high = float(d["High"].iloc[-1])
                ltp = float(d["Close"].iloc[-1])
                quote_status = "Latest daily close (not live)"

                if isinstance(intraday.columns, pd.MultiIndex):
                    intr = intraday[ticker].dropna(how="all") if ticker in intraday.columns.get_level_values(0) else pd.DataFrame()
                else:
                    intr = intraday.dropna(how="all") if len(batch) == 1 else pd.DataFrame()
                if not intr.empty and {"Open", "Low", "High", "Close"}.issubset(intr.columns):
                    intr = intr.dropna(subset=["Close"])
                    if not intr.empty:
                        day_open = float(intr["Open"].dropna().iloc[0]) if intr["Open"].notna().any() else day_open
                        day_low = float(intr["Low"].min())
                        day_high = float(intr["High"].max())
                        ltp = float(intr["Close"].iloc[-1])
                        quote_status = "Intraday quote (may be delayed)"

                # O2L is today's move from today's Open down to today's Low.
                # O2H is today's move from today's Open up to today's High.
                o2l_live = ((day_low - day_open) / day_open * 100) if pd.notna(day_open) and day_open > 0 and pd.notna(day_low) else None
                o2h_live = ((day_high - day_open) / day_open * 100) if pd.notna(day_open) and day_open > 0 and pd.notna(day_high) else None
                distance = ((ltp - lowest20) / lowest20 * 100) if lowest20 else None
                below_high = ((highest20 - ltp) / highest20 * 100) if highest20 else None
                results.append({
                    "Symbol": symbol, "Live price": ltp, "20-day lowest Low": lowest20,
                    "Above 20-day Low (%)": distance, "20-day highest High": highest20,
                    "Below 20-day High (%)": below_high, "Today Open": day_open,
                    "Today Low": day_low, "O2L live (%)": o2l_live,
                    "Today High": day_high, "O2H live (%)": o2h_live,
                    "Quote status": quote_status,
                    "20 sessions checked": len(last20),
                })
            except Exception as exc:
                results.append({"Symbol": symbol, "Quote status": f"Data error: {str(exc)[:100]}"})
        if progress:
            progress.progress(min((start + len(batch)) / len(symbols), 1.0))
    return pd.DataFrame(results)

if uploaded is None:
    st.info("Upload the daily Excel workbook to begin. The workbook can be replaced each day.")
    st.stop()


try:
    stocks, sheets = read_workbook(uploaded)
except Exception as exc:
    st.error(f"Could not read this workbook: {exc}")
    st.stop()

st.success(f"Loaded {len(stocks):,} unique stocks from {len(sheets)} worksheets: {', '.join(sheets)}")
dashboard_mode = st.radio(
    "Select dashboard",
    ["20-day Low Dashboard", "20-day High Dashboard"],
    horizontal=True,
    key="dashboard_mode",
)
dash_left, dash_right = st.columns([3, 1])
with dash_left:
    st.caption("O2L comparison = today's Open-to-Low % versus the workbook O2L average. O2H comparison = today's Open-to-High % versus the workbook O2H average. Market OHLC and 20-session price extremes are fetched separately.")
with dash_right:
    refresh = st.button("🔄 Refresh market data", use_container_width=True, type="primary")

if st.session_state.get("workbook_name") != uploaded.name:
    st.session_state.pop("market_data", None)
    st.session_state["workbook_name"] = uploaded.name

should_fetch = (
    refresh
    or "market_data" not in st.session_state
)

if should_fetch:
    bar = st.progress(0, text="Fetching market data. This can take a few minutes for a large stock list...")
    try:
        st.session_state["market_data"] = get_market_data(stocks["Symbol"].tolist(), bar)
        bar.empty()
        st.session_state["market_fetched_at"] = datetime.now().strftime("%d %b %Y, %I:%M:%S %p")
    except Exception as exc:
        bar.empty()
        st.error(f"Market-data download failed: {exc}")
        st.stop()

result = stocks.merge(st.session_state["market_data"], on="Symbol", how="left")

# Original 20-day-low dashboard: retain the existing three-condition logic.
result["Near 20-day Low"] = result["Above 20-day Low (%)"].notna() & (result["Above 20-day Low (%)"] <= near_pct)
result["Low dashboard O2L condition"] = result["O2L live (%)"].notna() & result["O2L Average (%)"].notna() & (result["O2L live (%)"] <= result["O2L Average (%)"])
result["Low dashboard O2H condition"] = result["O2H live (%)"].notna() & result["O2H Average (%)"].notna() & (result["O2H live (%)"] <= result["O2H Average (%)"])
result["Low dashboard MATCH"] = result["Near 20-day Low"] & result["Low dashboard O2L condition"] & result["Low dashboard O2H condition"]

# Opposite 20-day-high dashboard: near the 20-session highest High,
# with O2H strictly greater than its average and O2L strictly less.
result["Near 20-day High"] = result["Below 20-day High (%)"].notna() & (result["Below 20-day High (%)"] <= near_pct)
result["High dashboard O2H condition"] = result["O2H live (%)"].notna() & result["O2H Average (%)"].notna() & (result["O2H live (%)"] > result["O2H Average (%)"])
result["High dashboard O2L condition"] = result["O2L live (%)"].notna() & result["O2L Average (%)"].notna() & (result["O2L live (%)"] < result["O2L Average (%)"])
result["High dashboard MATCH"] = result["Near 20-day High"] & result["High dashboard O2H condition"] & result["High dashboard O2L condition"]

is_high_dashboard = dashboard_mode == "20-day High Dashboard"
if is_high_dashboard:
    result["MATCH"] = result["High dashboard MATCH"]
    result["Near selected 20-day extreme"] = result["Near 20-day High"]
    match_heading = "Near 20-day high"
    dashboard_description = "20-day High setup: price is within the selected distance below the highest High of the last 20 sessions; O2H is strictly greater than its average; O2L is strictly less than its average."
else:
    result["MATCH"] = result["Low dashboard MATCH"]
    result["Near selected 20-day extreme"] = result["Near 20-day Low"]
    match_heading = "Near 20-day low"
    dashboard_description = "20-day Low setup: price is within the selected distance above the lowest Low of the last 20 sessions; O2L and O2H are each equal to or below their averages."

available = result["Live price"].notna()
c1, c2, c3, c4 = st.columns(4)
c1.metric("Stocks in workbook", f"{len(result):,}")
c2.metric("Market quotes found", f"{available.sum():,}")
c3.metric(match_heading, f"{result['Near selected 20-day extreme'].sum():,}")
c4.metric("All conditions matched", f"{result['MATCH'].sum():,}")

st.caption(dashboard_description)
st.caption(f"Last market-data fetch: {st.session_state.get('market_fetched_at', '—')}")

def fmt_price(x):
    return f"₹{x:,.2f}" if pd.notna(x) else "—"

def fmt_pct(x):
    return f"{x:+.2f}%" if pd.notna(x) else "—"

def compare(live, avg):
    if pd.isna(live) or pd.isna(avg):
        return "—"
    return f"{live:+.2f}% vs {avg:+.2f}%"

display = result.copy()
display["Price vs 20-day Low"] = display.apply(
    lambda r: f"{fmt_price(r.get('Live price'))} vs {fmt_price(r.get('20-day lowest Low'))} ({fmt_pct(r.get('Above 20-day Low (%)'))})", axis=1)
display["Price vs 20-day High"] = display.apply(
    lambda r: f"{fmt_price(r.get('Live price'))} vs {fmt_price(r.get('20-day highest High'))} ({fmt_pct(r.get('Below 20-day High (%)'))} below)", axis=1)
display["O2L comparison"] = display.apply(lambda r: compare(r.get("O2L live (%)"), r.get("O2L Average (%)")), axis=1)
display["O2H comparison"] = display.apply(lambda r: compare(r.get("O2H live (%)"), r.get("O2H Average (%)")), axis=1)

if only_matches:
    display = display[display["MATCH"]]
else:
    sort_distance = "Below 20-day High (%)" if is_high_dashboard else "Above 20-day Low (%)"
    display = display.sort_values(
        ["MATCH", sort_distance], ascending=[False, True], na_position="last"
    )

if is_high_dashboard:
    columns = ["Symbol", "Price vs 20-day High", "O2H comparison", "O2L comparison",
               "Near 20-day High", "High dashboard O2H condition", "High dashboard O2L condition", "MATCH", "Quote status"]
else:
    columns = ["Symbol", "Price vs 20-day Low", "O2L comparison", "O2H comparison",
               "Near 20-day Low", "Low dashboard O2L condition", "Low dashboard O2H condition", "MATCH", "Quote status"]
st.subheader("Stock comparison")
st.dataframe(display[[c for c in columns if c in display.columns]],
             use_container_width=True, hide_index=True, height=620)

csv = result.to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️ Download complete results (CSV)", csv,
                   file_name=f"lowest_price_scan_{datetime.now():%Y%m%d_%H%M}.csv",
                   mime="text/csv")

st.caption("O2L comparison shows (Today Low - Today Open) / Today Open x 100 versus the workbook O2L average. O2H comparison shows (Today High - Today Open) / Today Open x 100 versus the workbook O2H average. Missing market data is never treated as a match. The High dashboard uses strict greater-than / less-than comparisons exactly as requested; the Low dashboard retains its original less-than-or-equal comparisons.")
