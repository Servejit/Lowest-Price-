import io
import re
import time
import colorsys
import xml.etree.ElementTree as ET
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles.colors import COLOR_INDEX

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

uploaded_file = st.file_uploader(
    "Upload today's Excel workbook (.xlsx)",
    type=["xlsx", "xls"],
    key="daily_workbook_uploader",
)

# Keep the uploaded workbook in this browser session across Streamlit reruns.
# Compare file contents, not just filenames, so a replacement with the same
# filename is detected correctly.
if uploaded_file is not None:
    new_file_bytes = uploaded_file.getvalue()
    import hashlib
    new_file_hash = hashlib.sha256(new_file_bytes).hexdigest()

    if new_file_hash != st.session_state.get("workbook_hash"):
        st.session_state["workbook_bytes"] = new_file_bytes
        st.session_state["workbook_name"] = uploaded_file.name
        st.session_state["workbook_hash"] = new_file_hash
        st.session_state.pop("market_data", None)
        st.session_state.pop("market_fetched_at", None)

if "workbook_bytes" not in st.session_state:
    st.info("Upload the daily Excel workbook to begin. It will stay loaded during this session until you upload a replacement.")
    st.stop()

# Re-create a file-like object from the saved bytes on every rerun.
uploaded = io.BytesIO(st.session_state["workbook_bytes"])
uploaded.name = st.session_state.get("workbook_name", "uploaded_workbook.xlsx")

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

def get_theme_rgb(workbook):
    """Return the workbook theme colours in Excel theme-index order."""
    if not workbook.loaded_theme:
        return []
    root = ET.fromstring(workbook.loaded_theme)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    scheme = root.find(".//a:clrScheme", ns)
    if scheme is None:
        return []
    theme_rgb = []
    for name in ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2",
                 "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]:
        node = scheme.find(f"a:{name}", ns)
        if node is None or len(node) == 0:
            theme_rgb.append(None)
            continue
        colour_node = node[0]
        value = colour_node.attrib.get("lastClr") or colour_node.attrib.get("val")
        theme_rgb.append(value[-6:] if value else None)
    return theme_rgb


def apply_tint(rgb, tint):
    if not tint:
        return rgb
    channels = [int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    adjusted = []
    for channel in channels:
        if tint < 0:
            channel = channel * (1 + tint)
        else:
            channel = channel * (1 - tint) + tint
        adjusted.append(max(0, min(255, round(channel * 255))))
    return "".join(f"{value:02X}" for value in adjusted)


def cell_fill_rgb(cell, workbook, theme_rgb):
    fill = cell.fill
    if fill.fill_type is None:
        return None
    colour = fill.fgColor
    rgb = None
    if colour.type == "rgb" and colour.rgb:
        rgb = colour.rgb[-6:]
    elif colour.type == "indexed" and colour.indexed is not None:
        try:
            indexed = COLOR_INDEX[colour.indexed]
            rgb = indexed[-6:] if indexed else None
        except (IndexError, TypeError):
            rgb = None
    elif colour.type == "theme" and colour.theme is not None:
        if colour.theme < len(theme_rgb):
            rgb = theme_rgb[colour.theme]
    if not rgb:
        return None
    try:
        return apply_tint(rgb, colour.tint)
    except (TypeError, ValueError):
        return rgb


def classify_symbol_cell(cell, workbook, theme_rgb):
    """Classify solid blue/green fills in column A; ignore other colours."""
    rgb = cell_fill_rgb(cell, workbook, theme_rgb)
    if not rgb:
        return None
    red, green, blue = [int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    hue, saturation, value = colorsys.rgb_to_hsv(red, green, blue)
    if saturation < 0.18 or value < 0.18:
        return None
    if 0.20 <= hue <= 0.46:
        return "Green"
    if 0.50 <= hue <= 0.75:
        return "Blue"
    return None


def read_workbook(file_obj):
    file_bytes = file_obj.getvalue()
    book = pd.ExcelFile(io.BytesIO(file_bytes))
    style_book = load_workbook(io.BytesIO(file_bytes), data_only=True)
    records = []
    marked_colours = {}
    theme_rgb = get_theme_rgb(style_book)

    for sheet in book.sheet_names:
        frame = pd.read_excel(book, sheet_name=sheet)
        if frame.empty:
            continue
        symbol_col = next((c for c in frame.columns if str(c).strip().lower().startswith("symbol")), None)
        if symbol_col is None:
            continue
        low_avg_col = next((c for c in frame.columns if "16<" in str(c) and "avg.4" in str(c).lower()), None)
        high_avg_col = next((c for c in frame.columns if "16>" in str(c) and "avg.4" in str(c).lower()), None)
        worksheet = style_book[sheet]

        for row_index, row in frame.iterrows():
            symbol = str(row.get(symbol_col, "")).strip().upper()
            if not symbol or symbol in {"NAN", "NONE"}:
                continue

            # The fill colour is read ONLY from column A of this stock's row.
            excel_row = int(row_index) + 2
            cell_colour = classify_symbol_cell(
                worksheet.cell(row=excel_row, column=1), style_book, theme_rgb
            )

            # Apply sheet-specific colour rules for the Blue/Green dashboard:
            # AutoBlue contributes BLUE cells only; AutoGreen contributes
            # GREEN cells only. Other worksheets do not contribute symbols.
            sheet_key = str(sheet).strip().lower()
            if sheet_key == "autoblue" and cell_colour == "Blue":
                marked_colours.setdefault(symbol, set()).add("Blue")
            elif sheet_key == "autogreen" and cell_colour == "Green":
                marked_colours.setdefault(symbol, set()).add("Green")

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
    data = data.drop(columns="_score")
    data["Symbol cell colour"] = data["Symbol"].map(
        lambda symbol: " + ".join(sorted(marked_colours.get(symbol, set()))) or "—"
    )
    data["Blue/Green"] = data["Symbol"].map(
        lambda symbol: bool(marked_colours.get(symbol, set()))
    )
    style_book.close()
    return data.sort_values("Symbol").reset_index(drop=True), book.sheet_names

def get_market_data(symbols, progress=None):
    results = []
    batch_size = 35

    for start in range(0, len(symbols), batch_size):
        batch = symbols[start:start + batch_size]
        tickers = [s + ".NS" for s in batch]

        daily = yf.download(
            tickers=tickers, period="2mo", interval="1d",
            group_by="ticker", auto_adjust=False, progress=False,
            threads=True, timeout=25
        )
        intraday = yf.download(
            tickers=tickers, period="1d", interval="1m",
            group_by="ticker", auto_adjust=False, progress=False,
            threads=True, timeout=25
        )

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

                daily_row = d.iloc[-1]
                day_open = float(daily_row["Open"])
                day_low = float(daily_row["Low"])
                day_high = float(daily_row["High"])
                ltp = float(daily_row["Close"])
                session_date = pd.Timestamp(d.index[-1]).strftime("%d %b %Y")
                quote_status = f"Daily candle only — {session_date}; not live"

                if isinstance(intraday.columns, pd.MultiIndex):
                    intr = intraday[ticker].dropna(how="all") if ticker in intraday.columns.get_level_values(0) else pd.DataFrame()
                else:
                    intr = intraday.dropna(how="all") if len(batch) == 1 else pd.DataFrame()

                # Use actual 1-minute session candles for Open, running Low,
                # running High and LTP instead of relying on a possibly stale
                # daily candle while the market is trading.
                required = {"Open", "High", "Low", "Close"}
                if not intr.empty and required.issubset(intr.columns):
                    intr = intr.dropna(subset=["Close"]).sort_index()
                    if not intr.empty:
                        intraday_date = pd.Timestamp(intr.index[-1]).strftime("%d %b %Y")
                        first_open = intr["Open"].dropna()
                        highs = intr["High"].dropna()
                        lows = intr["Low"].dropna()

                        if not first_open.empty and not highs.empty and not lows.empty:
                            day_open = float(first_open.iloc[0])
                            day_high = float(highs.max())
                            day_low = float(lows.min())
                            ltp = float(intr["Close"].iloc[-1])
                            session_date = intraday_date
                            quote_status = f"1-minute session OHLC — {session_date}; Yahoo data may be delayed"

                o2l_live = ((day_low - day_open) / day_open * 100) if day_open > 0 else None
                o2h_live = ((day_high - day_open) / day_open * 100) if day_open > 0 else None
                distance = ((ltp - lowest20) / lowest20 * 100) if lowest20 else None
                below_high = ((highest20 - ltp) / highest20 * 100) if highest20 else None

                results.append({
                    "Symbol": symbol,
                    "Live price": ltp,
                    "20-day lowest Low": lowest20,
                    "Above 20-day Low (%)": distance,
                    "20-day highest High": highest20,
                    "Below 20-day High (%)": below_high,
                    "Session date": session_date,
                    "Today Open": day_open,
                    "Today Low": day_low,
                    "O2L live (%)": o2l_live,
                    "Today High": day_high,
                    "O2H live (%)": o2h_live,
                    "Quote status": quote_status,
                    "20 sessions checked": len(last20),
                })
            except Exception as exc:
                results.append({"Symbol": symbol, "Quote status": f"Data error: {str(exc)[:100]}"})

        if progress:
            progress.progress(min((start + len(batch)) / len(symbols), 1.0))

    return pd.DataFrame(results)

try:
    stocks, sheets = read_workbook(uploaded)
except Exception as exc:
    st.error(f"Could not read this workbook: {exc}")
    st.stop()

st.success(f"Loaded {len(stocks):,} unique stocks from {len(sheets)} worksheets: {', '.join(sheets)}")
dashboard_mode = st.radio(
    "Select dashboard",
    ["20-day Low Dashboard", "20-day High Dashboard", "Blue/Green Stocks Dashboard"],
    horizontal=True,
    key="dashboard_mode",
)
dash_left, dash_right = st.columns([3, 1])
with dash_left:
    st.caption("O2L/O2H use the session Open, running Low and running High from 1-minute candles when available; the session date and quote source are shown in the table.")
with dash_right:
    refresh = st.button("🔄 Refresh market data", use_container_width=True, type="primary")

# Fetch only the symbols needed by the selected dashboard.
# The Blue/Green dashboard downloads market data exclusively for coloured
# column-A symbols; the other dashboards continue to use the full workbook.
if dashboard_mode == "Blue/Green Stocks Dashboard":
    selected_stocks = stocks[stocks["Blue/Green"]].copy()
    market_symbols = selected_stocks["Symbol"].drop_duplicates().tolist()
else:
    selected_stocks = stocks.copy()
    market_symbols = selected_stocks["Symbol"].drop_duplicates().tolist()

if not market_symbols:
    st.warning("No blue or green filled Symbol cells were found in column A of either worksheet.")
    st.stop()

cached_symbols = st.session_state.get("market_data_symbols", [])
should_fetch = (
    refresh
    or "market_data" not in st.session_state
    or set(cached_symbols) != set(market_symbols)
)

if should_fetch:
    bar = st.progress(0, text=f"Fetching market data for {len(market_symbols):,} selected stocks...")
    try:
        st.session_state["market_data"] = get_market_data(market_symbols, bar)
        st.session_state["market_data_symbols"] = market_symbols
        bar.empty()
        st.session_state["market_fetched_at"] = datetime.now().strftime("%d %b %Y, %I:%M:%S %p")
    except Exception as exc:
        bar.empty()
        st.error(f"Market-data download failed: {exc}")
        st.stop()

# Merge only the stocks selected for this dashboard, never the full workbook
# in the Blue/Green view.
result = selected_stocks.merge(st.session_state["market_data"], on="Symbol", how="left")

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
is_blue_green_dashboard = dashboard_mode == "Blue/Green Stocks Dashboard"

if is_blue_green_dashboard:
    # This view includes only symbols whose column-A cell is blue or green
    # in at least one worksheet. It shows both existing dashboard analyses.
    view_result = result[result["Blue/Green"]].copy()
    view_result["MATCH"] = view_result["Low dashboard MATCH"] | view_result["High dashboard MATCH"]
    view_result["Near selected 20-day extreme"] = view_result["Near 20-day Low"] | view_result["Near 20-day High"]
    match_heading = "Near 20-day low or high"
    dashboard_description = (
        "Blue/Green dashboard: only stocks with a blue or green fill in column A "
        "of either worksheet are included. Both the original 20-day Low and "
        "20-day High conditions are calculated without changing their rules. "
        "Overall MATCH means either dashboard's full conditions match."
    )
elif is_high_dashboard:
    view_result = result.copy()
    view_result["MATCH"] = view_result["High dashboard MATCH"]
    view_result["Near selected 20-day extreme"] = view_result["Near 20-day High"]
    match_heading = "Near 20-day high"
    dashboard_description = "20-day High setup: price is within the selected distance below the highest High of the last 20 sessions; O2H is strictly greater than its average; O2L is strictly less than its average."
else:
    view_result = result.copy()
    view_result["MATCH"] = view_result["Low dashboard MATCH"]
    view_result["Near selected 20-day extreme"] = view_result["Near 20-day Low"]
    match_heading = "Near 20-day low"
    dashboard_description = "20-day Low setup: price is within the selected distance above the lowest Low of the last 20 sessions; O2L and O2H are each equal to or below their averages."

available = view_result["Live price"].notna()
c1, c2, c3, c4 = st.columns(4)
c1.metric("Stocks in this dashboard", f"{len(view_result):,}")
c2.metric("Market quotes found", f"{available.sum():,}")
c3.metric(match_heading, f"{view_result['Near selected 20-day extreme'].sum():,}")
c4.metric("All conditions matched", f"{view_result['MATCH'].sum():,}")

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

display = view_result.copy()
display["Price vs 20-day Low"] = display.apply(
    lambda r: f"{fmt_price(r.get('Live price'))} vs {fmt_price(r.get('20-day lowest Low'))} ({fmt_pct(r.get('Above 20-day Low (%)'))})", axis=1)
display["Price vs 20-day High"] = display.apply(
    lambda r: f"{fmt_price(r.get('Live price'))} vs {fmt_price(r.get('20-day highest High'))} ({fmt_pct(r.get('Below 20-day High (%)'))} below)", axis=1)
display["O2L comparison"] = display.apply(lambda r: compare(r.get("O2L live (%)"), r.get("O2L Average (%)")), axis=1)
display["O2H comparison"] = display.apply(lambda r: compare(r.get("O2H live (%)"), r.get("O2H Average (%)")), axis=1)

if only_matches:
    display = display[display["MATCH"]]
else:
    if is_blue_green_dashboard:
        display = display.sort_values(
            ["MATCH", "Low dashboard MATCH", "High dashboard MATCH"],
            ascending=[False, False, False], na_position="last"
        )
    else:
        sort_distance = "Below 20-day High (%)" if is_high_dashboard else "Above 20-day Low (%)"
        display = display.sort_values(
            ["MATCH", sort_distance], ascending=[False, True], na_position="last"
        )

if is_blue_green_dashboard:
    columns = [
        "Symbol", "Symbol cell colour", "Session date", "Today Open", "Today High", "Today Low",
        "Price vs 20-day Low", "Price vs 20-day High", "O2L comparison", "O2H comparison",
        "Near 20-day Low", "Near 20-day High",
        "Low dashboard MATCH", "High dashboard MATCH", "MATCH", "Quote status"
    ]
elif is_high_dashboard:
    columns = ["Symbol", "Session date", "Today Open", "Today High", "Today Low", "Price vs 20-day High", "O2H comparison", "O2L comparison",
               "Near 20-day High", "High dashboard O2H condition", "High dashboard O2L condition", "MATCH", "Quote status"]
else:
    columns = ["Symbol", "Session date", "Today Open", "Today High", "Today Low", "Price vs 20-day Low", "O2L comparison", "O2H comparison",
               "Near 20-day Low", "Low dashboard O2L condition", "Low dashboard O2H condition", "MATCH", "Quote status"]
st.subheader("Stock comparison")
st.dataframe(display[[c for c in columns if c in display.columns]],
             use_container_width=True, hide_index=True, height=620)

csv = view_result.to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️ Download this dashboard's results (CSV)", csv,
                   file_name=f"lowest_price_scan_{datetime.now():%Y%m%d_%H%M}.csv",
                   mime="text/csv")

st.caption("For every stock: O2L (%) = (Today Low − Today Open) / Today Open × 100. O2H (%) = (Today High − Today Open) / Today Open × 100. Open, High and Low are shown beside the calculated percentages for verification. The session date and quote source are displayed. Yahoo Finance data may be delayed.")
