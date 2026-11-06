# Lowest Price Stock Scanner

Streamlit app for the daily Excel workbook workflow.

## What it does
- Upload a replacement Excel workbook each day.
- Reads stock symbols and the bracketed averages from:
  - `16< D-B / Avg.4` → O2L average
  - `16> C-B / Avg.4` → O2H average
- Fetches up to the latest 20 trading sessions of daily OHLC data for NSE symbols using Yahoo Finance through yfinance.
- Fetches intraday 1-minute OHLC data where available and calculates:
  - LTP vs lowest Low in the latest 20 sessions
  - Today's O2L = (today's low - today's open) / today's open × 100
  - Today's O2H = (today's high - today's open) / today's open × 100
- Highlights stocks meeting all three conditions and exports CSV results.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Important data notes
- Market data comes from Yahoo Finance and may be delayed, incomplete, or rate-limited. It is not an exchange-authoritative real-time feed.
- NSE symbols are looked up with the `.NS` suffix. Symbols unavailable from Yahoo Finance will show a data status instead of fabricated values.
- The app calculates the 20-session lowest absolute Low from market history, not from the workbook's O2L percentage columns.
- The uploaded workbook is read at runtime and is not committed to GitHub.
