# 🥇 MT5 Gold Trade Analyzer

A Streamlit dashboard for analyzing MetaTrader 5 trade history — built specifically for **XAUUSD (Gold)** trades on **Eightcap**, session hours **8AM–5PM**.

---

## Deploy to Streamlit Cloud (Free)

1. Create a GitHub account if you don't have one
2. Create a **new repository** called `mt5-gold-analyzer`
3. Upload `app.py` and `requirements.txt` to the repo
4. Go to **[share.streamlit.io](https://share.streamlit.io)**
5. Click **New app** → select your repo → set `app.py` as the main file → Deploy
6. Share the URL with her ✅

---

## How She Uses It

1. Open **MetaTrader 5**
2. Press `Ctrl + T` → go to **History** tab
3. Right-click → **Report** → **Open XML (MS Office Excel)**
4. Save the `.xml` file
5. Go to the app URL → Upload the file → Done

---

## Timezone Setting

Eightcap MT5 server runs on **EET (Eastern European Time)**:
- **UTC+2** in winter (Nov–Mar)
- **UTC+3** in summer (Mar–Oct)

Melbourne (AEST) is **UTC+10**, so:
- Winter: set slider to **+8**
- Summer: set slider to **+7**

The default is **+3** which is a middle ground — she should adjust based on whether her trade times look right.

---

## Features

- ✅ Robust XML parser (3 fallback strategies)
- ✅ Auto-detects all Gold symbol variants (XAUUSD, XAUUSDm, GOLD, etc.)
- ✅ 8AM–5PM session filter (toggleable)
- ✅ Equity curve, daily/monthly P&L bars
- ✅ Seasonality heatmap (profit AND win rate by hour/day)
- ✅ Day of week & hour of day analysis
- ✅ Streak analysis (max win/loss streak)
- ✅ Lot size vs profitability scatter
- ✅ Full trade log with CSV download
- ✅ Key metrics: Profit Factor, Expectancy, Max Drawdown, Win Rate
