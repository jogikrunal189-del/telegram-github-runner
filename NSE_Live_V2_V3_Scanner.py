"""
NSE LIVE V2 + V3 SCANNER
------------------------
LIVE scanner only. This is NOT a backtest.

What it does:
1. Downloads the latest daily NSE data from Yahoo Finance.
2. Calculates the same V2 (NEW) and V3 scoring logic used in your backtest.
3. Shows the top V2 and V3 stocks for the latest available market data.
4. Saves every scan into a CSV tracker so you can compare future scans.
5. Also saves an Excel file with the latest scan and the full history.

IMPORTANT:
- This is a DAILY STOCK-SELECTION scanner, not the JOGI 5-minute entry system.
- The stock is a candidate for the NEXT trading session; it is not an automatic BUY.
- Yahoo Finance data can be delayed/incomplete while the market is open.
- The scanner retries failed Yahoo downloads up to 3 times.
- NIFTY 50 uses Yahoo symbol ^NSEI (it must NOT become ^NSEI.NS).
- For stable end-of-day results, run after the NSE session closes.

Install once:
    pip install yfinance pandas openpyxl

Run:
    python NSE_Live_V2_V3_Scanner.py
"""

import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import yfinance as yf
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


# =========================
# SETTINGS
# =========================

TOP_N = 10
MIN_PRICE = 50
MAX_PRICE = 15000
MIN_VOLUME = 300_000

DATA_LOOKBACK_DAYS = 430

CSV_FILE = "NSE_Live_V2_V3_Tracker.csv"
XLSX_FILE = "NSE_Live_V2_V3_Tracker.xlsx"

# Broad liquid universe from your existing V3 code.
UNIVERSE = [
    "TCS","INFY","HCLTECH","WIPRO","TECHM","LTIM","MPHASIS","COFORGE","PERSISTENT","OFSS",
    "HDFCBANK","ICICIBANK","SBIN","KOTAKBANK","AXISBANK","BANDHANBNK","FEDERALBNK",
    "IDFCFIRSTB","INDUSINDBK","PNB","BANKBARODA","CANARABANK",
    "BAJFINANCE","BAJAJFINSV","CHOLAFIN","MUTHOOTFIN","LICHSGFIN","PFC","RECLTD",
    "BHARTIARTL","IDEA","TATACOMM","MARUTI","TATAMOTORS","M&M","BAJAJ-AUTO",
    "HEROMOTOCO","EICHERMOT","TVSMOTOR","ASHOKLEY","TATASTEEL","JSWSTEEL","HINDALCO",
    "VEDL","COALINDIA","NMDC","SAIL","NATIONALUM","SUNPHARMA","DRREDDY","CIPLA",
    "DIVISLAB","AUROPHARMA","LUPIN","BIOCON","TORNTPHARM","RELIANCE","ONGC","BPCL",
    "IOC","GAIL","POWERGRID","NTPC","ADANIGREEN","TATAPOWER","ADANIPORTS",
    "HINDUNILVR","ITC","NESTLEIND","BRITANNIA","DABUR","MARICO","GODREJCP","COLPAL",
    "ULTRACEMCO","GRASIM","AMBUJACEM","ACC","LT","SIEMENS","ABB","BHEL","TITAN",
    "ASIANPAINT","HAVELLS","VOLTAS","DMART","PIDILITIND","ADANIENT","ADANIPOWER",
    "ZOMATO","NAUKRI","IRCTC","RVNL","IRFC","RAILTEL","DELHIVERY","NYKAA","POLICYBZR",
    "PAYTM","BERGEPAINT","PREMIERPOL","MOREPENLAB","CUPID","AXISCADES","FIRSTCRY",
    "SBC","LLOYDSENGG","MARINE","SHUKRAPHAR","MANALIPETC"
]
UNIVERSE = list(dict.fromkeys(x.upper().replace(" ", "") for x in UNIVERSE))

YF_ALIASES = {
    "ZOMATO": "ETERNAL",
    "CANARABANK": "CANBK",
    "TATAMOTORS": "TMCV",
}


# =========================
# DATA
# =========================

def get_history(symbol: str) -> pd.DataFrame:
    """Download daily OHLCV data for one NSE symbol."""
    yf_symbol = YF_ALIASES.get(symbol, symbol)

    # Yahoo symbol handling:
    # - NSE equities need .NS
    # - Index symbols such as ^NSEI already contain their full Yahoo symbol
    ticker = yf_symbol if yf_symbol.startswith("^") else yf_symbol + ".NS"

    try:
        last_error = None

        for attempt in range(3):
            try:
                h = yf.download(
                    ticker,
                    period="1y",
                    auto_adjust=False,
                    progress=False,
                    actions=False,
                    threads=False,
                    timeout=20,
                )
                break
            except Exception as exc:
                last_error = exc
                if attempt < 2:
                    time.sleep(1.5)
                else:
                    return pd.DataFrame()


        if h is None or h.empty:
            return pd.DataFrame()

        if isinstance(h.columns, pd.MultiIndex):
            h.columns = h.columns.get_level_values(0)

        h.columns = [str(c).title() for c in h.columns]
        cols = ["Open", "High", "Low", "Close", "Volume"]

        if not all(c in h.columns for c in cols):
            return pd.DataFrame()

        h = h[cols].dropna(subset=["Close"]).copy()
        h.index = pd.to_datetime(h.index).tz_localize(None)
        return h

    except Exception:
        return pd.DataFrame()


def rsi14(close: pd.Series) -> pd.Series:
    d = close.diff()
    gain = d.clip(lower=0).rolling(14).mean()
    loss = (-d.clip(upper=0)).rolling(14).mean()
    return 100 - 100 / (1 + gain / loss.replace(0, 0.001))


def atr14(h: pd.DataFrame) -> pd.Series:
    prev = h["Close"].shift(1)

    tr = pd.concat(
        [
            h["High"] - h["Low"],
            (h["High"] - prev).abs(),
            (h["Low"] - prev).abs(),
        ],
        axis=1,
    ).max(axis=1)

    return tr.rolling(14).mean()


# =========================
# V2 = YOUR NEW MODEL
# =========================

def v2_score(sym: str, h: pd.DataFrame, nifty: pd.DataFrame):
    if len(h) < 60:
        return None

    close = h["Close"]
    vol = h["Volume"]

    c = float(close.iloc[-1])

    if not (MIN_PRICE <= c <= MAX_PRICE):
        return None

    if float(vol.iloc[-1]) < MIN_VOLUME:
        return None

    prev = float(close.iloc[-2])

    sma10 = float(close.tail(10).mean())
    sma20 = float(close.tail(20).mean())
    sma50 = float(close.tail(50).mean())

    rsi = float(rsi14(close).iloc[-1])
    atr = float(atr14(h).iloc[-1])

    if pd.isna(rsi) or pd.isna(atr) or atr <= 0:
        return None

    vr = float(vol.iloc[-1] / max(float(vol.tail(20).mean()), 1))

    chg = (c / prev - 1) * 100
    mom3 = (c / float(close.iloc[-4]) - 1) * 100
    mom5 = (c / float(close.iloc[-6]) - 1) * 100

    day_range = max(
        float(h["High"].iloc[-1] - h["Low"].iloc[-1]), 0.01
    )
    range_pos = (
        c - float(h["Low"].iloc[-1])
    ) / day_range

    atr_pct = atr / c * 100
    dist20 = (c / sma20 - 1) * 100

    score = 0.0

    # Same V2/New scoring logic.
    score += max(-2, min(4, mom3 * 0.9))
    score += max(-2, min(3, mom5 * 0.5))

    score += (
        2 if c > sma10 > sma20
        else 1 if c > sma20
        else 0
    )

    score += 1.5 if sma20 > sma50 else 0

    score += (
        2 if 52 <= rsi <= 68
        else 1 if 45 <= rsi < 52 or 68 < rsi <= 72
        else 0
    )

    score += (
        2 if vr >= 1.5
        else 1 if vr >= 1.1
        else 0
    )

    score += 1 if range_pos >= 0.65 else 0

    score += 1 if -3 <= dist20 <= 6 else 0

    score -= 2 if dist20 > 8 else 0

    score -= 1 if atr_pct < 1 else 0
    score -= 1 if atr_pct > 7 else 0

    # Relative strength vs Nifty.
    rs = 0.0

    if nifty is not None and not nifty.empty and len(nifty) >= 6:
        nmom5 = (
            float(nifty["Close"].iloc[-1])
            / float(nifty["Close"].iloc[-6])
            - 1
        ) * 100

        rs = mom5 - nmom5

        score += (
            2 if rs > 2
            else 1 if rs > 0.5
            else -1 if rs < -2
            else 0
        )

    return {
        "Symbol": sym,
        "V2 Score": round(score, 2),
        "Close": round(c, 2),
        "RSI": round(rsi, 1),
        "Vol Ratio": round(vr, 2),
        "Mom3 %": round(mom3, 2),
        "Mom5 %": round(mom5, 2),
        "ATR %": round(atr_pct, 2),
        "RS vs Nifty %": round(rs, 2),
    }


# =========================
# V3 = YOUR V3 MODEL
# =========================

def v3_score(sym: str, h: pd.DataFrame, nifty: pd.DataFrame):
    if len(h) < 80:
        return None

    close = h["Close"]
    vol = h["Volume"]

    c = float(close.iloc[-1])

    if not (MIN_PRICE <= c <= MAX_PRICE):
        return None

    if float(vol.iloc[-1]) < MIN_VOLUME:
        return None

    sma5 = float(close.tail(5).mean())
    sma10 = float(close.tail(10).mean())
    sma20 = float(close.tail(20).mean())
    sma50 = float(close.tail(50).mean())

    r = float(rsi14(close).iloc[-1])
    atr = float(atr14(h).iloc[-1])

    if pd.isna(r) or pd.isna(atr) or atr <= 0:
        return None

    atr_pct = atr / c * 100
    vr = float(vol.iloc[-1] / max(float(vol.tail(20).mean()), 1))

    mom1 = (c / float(close.iloc[-2]) - 1) * 100
    mom3 = (c / float(close.iloc[-4]) - 1) * 100
    mom5 = (c / float(close.iloc[-6]) - 1) * 100
    mom10 = (c / float(close.iloc[-11]) - 1) * 100

    recent = close.tail(6).diff().dropna()
    up_days = int((recent > 0).sum())
    down_days = int((recent < 0).sum())

    hi20 = float(h["High"].tail(20).max())
    lo20 = float(h["Low"].tail(20).min())

    range20 = max(hi20 - lo20, 0.01)
    range_pos20 = (c - lo20) / range20

    dist20 = (c / sma20 - 1) * 100
    near_high = (c / hi20 - 1) * 100

    score = 0.0

    score += max(-2, min(3, mom3 * 0.75))
    score += max(-2, min(3, mom5 * 0.45))
    score += max(-1.5, min(2, mom10 * 0.20))

    score += (
        2.5 if c > sma5 > sma10 > sma20
        else 1.5 if c > sma10 > sma20
        else 0.5 if c > sma20
        else 0
    )

    score += 1.5 if sma20 > sma50 else 0
    score += 1.0 if (sma10 / sma20 - 1) * 100 > 0.4 else 0

    score += (
        2.0 if 55 <= r <= 68
        else 1.0 if 50 <= r < 55 or 68 < r <= 72
        else -1.0 if r > 75
        else 0
    )

    score += (
        2.0 if vr >= 1.8
        else 1.0 if vr >= 1.25
        else 0
    )

    score += (
        1.5 if up_days >= 4
        else 0.75 if up_days == 3
        else 0
    )

    score -= 1.0 if down_days >= 4 else 0

    score += (
        1.5 if range_pos20 >= 0.80
        else 0.75 if range_pos20 >= 0.65
        else 0
    )

    score += 1.0 if near_high >= -2.0 else 0

    score += (
        1.0 if 1.5 <= atr_pct <= 4.5
        else 0.25 if 1.0 <= atr_pct < 1.5 or 4.5 < atr_pct <= 6.0
        else -1.5 if atr_pct > 7.0
        else -1.0
    )

    score += (
        1.0 if -2 <= dist20 <= 5
        else -1.5 if dist20 > 8
        else 0
    )

    rs = 0.0

    if nifty is not None and not nifty.empty and len(nifty) >= 6:
        nmom5 = (
            float(nifty["Close"].iloc[-1])
            / float(nifty["Close"].iloc[-6])
            - 1
        ) * 100

        rs = mom5 - nmom5

        score += (
            2.0 if rs > 2.0
            else 1.0 if rs > 0.75
            else -1.0 if rs < -2.0
            else 0
        )

    # Avoid chasing a single abnormal daily jump.
    if mom1 > 5:
        score -= 2.0
    elif mom1 > 3:
        score -= 1.0

    return {
        "Symbol": sym,
        "V3 Score": round(score, 2),
        "Close": round(c, 2),
        "RSI": round(r, 1),
        "Vol Ratio": round(vr, 2),
        "Mom1 %": round(mom1, 2),
        "Mom3 %": round(mom3, 2),
        "Mom5 %": round(mom5, 2),
        "Mom10 %": round(mom10, 2),
        "ATR %": round(atr_pct, 2),
        "RS vs Nifty %": round(rs, 2),
        "Up Days/5": up_days,
        "Dist SMA20 %": round(dist20, 2),
    }


# =========================
# TRACKER
# =========================

def save_history(rows):
    new_df = pd.DataFrame(rows)

    csv_path = Path(CSV_FILE)

    if csv_path.exists():
        try:
            old_df = pd.read_csv(csv_path)
            all_df = pd.concat([old_df, new_df], ignore_index=True)
        except Exception:
            all_df = new_df.copy()
    else:
        all_df = new_df.copy()

    # Prevent duplicates when you run the scanner more than once
    # for the same model/date/rank.
    if not all_df.empty:
        all_df = all_df.drop_duplicates(
            subset=["Scan Date", "Model", "Rank"],
            keep="last",
        )

    all_df.to_csv(csv_path, index=False)

    # Excel output.
    wb = Workbook()
    ws = wb.active
    ws.title = "Latest Scan"

    latest_date = new_df["Scan Date"].iloc[0]
    latest = new_df[new_df["Scan Date"] == latest_date].copy()

    headers = list(latest.columns)

    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1A3A5C")

    for r, row in enumerate(latest.itertuples(index=False), 2):
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for c, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(c)].width = min(
            max(len(str(h)) + 2, 12), 24
        )

    hist_ws = wb.create_sheet("History")
    hist_headers = list(all_df.columns)

    for c, h in enumerate(hist_headers, 1):
        cell = hist_ws.cell(1, c, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1A3A5C")

    for r, row in enumerate(all_df.itertuples(index=False), 2):
        for c, value in enumerate(row, 1):
            hist_ws.cell(r, c, value)

    hist_ws.freeze_panes = "A2"
    hist_ws.auto_filter.ref = hist_ws.dimensions

    for c, h in enumerate(hist_headers, 1):
        hist_ws.column_dimensions[get_column_letter(c)].width = min(
            max(len(str(h)) + 2, 12), 24
        )

    wb.save(XLSX_FILE)

    return all_df


# =========================
# LIVE SCAN
# =========================

def main():
    print("\n" + "=" * 78)
    print("NSE LIVE V2 + V3 STOCK SCANNER")
    print("=" * 78)

    print("\nDownloading latest daily market data...\n")

    data = {}
    failures = []

    for i, sym in enumerate(UNIVERSE, 1):
        h = get_history(sym)

        if not h.empty:
            data[sym] = h
        else:
            failures.append(sym)

        if i % 10 == 0 or i == len(UNIVERSE):
            print(f"{i:>3}/{len(UNIVERSE)} processed")

        time.sleep(0.03)

    nifty = get_history("^NSEI")

    if nifty.empty:
        print("\nERROR: Could not download Nifty data.")
        return

    # Use the latest available daily market date.
    scan_date = nifty.index[-1].date()

    scan_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    v2_rows = []
    v3_rows = []

    for sym, h in data.items():

        h = h[h.index.date <= scan_date].copy()

        if h.empty:
            continue

        v2 = v2_score(sym, h, nifty[nifty.index.date <= scan_date])
        v3 = v3_score(sym, h, nifty[nifty.index.date <= scan_date])

        if v2:
            v2_rows.append(v2)

        if v3:
            v3_rows.append(v3)

    if not v2_rows and not v3_rows:
        print("\nNo stocks passed the filters.")
        return

    v2_df = pd.DataFrame(v2_rows)
    v3_df = pd.DataFrame(v3_rows)

    if not v2_df.empty:
        v2_df = v2_df.sort_values(
            ["V2 Score", "RSI"],
            ascending=[False, False],
        ).head(TOP_N).reset_index(drop=True)

    if not v3_df.empty:
        v3_df = v3_df.sort_values(
            ["V3 Score", "RSI"],
            ascending=[False, False],
        ).head(TOP_N).reset_index(drop=True)

    rows = []

    for rank, row in enumerate(v2_df.to_dict("records"), 1):
        row["Scan Date"] = str(scan_date)
        row["Scan Time"] = scan_time
        row["Model"] = "V2"
        row["Rank"] = rank
        row["Trade Session"] = "NEXT NSE SESSION"
        rows.append(row)

    for rank, row in enumerate(v3_df.to_dict("records"), 1):
        row["Scan Date"] = str(scan_date)
        row["Scan Time"] = scan_time
        row["Model"] = "V3"
        row["Rank"] = rank
        row["Trade Session"] = "NEXT NSE SESSION"
        rows.append(row)

    # Put tracking columns first.
    ordered = []

    for row in rows:
        base = {
            "Scan Date": row.pop("Scan Date"),
            "Scan Time": row.pop("Scan Time"),
            "Trade Session": row.pop("Trade Session"),
            "Model": row.pop("Model"),
            "Rank": row.pop("Rank"),
        }
        base.update(row)
        ordered.append(base)

    history = save_history(ordered)

    # =========================
    # CONSOLE OUTPUT
    # =========================

    print("\n" + "=" * 78)
    print(f"SCAN DATE: {scan_date}")
    print("TRADE:     NEXT NSE SESSION")
    print("=" * 78)

    print("\nV2 TOP PICKS")
    print("-" * 78)
    if v2_df.empty:
        print("No V2 candidates.")
    else:
        print(
            v2_df[
                ["Symbol", "V2 Score", "Close", "RSI",
                 "Vol Ratio", "Mom3 %", "Mom5 %", "RS vs Nifty %"]
            ].to_string(index=False)
        )

    print("\nV3 TOP PICKS")
    print("-" * 78)
    if v3_df.empty:
        print("No V3 candidates.")
    else:
        print(
            v3_df[
                ["Symbol", "V3 Score", "Close", "RSI",
                 "Vol Ratio", "Mom3 %", "Mom5 %",
                 "ATR %", "RS vs Nifty %"]
            ].to_string(index=False)
        )

    # Common candidates in both V2 and V3.
    common = sorted(
        set(v2_df["Symbol"].tolist()) &
        set(v3_df["Symbol"].tolist())
    )

    print("\nCOMMON V2 + V3 PICKS")
    print("-" * 78)

    if common:
        print(", ".join(common))
    else:
        print("No common top picks today.")

    print("\n" + "=" * 78)
    print("FILES SAVED")
    print(f"CSV : {Path(CSV_FILE).resolve()}")
    print(f"XLSX: {Path(XLSX_FILE).resolve()}")
    print("=" * 78)

    if failures:
        print(
            f"\nNote: {len(failures)} symbols returned no data "
            f"and were skipped."
        )


if __name__ == "__main__":
    main()
