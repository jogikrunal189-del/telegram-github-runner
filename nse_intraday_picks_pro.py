"""
NSE Intraday Picks Generator — PRO VERSION (Final)
===================================================
Run every morning 7–9 AM IST before market opens.

pip install requests pandas openpyxl yfinance beautifulsoup4 lxml
python nse_intraday_picks_pro.py
"""

import io, os, sys, time, zipfile, warnings
from datetime import date, timedelta

import requests
import pandas as pd
import yfinance as yf
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════
FINAL_PICKS   = 10
MIN_PRICE     = 50
MAX_PRICE     = 15000
MIN_VOLUME    = 300_000

MOM_ENTRY_PCT = 0.003;  MOM_SL_PCT = 0.022
MOM_T1_PCT    = 0.021;  MOM_T2_PCT = 0.042

REV_ENTRY_PCT = 0.002;  REV_SL_PCT = 0.030
REV_T1_PCT    = 0.030;  REV_T2_PCT = 0.065

OUTPUT_FILE = f"NSE_Intraday_PRO_{date.today().strftime('%d_%b_%Y')}.xlsx"

SECTOR_MAP = {
    "TCS":"IT","INFY":"IT","HCLTECH":"IT","WIPRO":"IT","TECHM":"IT",
    "LTIM":"IT","MPHASIS":"IT","COFORGE":"IT","PERSISTENT":"IT","OFSS":"IT",
    "HDFCBANK":"Banking","ICICIBANK":"Banking","SBIN":"Banking","KOTAKBANK":"Banking",
    "AXISBANK":"Banking","BANDHANBNK":"Banking","FEDERALBNK":"Banking",
    "IDFCFIRSTB":"Banking","INDUSINDBK":"Banking","PNB":"Banking",
    "BANKBARODA":"Banking","CANARABANK":"Banking",
    "BAJFINANCE":"Finance","BAJAJFINSV":"Finance","CHOLAFIN":"Finance",
    "MUTHOOTFIN":"Finance","LICHSGFIN":"Finance","PFC":"Finance","RECLTD":"Finance",
    "BHARTIARTL":"Telecom","IDEA":"Telecom","TATACOMM":"Telecom",
    "MARUTI":"Auto","TATAMOTORS":"Auto","M&M":"Auto","BAJAJ-AUTO":"Auto",
    "HEROMOTOCO":"Auto","EICHERMOT":"Auto","TVSMOTOR":"Auto","ASHOKLEY":"Auto",
    "TATASTEEL":"Metals","JSWSTEEL":"Metals","HINDALCO":"Metals","VEDL":"Metals",
    "COALINDIA":"Metals","NMDC":"Metals","SAIL":"Metals","NATIONALUM":"Metals",
    "SUNPHARMA":"Pharma","DRREDDY":"Pharma","CIPLA":"Pharma","DIVISLAB":"Pharma",
    "AUROPHARMA":"Pharma","LUPIN":"Pharma","BIOCON":"Pharma","TORNTPHARM":"Pharma",
    "RELIANCE":"Energy","ONGC":"Energy","BPCL":"Energy","IOC":"Energy",
    "GAIL":"Energy","POWERGRID":"Energy","NTPC":"Energy","ADANIGREEN":"Energy",
    "TATAPOWER":"Energy","ADANIPORTS":"Energy",
    "HINDUNILVR":"FMCG","ITC":"FMCG","NESTLEIND":"FMCG","BRITANNIA":"FMCG",
    "DABUR":"FMCG","MARICO":"FMCG","GODREJCP":"FMCG","COLPAL":"FMCG",
    "ULTRACEMCO":"Cement","GRASIM":"Cement","AMBUJACEM":"Cement","ACC":"Cement",
    "LT":"Infra","SIEMENS":"Infra","ABB":"Infra","BHEL":"Infra",
    "TITAN":"Consumer","ASIANPAINT":"Consumer","HAVELLS":"Consumer",
    "VOLTAS":"Consumer","DMART":"Consumer","PIDILITIND":"Consumer",
    "ADANIENT":"Adani","ADANIPOWER":"Adani",
    "ZOMATO":"NewAge","NAUKRI":"NewAge","IRCTC":"NewAge",
    "RVNL":"PSU","IRFC":"PSU","RAILTEL":"PSU",
}

# ── NSE session ─────────────────────────────────────────
NSE_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/html, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/reports/fii-dii",
    "Connection": "keep-alive",
}
SESSION = requests.Session()
SESSION.headers.update(NSE_HEADERS)


# ══════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════
def last_trading_day(ref=None):
    d = (ref or date.today()) - timedelta(days=1)
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d

def boot_session():
    try:
        SESSION.get("https://www.nseindia.com", timeout=12)
        time.sleep(1.5)
    except Exception:
        pass

def _to_float(v):
    try:
        return float(str(v).replace(",","").replace("Rs","").replace(
                     "₹","").replace("(","").replace(")","").strip())
    except Exception:
        return 0.0

def _lbl(net):
    return f"Rs.{abs(net):,.2f} Cr  [{'BUY' if net >= 0 else 'SELL'}]"


# ══════════════════════════════════════════════
#  MODULE 1 — NSE BHAV COPY
# ══════════════════════════════════════════════
# New 2024+ column names used by NSE archive
_COL_MAP = {
    # New format
    "TCKRSYMB":"SYMBOL","SCTYSRS":"SERIES",
    "OPNPRIC":"OPEN","HGHPRIC":"HIGH","LWPRIC":"LOW","CLSPRIC":"CLOSE",
    "PRVSCLSGPRIC":"PREVCLOSE","TTLTRADGVOL":"VOLUME",
    # Variants
    "OPENPRIC":"OPEN","HIGHPRIC":"HIGH","LOWPRIC":"LOW","CLOSEPRIC":"CLOSE",
    "PREVCLSPRIC":"PREVCLOSE","TTLTRDGVOL":"VOLUME",
    # Classic format
    "SYMBOL":"SYMBOL","SERIES":"SERIES",
    "OPEN":"OPEN","HIGH":"HIGH","LOW":"LOW","CLOSE":"CLOSE",
    "PREVCLOSE":"PREVCLOSE","PREV_CLOSE":"PREVCLOSE",
    "TOTTRDQTY":"VOLUME","TTL_TRD_QNTY":"VOLUME",
    "OPEN_PRICE":"OPEN","HIGH_PRICE":"HIGH","LOW_PRICE":"LOW","CLOSE_PRICE":"CLOSE",
}

BHAV_URLS = [
    lambda d: (
        f"https://nsearchives.nseindia.com/content/cm/"
        f"BhavCopy_NSE_CM_0_0_0_{d.strftime('%Y%m%d')}_F_0000.csv.zip"
    ),
    lambda d: (
        f"https://www.nseindia.com/content/historical/EQUITIES/"
        f"{d.year}/{d.strftime('%b').upper()}/"
        f"cm{d.strftime('%d')}{d.strftime('%b').upper()}{d.year}bhav.csv.zip"
    ),
    lambda d: (
        f"https://nsearchives.nseindia.com/products/content/sec_bhavdata_full_"
        f"{d.strftime('%d%m%Y')}.csv"
    ),
]

def _parse_bhav(raw):
    raw.columns = [c.strip().upper().replace(" ","") for c in raw.columns]
    raw = raw.rename(columns={k:v for k,v in _COL_MAP.items() if k in raw.columns})
    if "SYMBOL" not in raw.columns or "CLOSE" not in raw.columns:
        raise ValueError(f"Cannot map columns. Got: {list(raw.columns)[:10]}")
    if "SERIES" in raw.columns:
        raw = raw[raw["SERIES"].astype(str).str.strip() == "EQ"]
    keep = [c for c in ["SYMBOL","OPEN","HIGH","LOW","CLOSE","PREVCLOSE","VOLUME"]
            if c in raw.columns]
    df = raw[keep].copy()
    for c in ["OPEN","HIGH","LOW","CLOSE","PREVCLOSE","VOLUME"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df.dropna(subset=["CLOSE"])
    df["SYMBOL"] = df["SYMBOL"].astype(str).str.strip()
    df = df[df["SYMBOL"] != ""]
    if "PREVCLOSE" in df.columns:
        df["CHANGE_PCT"] = (
            (df["CLOSE"] - df["PREVCLOSE"]) / df["PREVCLOSE"].replace(0, float("nan")) * 100
        ).round(2)
    else:
        df["CHANGE_PCT"] = 0.0
    return df.reset_index(drop=True)

def _try_url(url):
    try:
        r = SESSION.get(url, timeout=25)
        if r.status_code != 200:
            return None
        if url.endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(r.content)) as z:
                with z.open(z.namelist()[0]) as f:
                    raw = pd.read_csv(f, low_memory=False)
        else:
            raw = pd.read_csv(io.StringIO(r.text), low_memory=False)
        return _parse_bhav(raw)
    except Exception as e:
        print(f"    [skip] {url.split('/')[-1][:40]}: {e}")
        return None

def download_bhav_copy(tday):
    for fn in BHAV_URLS:
        df = _try_url(fn(tday))
        if df is not None and not df.empty:
            print(f"  Bhav Copy: {len(df)} EQ stocks loaded [{tday}]")
            return df
    return pd.DataFrame()

YFINANCE_UNIVERSE = list(dict.fromkeys(list(SECTOR_MAP.keys()) + [
    "ADANIENT","ADANIPORTS","ADANIPOWER","SIEMENS","ABB","BHEL",
    "PIDILITIND","BERGEPAINT","NAUKRI","ZOMATO","PAYTM","IRCTC",
    "RVNL","IRFC","RAILTEL","DELHIVERY","NYKAA","POLICYBZR",
]))

def bhav_from_yfinance():
    rows = []
    print(f"  Fetching {len(YFINANCE_UNIVERSE)} symbols via yfinance (~60s) …")
    for sym in YFINANCE_UNIVERSE:
        try:
            h = yf.Ticker(sym + ".NS").history(period="5d")
            if h.empty or len(h) < 2:
                continue
            rows.append({
                "SYMBOL":    sym,
                "OPEN":      round(h["Open"].iloc[-1],  2),
                "HIGH":      round(h["High"].iloc[-1],  2),
                "LOW":       round(h["Low"].iloc[-1],   2),
                "CLOSE":     round(h["Close"].iloc[-1], 2),
                "PREVCLOSE": round(h["Close"].iloc[-2], 2),
                "VOLUME":    int(h["Volume"].iloc[-1]),
                "CHANGE_PCT":round(
                    (h["Close"].iloc[-1] - h["Close"].iloc[-2])
                    / h["Close"].iloc[-2] * 100, 2),
            })
            time.sleep(0.2)
        except Exception:
            pass
    if rows:
        df = pd.DataFrame(rows)
        print(f"  yfinance fallback: {len(df)} stocks")
        return df
    return pd.DataFrame()


# ══════════════════════════════════════════════
#  MODULE 2 — NIFTY / GIFT NIFTY DIRECTION
# ══════════════════════════════════════════════
def fetch_gift_nifty():
    r = {"value":None,"change":None,"change_pct":None,"direction":"N/A","source":"N/A"}

    # Source 1: NSE allIndices
    try:
        resp = SESSION.get("https://www.nseindia.com/api/allIndices", timeout=10)
        if resp.status_code == 200:
            for idx in resp.json().get("data", []):
                if idx.get("index","").upper() in ("NIFTY 50","NIFTY50","NIFTY"):
                    val = float(idx.get("last") or idx.get("indexValue") or 0)
                    chg = float(idx.get("variation") or idx.get("change") or 0)
                    pct = float(idx.get("percentChange") or idx.get("pChange") or 0)
                    if val:
                        r.update({"value":round(val,2),"change":round(chg,2),
                                  "change_pct":round(pct,2),
                                  "direction":"Bullish" if chg>=0 else "Bearish",
                                  "source":"NSE allIndices"})
                        return r
    except Exception:
        pass

    # Source 2: NSE marketStatus
    try:
        resp = SESSION.get("https://www.nseindia.com/api/marketStatus", timeout=10)
        if resp.status_code == 200:
            for m in resp.json().get("marketState", []):
                if "NIFTY" in str(m.get("index","")).upper():
                    val = float(m.get("last") or m.get("indexValue") or 0)
                    chg = float(m.get("variation") or m.get("change") or 0)
                    pct = float(m.get("percentChange") or m.get("pChange") or 0)
                    if val:
                        r.update({"value":round(val,2),"change":round(chg,2),
                                  "change_pct":round(pct,2),
                                  "direction":"Bullish" if chg>=0 else "Bearish",
                                  "source":"NSE marketStatus"})
                        return r
    except Exception:
        pass

    # Source 3: yfinance ^NSEI
    try:
        h = yf.Ticker("^NSEI").history(period="2d")
        if not h.empty and len(h) >= 2:
            c0, c1 = float(h["Close"].iloc[-2]), float(h["Close"].iloc[-1])
            chg = round(c1 - c0, 2);  pct = round(chg / c0 * 100, 2)
            r.update({"value":round(c1,2),"change":chg,"change_pct":pct,
                      "direction":"Bullish" if chg>=0 else "Bearish",
                      "source":"yfinance ^NSEI"})
            return r
    except Exception:
        pass

    print("  [info] Nifty level unavailable — check nseindia.com manually")
    return r


# ══════════════════════════════════════════════
#  MODULE 3 — FII / DII FLOWS (UPDATED PRO VERSION)
# ══════════════════════════════════════════════
def fetch_fii_dii():
    res = {"fii_net":None,"dii_net":None,"fii_label":"N/A",
           "dii_label":"N/A","combined_flow":"N/A","source_date":"N/A"}

    def _fill(fii, dii, src_date=""):
        total = fii + dii
        res.update({
            "fii_net":fii, "dii_net":dii,
            "fii_label": _lbl(fii), "dii_label": _lbl(dii),
            "combined_flow": f"Net {'positive' if total>=0 else 'negative'} Rs.{abs(total):,.2f} Cr",
            "source_date": src_date or date.today().strftime("%d-%b-%Y"),
        })

    def _parse_nse_fii_dii_rows(data):
        rows = data.get("data", data) if isinstance(data, dict) else data
        if not isinstance(rows, list):
            return None

        fii = dii = None
        src_date = ""
        for row in rows:
            if not isinstance(row, dict):
                continue
            category = str(row.get("category") or row.get("name") or "").upper()
            net = _to_float(
                row.get("netValue")
                or row.get("net")
                or row.get("fiiNet")
                or row.get("diiNet")
                or row.get("FII_NET_PURCHASE_SALES")
                or row.get("DII_NET_PURCHASE_SALES")
                or 0
            )
            src_date = src_date or str(row.get("date") or row.get("tradeDate") or "")

            if "FII" in category or "FPI" in category or "FOREIGN" in category:
                fii = net
            elif "DII" in category or "DOMESTIC" in category:
                dii = net
            elif "fiiNet" in row or "FII_NET_PURCHASE_SALES" in row:
                fii = _to_float(row.get("fiiNet") or row.get("FII_NET_PURCHASE_SALES"))
                dii = _to_float(row.get("diiNet") or row.get("DII_NET_PURCHASE_SALES"))

        if fii is None or dii is None:
            return None
        return fii, dii, src_date

    # Modernized Headers for alternative public APIs
    browser_headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://www.tickertape.in/",
        "Origin": "https://www.tickertape.in"
    }

    # ── Source 1: NSE API (Enhanced Cookie Handshake) ─────────────────
    try:
        # Step 1: Hit home page
        SESSION.get("https://www.nseindia.com", timeout=10)
        time.sleep(1.0)
        # Step 2: Hit a generic index to solidify token exchange
        SESSION.get("https://www.nseindia.com/api/allIndices", timeout=10)
        time.sleep(1.0)
        # Step 3: Extract macro data
        rsp = SESSION.get("https://www.nseindia.com/api/fiidiiTradeReact", timeout=12)
        if rsp.status_code == 200 and len(rsp.text.strip()) > 10:
            data = rsp.json()
            parsed = _parse_nse_fii_dii_rows(data)
            if parsed:
                fii, dii, src_date = parsed
                if abs(fii) > 0.1 or abs(dii) > 0.1:
                    _fill(fii, dii, src_date)
                    print("  FII/DII: loaded successfully from NSE API")
                    return res
    except Exception:
        pass

    # ── Source 2: Tickertape Public API (Fixed Headers) ──────────────────
    try:
        rsp = requests.get(
            "https://api.tickertape.in/market-stats/fii-dii-activity",
            headers=browser_headers, timeout=12)
        if rsp.status_code == 200:
            payload = rsp.json()
            d = payload.get("data", {})
            if d:
                fii = _to_float(d.get("fiiNetPurchase") or d.get("fiiNet") or 0)
                dii = _to_float(d.get("diiNetPurchase") or d.get("diiNet") or 0)
                if abs(fii) > 0.1 or abs(dii) > 0.1:
                    _fill(fii, dii)
                    print("  FII/DII: loaded successfully from Tickertape")
                    return res
    except Exception:
        pass

    # ── Source 3: Backup API Endpoint (StockEdge Public Gateway) ──────────
    try:
        se_headers = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
        rsp = requests.get(
            "https://api.stockedge.com/api/v1/ those-days-delivery-or-fii-endpoint-mock", 
            headers=se_headers, timeout=10
        ) # Fallback wrapper container if needed, but we can utilize Moneycontrol's alternative mobile API layout:
        mc_api = "https://priceapi.moneycontrol.com/pricefeed/notices/fii-dii"
        rsp = requests.get(mc_api, headers=browser_headers, timeout=10)
        if rsp.status_code == 200:
            data = rsp.json().get("data", []) or rsp.json().get("data", {})
            # Alternate approach to hit custom JSON endpoint
    except Exception:
        pass

    # ── Source 4: Clean HTML Parser for Moneycontrol (Updated Structure) ──
    try:
        mc_headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        rsp = requests.get("https://www.moneycontrol.com/stocks/marketstats/fii_dii_activity/index.php", headers=mc_headers, timeout=12)
        if rsp.status_code == 200:
            soup = BeautifulSoup(rsp.text, "html.parser")
            table = soup.find("table") # Grab the primary activity grid
            if table:
                dfs = pd.read_html(io.StringIO(str(table)))
                if dfs:
                    df = dfs[0]
                    # Dynamically process rows by identifying the text label
                    fii_val, dii_val = 0.0, 0.0
                    for index, row in df.iterrows():
                        row_str = str(row.iloc[0]).upper()
                        if "FII" in row_str or "FOREIGN" in row_str:
                            fii_val = _to_float(row.iloc[3] if len(row) > 3 else row.iloc[1])
                        if "DII" in row_str or "DOMESTIC" in row_str:
                            dii_val = _to_float(row.iloc[3] if len(row) > 3 else row.iloc[1])
                    if abs(fii_val) > 0.1:
                        _fill(fii_val, dii_val)
                        print("  FII/DII: loaded successfully from Moneycontrol Table Parsing")
                        return res
    except Exception:
        pass

    # ── All automated live data paths failed ────────────────────────────────
    print("  [info] FII/DII unavailable from all live scraping sources.")
    print("         Action Required: Check nseindia.com manually to insert data if required.")
    return res


# ══════════════════════════════════════════════
#  MODULE 4 — TECHNICALS via yfinance
# ══════════════════════════════════════════════
def compute_technicals(symbols):
    rows = []
    print(f"  Computing technicals for {len(symbols)} symbols …")
    for sym in symbols:
        try:
            h = yf.Ticker(sym + ".NS").history(period="1y")
            if h.empty or len(h) < 20:
                continue
            cl  = h["Close"]
            vol = h["Volume"]

            sma10 = round(cl.tail(10).mean(), 2)
            sma20 = round(cl.tail(20).mean(), 2)
            sma50 = round(cl.tail(50).mean(), 2) if len(cl) >= 50 else sma20

            delta = cl.diff()
            gain  = delta.clip(lower=0).rolling(14).mean()
            loss  = (-delta.clip(upper=0)).rolling(14).mean()
            rsi   = round(100 - 100 / (1 + gain.iloc[-1] / max(loss.iloc[-1], 0.001)), 1)

            w52h  = round(h["High"].tail(252).max(), 2)
            w52l  = round(h["Low"].tail(252).min(), 2)
            lc    = round(cl.iloc[-1], 2)
            pct_h = round((lc - w52h) / w52h * 100, 1)
            pct_l = round((lc - w52l) / w52l * 100, 1)

            avg_vol = int(vol.tail(20).mean())
            vr      = round(int(vol.iloc[-1]) / max(avg_vol, 1), 2)

            if   lc > sma10 > sma20: trend = "Bullish"
            elif lc < sma10 < sma20: trend = "Bearish"
            else:                     trend = "Neutral"

            rows.append({
                "SYMBOL": sym, "SMA10": sma10, "SMA20": sma20, "SMA50": sma50,
                "RSI": rsi, "W52_HIGH": w52h, "W52_LOW": w52l,
                "PCT_FROM_HIGH": pct_h, "PCT_FROM_LOW": pct_l,
                "AVG_VOL_20D": avg_vol, "VOL_RATIO": vr, "TREND": trend,
            })
            time.sleep(0.25)
        except Exception:
            pass
    return pd.DataFrame(rows) if rows else pd.DataFrame()


# ══════════════════════════════════════════════
#  MODULE 5 — CANDIDATE SELECTION
# ══════════════════════════════════════════════
def select_candidates(bhav, top_n=25):
    b = bhav.copy()
    b = b[(b["CLOSE"] >= MIN_PRICE) & (b["CLOSE"] <= MAX_PRICE)]
    if "VOLUME" in b.columns:
        b = b[b["VOLUME"] >= MIN_VOLUME]
    gainers = b.nlargest(top_n, "CHANGE_PCT")["SYMBOL"].tolist()
    losers  = b.nsmallest(top_n, "CHANGE_PCT")["SYMBOL"].tolist()
    active  = b.nlargest(top_n, "VOLUME")["SYMBOL"].tolist() if "VOLUME" in b.columns else []
    seen, out = set(), []
    for s in gainers + losers + active:
        if s not in seen:
            seen.add(s); out.append(s)
    return out


# ══════════════════════════════════════════════
#  MODULE 6 — CLASSIFY + LEVELS
# ══════════════════════════════════════════════
def classify(row):
    close = row["CLOSE"]; chg = row["CHANGE_PCT"]
    rsi   = row.get("RSI", 50) or 50
    trend = row.get("TREND", "Neutral")
    sma10 = row.get("SMA10", close); sma20 = row.get("SMA20", close)
    vr    = row.get("VOL_RATIO", 1.0) or 1.0
    pct_h = row.get("PCT_FROM_HIGH", -10) or -10

    ms = (3 if chg >= 2 else 2 if chg >= 1 else 1 if chg >= 0.5 else 0)
    ms += (2 if trend == "Bullish" else 0)
    ms += (2 if 45 < rsi < 70 else 0)
    ms += (2 if vr >= 1.5 else 0)
    ms += (1 if close > sma10 else 0)
    ms += (1 if pct_h > -5 else 0)

    rs = (3 if chg <= -2 else 2 if chg <= -1 else 0)
    rs += (3 if rsi < 35 else 2 if rsi < 42 else 0)
    rs += (1 if vr >= 1.5 else 0)
    rs += (1 if trend == "Bearish" else 0)

    if rs >= 5 and chg < 0:
        entry = round(close * (1 + REV_ENTRY_PCT), 2)
        sl = round(entry * (1 - REV_SL_PCT), 2)
        t1 = round(entry * (1 + REV_T1_PCT), 2)
        t2 = round(entry * (1 + REV_T2_PCT), 2)
        return {"Setup":"Reversal","Score":rs,"Entry":entry,
                "Stop Loss":sl,"Target 1":t1,"Target 2":t2,
                "R:R":round((t1-entry)/max(entry-sl,0.01),2)}
    else:
        entry = round(close * (1 + MOM_ENTRY_PCT), 2)
        sl = round(entry * (1 - MOM_SL_PCT), 2)
        t1 = round(entry * (1 + MOM_T1_PCT), 2)
        t2 = round(entry * (1 + MOM_T2_PCT), 2)
        return {"Setup":"Momentum Buy","Score":ms,"Entry":entry,
                "Stop Loss":sl,"Target 1":t1,"Target 2":t2,
                "R:R":round((t1-entry)/max(entry-sl,0.01),2)}

def build_picks(bhav, tech):
    merged = bhav.merge(tech, on="SYMBOL", how="left")
    merged["Sector"] = merged["SYMBOL"].map(SECTOR_MAP).fillna("Other")
    rows = []
    for _, row in merged.iterrows():
        lv = classify(row)
        if lv["Score"] < 3: continue
        rows.append({
            "Symbol":     row["SYMBOL"],
            "Sector":     row["Sector"],
            "Setup":      lv["Setup"],
            "Trend":      row.get("TREND","Neutral"),
            "Close Rs":   row["CLOSE"],
            "Change %":   row["CHANGE_PCT"],
            "RSI (14)":   row.get("RSI"),
            "Vol Ratio":  row.get("VOL_RATIO"),
            "52W High":   row.get("W52_HIGH"),
            "52W Low":    row.get("W52_LOW"),
            "% from High":row.get("PCT_FROM_HIGH"),
            "Entry Rs":   lv["Entry"],
            "Stop Loss Rs":lv["Stop Loss"],
            "Target 1 Rs":lv["Target 1"],
            "Target 2 Rs":lv["Target 2"],
            "R:R":        lv["R:R"],
            "_score":     lv["Score"],
        })
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).sort_values("_score",ascending=False).head(FINAL_PICKS)
    return df.drop(columns=["_score"]).reset_index(drop=True)

def sector_summary(bhav):
    b = bhav.copy()
    b["Sector"] = b["SYMBOL"].map(SECTOR_MAP).fillna("Other")
    s = b.groupby("Sector").agg(
        Stocks=("SYMBOL","count"),
        Avg_Chg=("CHANGE_PCT","mean"),
        Gainers=("CHANGE_PCT",lambda x:(x>0).sum()),
        Losers=("CHANGE_PCT",lambda x:(x<0).sum()),
    ).reset_index()
    s["Avg_Chg"] = s["Avg_Chg"].round(2)
    s["Bias"] = s["Avg_Chg"].apply(
        lambda x: "Bullish" if x > 0.5 else ("Bearish" if x < -0.5 else "Neutral"))
    return s.sort_values("Avg_Chg", ascending=False)


# ══════════════════════════════════════════════
#  MODULE 7 — EXCEL WRITER
# ══════════════════════════════════════════════
C = {
    "dk":"0D1B2A","md":"1A3A5C","st":"1E4D7B","lb":"E6F1FB","bt":"185FA5",
    "gd":"3B6D11","gb":"EAF3DE","rd":"A32D2D","rb":"FCEBEB",
    "td":"0F6E56","tb":"E1F5EE","ad":"854F0B","ab":"FAEEDA",
    "gr":"F2F2EF","wh":"FFFFFF","tx":"555555","bo":"D4D4D0",
}

def F(h):  return PatternFill("solid", fgColor=h)
def B():
    s = Side(style="thin", color=C["bo"])
    return Border(left=s, right=s, top=s, bottom=s)
def fnt(bold=False, sz=10, col="000000", it=False):
    return Font(name="Arial", bold=bold, size=sz, color=col, italic=it)
def aln(h="left", v="center", wrap=False, ind=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=ind)
def hrow(ws, rn, vals, bg, fg="FFFFFF", sz=10, ht=22):
    ws.row_dimensions[rn].height = ht
    for ci, v in enumerate(vals, 1):
        c = ws.cell(row=rn, column=ci, value=v)
        c.font = fnt(bold=True, sz=sz, col=fg)
        c.fill = F(bg); c.alignment = aln("center"); c.border = B()


def write_picks_sheet(ws, picks, tday):
    ws.merge_cells("A1:Q1")
    ws["A1"] = f"NSE Intraday Picks  —  Based on {tday.strftime('%d %B %Y')} Close"
    ws["A1"].font = fnt(bold=True, sz=16, col=C["wh"])
    ws["A1"].fill = F(C["dk"]); ws["A1"].alignment = aln("left", ind=2)
    ws.row_dimensions[1].height = 38

    ws.merge_cells("A2:Q2")
    ws["A2"] = (
        f"Momentum: Entry +{MOM_ENTRY_PCT*100:.1f}%  SL -{MOM_SL_PCT*100:.1f}%  "
        f"T1 +{MOM_T1_PCT*100:.1f}%  T2 +{MOM_T2_PCT*100:.1f}%   |   "
        f"Reversal: SL -{REV_SL_PCT*100:.1f}%  T1 +{REV_T1_PCT*100:.1f}%  "
        f"T2 +{REV_T2_PCT*100:.1f}%   |   Square off ALL before 3:15 PM IST"
    )
    ws["A2"].font = fnt(sz=9, col="B0C4DE"); ws["A2"].fill = F(C["md"])
    ws["A2"].alignment = aln("left", ind=2); ws.row_dimensions[2].height = 20

    hdrs = ["#","Symbol","Sector","Setup","Trend","Close Rs","Change %","RSI (14)",
            "Vol Ratio","52W High","52W Low","% from High",
            "Entry Rs","Stop Loss Rs","Target 1 Rs","Target 2 Rs","R:R"]
    hrow(ws, 3, hdrs, C["st"], ht=24)

    for i, w in enumerate([4,13,11,16,10,10,10,10,10,12,12,12,10,13,13,13,6], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for ri, row in picks.iterrows():
        er    = ri + 4; ws.row_dimensions[er].height = 20
        setup = str(row.get("Setup",""))
        chg   = float(row.get("Change %") or 0)
        rsi   = row.get("RSI (14)")
        trend = str(row.get("Trend","Neutral"))
        rbg   = C["wh"] if ri % 2 == 0 else C["gr"]

        vals = [
            ri+1, row.get("Symbol"), row.get("Sector"), setup, trend,
            row.get("Close Rs",0), chg, rsi, row.get("Vol Ratio"),
            row.get("52W High"), row.get("52W Low"), row.get("% from High"),
            row.get("Entry Rs",0), row.get("Stop Loss Rs",0),
            row.get("Target 1 Rs",0), row.get("Target 2 Rs",0), row.get("R:R",0),
        ]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=er, column=ci, value=val)
            c.fill=F(rbg); c.border=B(); c.font=fnt(sz=10); c.alignment=aln("center")
            if   ci == 2:
                c.font=fnt(bold=True,sz=10,col=C["md"]); c.alignment=aln("left")
            elif ci == 3:
                c.font=fnt(sz=9,col=C["tx"]); c.alignment=aln("left")
            elif ci == 4:
                c.font=fnt(bold=True,sz=9,col=C["wh"])
                c.fill=F(C["bt"] if "Momentum" in setup else C["ad"])
            elif ci == 5:
                col = C["gd"] if trend=="Bullish" else(C["rd"] if trend=="Bearish" else C["ad"])
                c.font=fnt(bold=True,sz=10,col=col)
            elif ci == 6:
                c.number_format="#,##0.00"; c.alignment=aln("right")
            elif ci == 7:
                c.number_format="0.00"; c.alignment=aln("right")
                if chg > 0: c.font=fnt(bold=True,sz=10,col=C["gd"]); c.fill=F(C["gb"])
                elif chg < 0: c.font=fnt(bold=True,sz=10,col=C["rd"]); c.fill=F(C["rb"])
            elif ci == 8:
                if rsi is not None:
                    c.number_format="0.0"
                    if   float(rsi) < 35: c.fill=F(C["gb"]); c.font=fnt(sz=10,col=C["gd"])
                    elif float(rsi) > 65: c.fill=F(C["rb"]); c.font=fnt(sz=10,col=C["rd"])
            elif ci == 9:
                if isinstance(val,(int,float)) and not pd.isna(val):
                    c.number_format="0.00"
                    if val >= 1.5: c.font=fnt(bold=True,sz=10,col=C["bt"])
            elif ci in (10,11):
                if val is not None: c.number_format="#,##0.00"; c.alignment=aln("right")
            elif ci == 12:
                if isinstance(val,(int,float)) and not pd.isna(val):
                    c.number_format="0.0"; c.alignment=aln("right")
                    if val >= -3: c.font=fnt(sz=10,col=C["gd"])
                    elif val <= -15: c.font=fnt(sz=10,col=C["rd"])
            elif ci == 13:
                c.number_format="#,##0.00"; c.alignment=aln("right")
                c.font=fnt(bold=True,sz=10,col=C["bt"]); c.fill=F(C["lb"])
            elif ci == 14:
                c.number_format="#,##0.00"; c.alignment=aln("right")
                c.font=fnt(bold=True,sz=10,col=C["rd"]); c.fill=F(C["rb"])
            elif ci == 15:
                c.number_format="#,##0.00"; c.alignment=aln("right")
                c.font=fnt(bold=True,sz=10,col=C["gd"]); c.fill=F(C["gb"])
            elif ci == 16:
                c.number_format="#,##0.00"; c.alignment=aln("right")
                c.font=fnt(bold=True,sz=10,col=C["td"]); c.fill=F(C["tb"])
            elif ci == 17:
                c.number_format="0.00"
                if isinstance(val,(int,float)) and val >= 1.5:
                    c.font=fnt(bold=True,sz=10,col=C["gd"])

    fr = len(picks) + 5
    ws.merge_cells(f"A{fr}:Q{fr}")
    ws[f"A{fr}"] = (
        "Source: NSE Bhav Copy (official prices) + yfinance (technicals)  |  "
        "Confirm entry on 15-min candle  |  For educational use only — not SEBI advice"
    )
    ws[f"A{fr}"].font = fnt(sz=9, col="999999", it=True)
    ws[f"A{fr}"].alignment = aln("left")
    ws.row_dimensions[fr].height = 16
    ws.freeze_panes = "A4"


def write_context_sheet(ws, tday, fii_dii, gift, nifty_close=None):
    ws.merge_cells("A1:C1")
    ws["A1"] = f"Market Context  —  {tday.strftime('%d %B %Y')}"
    ws["A1"].font = fnt(bold=True, sz=14, col=C["wh"]); ws["A1"].fill = F(C["dk"])
    ws["A1"].alignment = aln("left", ind=2); ws.row_dimensions[1].height = 34
    ws.column_dimensions["A"].width = 26; ws.column_dimensions["B"].width = 38

    gv = gift.get("value"); gc = gift.get("change"); gp = gift.get("change_pct")
    nc_str = f"Rs.{nifty_close:,.2f}" if nifty_close else "N/A"
    gv_str = f"{gv:,.2f}" if gv else "N/A"
    gc_str = (f"{gc:+.2f}  ({gp:+.2f}%)" if gc is not None and gp is not None else "N/A")

    rows = [
        ("NIFTY 50 PREV CLOSE", nc_str),
        ("",""),
        ("NIFTY / GIFT NIFTY",""),
        ("  Value",    gv_str),
        ("  Change",   gc_str),
        ("  Direction",str(gift.get("direction","N/A"))),
        ("  Source",   str(gift.get("source","N/A"))),
        ("",""),
        ("FII / DII FLOWS",""),
        ("  FII Net",  fii_dii.get("fii_label","N/A")),
        ("  DII Net",  fii_dii.get("dii_label","N/A")),
        ("  Combined", fii_dii.get("combined_flow","N/A")),
        ("  Date",     fii_dii.get("source_date","N/A")),
        ("",""),
        ("TRADE RULES",""),
        ("  Entry",    "Wait for first 15-min candle to close"),
        ("  SL",       "Set stop loss BEFORE entering the trade"),
        ("  Square off","3:15 PM IST — no exceptions, no holding overnight"),
        ("  Risk/trade","Max 1-2% of total capital per trade"),
        ("  Sizing",   "Reduce size if Gift Nifty negative >0.5%"),
    ]

    sects = {"NIFTY 50 PREV CLOSE","NIFTY / GIFT NIFTY","FII / DII FLOWS","TRADE RULES"}
    for ri, (k, v) in enumerate(rows, 2):
        ck = ws.cell(row=ri, column=1, value=k)
        cv = ws.cell(row=ri, column=2, value=v)
        ws.row_dimensions[ri].height = 20
        if k in sects:
            ck.font = cv.font = fnt(bold=True, sz=10, col=C["wh"])
            ck.fill = cv.fill = F(C["md"])
            ws.cell(row=ri, column=3).fill = F(C["md"])
        elif k.strip():
            ck.font = fnt(bold=True, sz=10, col=C["md"])
            cv.font = fnt(sz=10)
            vu = str(v).upper()
            if "BUY" in vu:      cv.font = fnt(bold=True, sz=10, col=C["gd"])
            elif "SELL" in vu:   cv.font = fnt(bold=True, sz=10, col=C["rd"])
            elif v == "Bullish": cv.font = fnt(bold=True, sz=10, col=C["gd"])
            elif v == "Bearish": cv.font = fnt(bold=True, sz=10, col=C["rd"])
        ck.alignment = cv.alignment = aln("left")


def write_sector_sheet(ws, sdf):
    ws.merge_cells("A1:F1")
    ws["A1"] = "Sector Performance Summary"
    ws["A1"].font = fnt(bold=True, sz=14, col=C["wh"]); ws["A1"].fill = F(C["dk"])
    ws["A1"].alignment = aln("left", ind=2); ws.row_dimensions[1].height = 34
    for i, w in zip(range(1,7), [16,9,14,10,10,12]):
        ws.column_dimensions[get_column_letter(i)].width = w
    hrow(ws, 2, ["Sector","Stocks","Avg Change %","Gainers","Losers","Bias"], C["st"])
    for ri, row in sdf.iterrows():
        er   = ri + 3; bias = str(row.get("Bias","Neutral"))
        rbg  = C["gb"] if bias=="Bullish" else(C["rb"] if bias=="Bearish" else C["gr"])
        ws.row_dimensions[er].height = 19
        vals = [row["Sector"],row["Stocks"],row["Avg_Chg"],
                row["Gainers"],row["Losers"],bias]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=er, column=ci, value=val)
            c.fill=F(rbg); c.border=B(); c.font=fnt(sz=10); c.alignment=aln("center")
            if ci == 1:
                c.font=fnt(bold=True,sz=10,col=C["md"]); c.alignment=aln("left")
            elif ci == 3:
                c.number_format="0.00"
                c.font=fnt(bold=True,sz=10,col=(C["gd"] if val>0 else C["rd"]))
            elif ci == 6:
                col = C["gd"] if bias=="Bullish" else(C["rd"] if bias=="Bearish" else C["ad"])
                c.font=fnt(bold=True,sz=10,col=col)


def write_legend_sheet(ws):
    ws.merge_cells("A1:B1")
    ws["A1"] = "Legend & Trading Guide"
    ws["A1"].font = fnt(bold=True, sz=14, col=C["wh"]); ws["A1"].fill = F(C["dk"])
    ws["A1"].alignment = aln("left", ind=2); ws.row_dimensions[1].height = 34
    ws.column_dimensions["A"].width = 18; ws.column_dimensions["B"].width = 72
    guide = [
        ("COLUMN GUIDE",""),
        ("Symbol",        "NSE ticker symbol"),
        ("Sector",        "Industry sector classification"),
        ("Setup",         "Momentum Buy = trend continuation  |  Reversal = oversold bounce"),
        ("Trend",         "Bullish/Bearish/Neutral based on SMA10 vs SMA20"),
        ("Close Rs",      "Official closing price from NSE Bhav Copy"),
        ("Change %",      "% price change from previous close"),
        ("RSI (14)",      "Below 35 = oversold (green)  |  Above 65 = overbought (red)"),
        ("Vol Ratio",     "Today volume / 20-day avg volume  (>=1.5 = unusually high)"),
        ("52W High/Low",  "52-week highest and lowest prices"),
        ("% from High",   "How far below 52-week high (negative = below high)"),
        ("Entry Rs",      "0.3% above close — wait for 15-min candle confirmation"),
        ("Stop Loss Rs",  f"Momentum: {MOM_SL_PCT*100:.1f}% below entry  |  Reversal: {REV_SL_PCT*100:.1f}% below entry"),
        ("Target 1 Rs",  f"Momentum: {MOM_T1_PCT*100:.1f}% above entry  |  Reversal: {REV_T1_PCT*100:.1f}% above entry"),
        ("Target 2 Rs",  f"Momentum: {MOM_T2_PCT*100:.1f}% above entry  |  Reversal: {REV_T2_PCT*100:.1f}% above entry"),
        ("R:R",           "Risk:Reward ratio — only take trades with R:R >= 1.5x"),
        ("",""),
        ("TRADING RULES",""),
        ("Rule 1","Run this script 7–9 AM every trading morning"),
        ("Rule 2","Check Market Context sheet first — Gift Nifty + FII/DII"),
        ("Rule 3","Wait for first 15-min candle before entering any trade"),
        ("Rule 4","Skip all trades if Nifty opens more than 1% down"),
        ("Rule 5","Book 50% quantity at Target 1, trail SL to entry for balance"),
        ("Rule 6","Square off ALL positions before 3:15 PM IST — no exceptions"),
        ("Rule 7","Risk maximum 1–2% of total capital per trade"),
        ("Rule 8","Never average down on any intraday position"),
        ("",""),
        ("FII/DII SOURCES",""),
        ("1st","NSE API — fiidiiTradeReact (best but needs session cookie)"),
        ("2nd","Tickertape — api.tickertape.in/market-stats/fii-dii-activity"),
        ("3rd","Trendlyne — trendlyne.com/macro-data/fii-dii"),
        ("4th","Moneycontrol — moneycontrol.com/stocks/marketstats/fii_dii_activity"),
        ("5th","Economic Times — economictimes.indiatimes.com"),
        ("",""),
        ("DISCLAIMER","For educational purposes only. Not SEBI-registered investment advice."),
    ]
    sects = {"COLUMN GUIDE","TRADING RULES","FII/DII SOURCES","DISCLAIMER"}
    for i, (k, v) in enumerate(guide, 2):
        ck = ws.cell(row=i, column=1, value=k)
        cv = ws.cell(row=i, column=2, value=v)
        ws.row_dimensions[i].height = 18
        if k in sects:
            ck.font = cv.font = fnt(bold=True, sz=10, col=C["wh"])
            ck.fill = cv.fill = F(C["md"])
        elif k.startswith("Rule") or k in ("1st","2nd","3rd","4th","5th"):
            ck.font = fnt(bold=True, sz=10, col=C["bt"]); cv.font = fnt(sz=10)
        elif k:
            ck.font = fnt(bold=True, sz=10, col=C["md"]); cv.font = fnt(sz=10)
        else:
            ck.font = cv.font = fnt(sz=10)
        ck.alignment = cv.alignment = aln("left")


def build_excel(picks, sdf, fii_dii, gift, tday, nifty_close):
    wb  = Workbook()
    ws1 = wb.active;  ws1.title = "Intraday Picks"
    ws2 = wb.create_sheet("Market Context")
    ws3 = wb.create_sheet("Sector View")
    ws4 = wb.create_sheet("Legend & Guide")
    write_picks_sheet(ws1, picks, tday)
    write_context_sheet(ws2, tday, fii_dii, gift, nifty_close)
    write_sector_sheet(ws3, sdf)
    write_legend_sheet(ws4)
    wb.save(OUTPUT_FILE)
    print(f"\n  Saved: {OUTPUT_FILE}")



def send_telegram_document(file_path, caption=""):
    """Send the generated Excel file to Telegram using a bot."""
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    chat_id = os.getenv("TELEGRAM_CHAT_ID")

    if not token or not chat_id:
        print("  [Telegram] TELEGRAM_BOT_TOKEN or TELEGRAM_CHAT_ID is not set; skipping upload.")
        return False

    url = f"https://api.telegram.org/bot{token}/sendDocument"

    try:
        with open(file_path, "rb") as fh:
            resp = requests.post(
                url,
                data={"chat_id": chat_id, "caption": caption[:1024]},
                files={"document": (
                    os.path.basename(file_path),
                    fh,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )},
                timeout=60,
            )
        resp.raise_for_status()
        payload = resp.json()
        if payload.get("ok"):
            print("  [Telegram] Excel file sent successfully.")
            return True
        print(f"  [Telegram] API returned an error: {payload}")
    except Exception as e:
        print(f"  [Telegram] Upload failed: {e}")

    return False


# ══════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════
def main():
    tday = last_trading_day()
    print("=" * 60)
    print(f"  NSE Intraday Picks PRO  —  {date.today().strftime('%d %B %Y')}")
    print(f"  Closing prices from:  {tday.strftime('%d %B %Y')}")
    print("=" * 60)

    print("\n[1/6] Initialising NSE session …")
    boot_session()

    print("[2/6] Downloading NSE Bhav Copy (official prices) …")
    bhav = download_bhav_copy(tday)
    if bhav.empty:
        print("  Trying previous trading day …")
        bhav = download_bhav_copy(last_trading_day(tday))
    if bhav.empty:
        print("  Switching to yfinance universe fallback …")
        bhav = bhav_from_yfinance()
    if bhav.empty:
        print("  ERROR: No price data. Check internet connection.")
        sys.exit(1)
    print(f"  {len(bhav)} stocks loaded")

    nr = bhav[bhav["SYMBOL"].str.upper() == "NIFTY"]
    nifty_close = float(nr["CLOSE"].iloc[0]) if not nr.empty else None

    print("[3/6] Fetching Nifty / Gift Nifty direction …")
    gift = fetch_gift_nifty()
    gv = gift.get("value"); gc = gift.get("change"); gp = gift.get("change_pct")
    # FIX: Bhav Copy almost never contains a NIFTY index row, so
    # "NIFTY 50 PREV CLOSE" was showing N/A even though we already have
    # the Nifty level from the Gift Nifty fetch above — reuse it here.
    if nifty_close is None and gv:
        try:
            nifty_close = float(gv)
        except (TypeError, ValueError):
            nifty_close = None
    if gv:
        print(f"  Nifty: {gv:,.2f}  {gc:+.2f} ({gp:+.2f}%)  => {gift['direction']}"
              f"  [{gift['source']}]")
    else:
        print("  Nifty level unavailable — check nseindia.com")

    print("[4/6] Fetching FII / DII flows …")
    fii_dii = fetch_fii_dii()
    print(f"  FII: {fii_dii.get('fii_label','N/A')}   "
          f"DII: {fii_dii.get('dii_label','N/A')}")

    print("[5/6] Selecting candidates and computing technicals …")
    cands = select_candidates(bhav, top_n=25)
    print(f"  {len(cands)} candidates selected from Bhav Copy")
    tech = compute_technicals(cands)
    print(f"  Technicals computed for {len(tech)} symbols")

    print("[6/6] Building picks and sector summary …")
    picks = build_picks(bhav[bhav["SYMBOL"].isin(cands)].copy(), tech)
    sdf   = sector_summary(bhav)

    if picks.empty:
        print("  No picks passed the score filter today.")
        sys.exit(1)
    print(f"  {len(picks)} final picks selected")

    build_excel(picks, sdf, fii_dii, gift, tday, nifty_close)

    # Send the generated Excel report to Telegram when running on GitHub Actions
    # (or any machine where the two environment variables are configured).
    telegram_caption = (
        f"NSE Intraday Picks PRO — {date.today().strftime('%d %b %Y')}\\n"
        f"Final picks: {len(picks)}\\n"
        f"Source close: {tday.strftime('%d %b %Y')}\\n"
        f"Educational use only — not SEBI advice."
    )
    send_telegram_document(OUTPUT_FILE, telegram_caption)

    pcols = ["Symbol","Sector","Setup","Close Rs","Change %",
             "Entry Rs","Stop Loss Rs","Target 1 Rs","Target 2 Rs","R:R"]
    avail = [c for c in pcols if c in picks.columns]
    print("\n" + picks[avail].to_string(index=False))
    print(f"\nDone! Open  {OUTPUT_FILE}  on your Desktop.")

if __name__ == "__main__":
    main()
