"""
India wholesale power prices: IEX (Indian Energy Exchange) Day-Ahead Market, all-India market clearing
price. Found via discovery_archive/subcontinent/SEA_DISCOVERY_INDIA.py.

The IEX market-data pages load JSON (no key):
  https://www.iexindia.com/api/v1/dam/market-snapshot?interval=ONE_HOUR&fromDate=DD-MM-YYYY&toDate=DD-MM-YYYY
  -> per delivery hour: purchase_bid, sell_bid, mcv (market clearing volume, MW), final_scheduled_volume,
     mcp (market clearing price, Rs/MWh), weighted_mcp. History from at least 2022.

Writes output/Data and Chart Outputs/india_power_prices.xlsx:
  Daily   per delivery day: average MCP (time-weighted), volume-weighted MCP, max and min hourly MCP,
          cleared volume (MWh), purchase / sell bids (MWh) - Rs/MWh as published
  Hourly  the hourly rows, last 120 days

The DAM price cap is Rs 10,000/MWh (hours at the cap show the bid-ask shortfall in purchase_bid vs sell_bid).

Incremental: the Daily sheet is the history store; only delivery days not saved yet (plus the last
REVISION_DAYS) are fetched. Runs on the 1st and 15th.

    python3 asia/INDIA_IEX_PRICES.py
"""
import argparse
import os
import sys
import time
from datetime import date, timedelta

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

API = "https://www.iexindia.com/api/v1/dam/market-snapshot"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "application/json", "Referer": "https://www.iexindia.com/market-data/day-ahead-market/market-snapshot"}
T = (15, 120)
DATA_START = date(2022, 1, 1)
CHUNK_DAYS = 7
REVISION_DAYS = 16   # runs are 14-17 days apart: re-read everything since the last run (provisional days get final)
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "india_power_prices.xlsx")
NUM = ["purchase_bid", "sell_bid", "mcv", "final_scheduled_volume", "mcp", "weighted_mcp"]


def out(*a):
    print(*a, flush=True)


def fetch(d0, d1):
    for i in range(4):
        try:
            r = requests.get(API, params={"interval": "ONE_HOUR", "fromDate": d0.strftime("%d-%m-%Y"),
                                          "toDate": d1.strftime("%d-%m-%Y")}, headers=H, timeout=T)
            r.raise_for_status()
            df = pd.DataFrame(r.json().get("data") or [])
            if df.empty:
                return df
            df["date"] = pd.to_datetime(df["date"], format="%d-%m-%Y")
            df["hour"] = pd.to_numeric(df["hour"], errors="coerce")
            df = df[df["hour"].notna()].astype({"hour": int})
            df[NUM] = df[NUM].apply(pd.to_numeric, errors="coerce")
            return df[["date", "hour"] + NUM]
        except (requests.RequestException, ValueError, KeyError) as e:
            if i == 3:
                raise
            out(f"  retry {d0}: {e}")
            time.sleep(5 * (i + 1))


def daily_from(h):
    g = h.groupby("date")
    vwap = (h["mcp"] * h["mcv"]).groupby(h["date"]).sum() / g["mcv"].sum()
    return pd.DataFrame({"MCP_avg_Rs_per_MWh": g["mcp"].mean(), "MCP_volume_weighted_Rs_per_MWh": vwap,
                         "MCP_max_Rs_per_MWh": g["mcp"].max(), "MCP_min_Rs_per_MWh": g["mcp"].min(),
                         "Cleared_volume_MWh": g["mcv"].sum(), "Purchase_bid_MWh": g["purchase_bid"].sum(),
                         "Sell_bid_MWh": g["sell_bid"].sum(), "Hours": g.size()}).round(2)


def read_sheet(path, sheet, index_col=0):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=index_col)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    if index_col is not None:
        df.index = pd.to_datetime(df.index)
    return df


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start", default=DATA_START.isoformat())
    args = ap.parse_args()
    start = date.fromisoformat(args.start)
    old = read_sheet(args.out, "Daily")
    old_h = read_sheet(args.out, "Hourly", None)
    last = date.today()   # the DAM clears today for tomorrow; take delivery days up to today
    have = set(old.index.date) if not old.empty else set()
    revise = {last - timedelta(days=k) for k in range(REVISION_DAYS)}
    todo = [d for d in (start + timedelta(days=k) for k in range((last - start).days + 1)) if d not in have or d in revise]
    out(f"{len(have)} days saved; fetching {len(todo)}")
    frames, i = [], 0
    while i < len(todo):
        d0 = todo[i]
        d1 = min(d0 + timedelta(days=CHUNK_DAYS - 1), last)
        try:
            df = fetch(d0, d1)
        except Exception as e:  # noqa: BLE001  (IEX down: save what was fetched; the rest waits for the next run)
            out(f"  {d0}..{d1} failed after retries ({type(e).__name__}: {e}); saving what was fetched")
            break
        if not df.empty:
            frames.append(df)
        while i < len(todo) and todo[i] <= d1:
            i += 1
        if len(frames) % 20 == 0:
            out(f"  {d0}..{d1}")
        time.sleep(0.5)
    hourly = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=["date", "hour"] + NUM)
    hourly = hourly[hourly["date"].dt.date.isin(todo)].drop_duplicates(["date", "hour"], keep="last")
    new = daily_from(hourly) if not hourly.empty else pd.DataFrame()
    daily = new if old.empty else pd.concat([old[~old.index.isin(new.index)], new]).sort_index()
    if daily.empty:
        raise SystemExit("No IEX data")
    daily.index.name = "date"
    if not old_h.empty:
        old_h["date"] = pd.to_datetime(old_h["date"])
        hourly = pd.concat([old_h[~old_h["date"].isin(hourly["date"])], hourly], ignore_index=True)
    hourly = hourly[hourly["date"] >= hourly["date"].max() - pd.Timedelta(days=120)].sort_values(["date", "hour"])
    notes = [
        "UNITS",
        "Rs/MWh (Indian rupees per MWh) as published; volumes MWh (hourly MW x 1 h). Daily: MCP_avg = simple average "
        "of the 24 hourly market clearing prices; MCP_volume_weighted = weighted by cleared volume; max/min hourly MCP; "
        "Cleared_volume_MWh, Purchase_bid_MWh, Sell_bid_MWh = day totals. Hours = hourly rows returned (24 = complete).",
        "Hourly: the hourly rows for the last 120 days (date = delivery day, hour 1 = 00:00-01:00).",
        "",
        "COVERAGE",
        f"IEX Day-Ahead Market (DAM), all-India unconstrained MCP. Daily from {daily.index.min():%Y-%m-%d} to "
        f"{daily.index.max():%Y-%m-%d}. The DAM price cap is Rs 10,000/MWh. Regional area prices (after "
        "transmission congestion) are not included.",
        "",
        "SOURCE",
        "IEX (Indian Energy Exchange), Day-Ahead Market snapshot: "
        "https://www.iexindia.com/market-data/day-ahead-market/market-snapshot (JSON at /api/v1/dam/market-snapshot).",
    ]
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Hourly": hourly.set_index("date")}, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}: {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d}")
    out(daily.tail(3).to_string())


if __name__ == "__main__":
    main()
