"""
Europe gas storage and LNG terminals from Gas Infrastructure Europe: AGSI+ (storage) and ALSI (LNG), two workbooks:

  output/Data and Chart Outputs/eu_gas_storage_daily.xlsx   sheet "Daily": date, <CC>_TWh (gas in storage),
                                                            <CC>_full_pct (% of working gas volume), <CC>_injection_GWhd
                                                            and <CC>_withdrawal_GWhd per country + EU
  output/Data and Chart Outputs/eu_lng_terminals_daily.xlsx sheet "Daily": date, <CC>_sendout_GWhd (LNG send-out into
                                                            the grid, GWh per day) and EU_inventory_GWh

Gas days (not calendar-adjusted). Storage charts are AGSI-style water-year charts (Oct-Sep); LNG send-out is charted
monthly by country (add_charts.py REGISTRY).

Incremental: reads the committed workbook and re-fetches the last 14 days (GIE revises) plus any gap, from 2012-01-01
(storage) / 2018-01-01 (LNG) on the first run. A first backfill is ~300 requests of up to 400 days each.

Usage: python3 GIE_STORAGE_LNG.py [--out-dir DIR] [--what storage,lng]
Requires GIE_API_KEY (environment / GitHub secret, or api_keys.py): one free key serves AGSI and ALSI.
"""
import argparse
import os
import sys
import time
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

try:
    import api_keys  # type: ignore  # noqa: F401
except Exception:  # noqa: BLE001
    api_keys = None

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
REVISION_DAYS = 14
STORAGE_START = "2012-01-01"
LNG_START = "2018-01-01"

# AGSI+ countries with storage (EU members + GB + UA); "EU" is the continent aggregate
STORAGE = ["AT", "BE", "BG", "HR", "CZ", "DK", "FR", "DE", "HU", "IE", "IT", "LV", "NL", "PL", "PT", "RO", "SK", "ES",
           "SE", "GB", "UA"]
# ALSI countries with LNG terminals
LNG = ["BE", "FR", "DE", "GR", "HR", "IT", "LT", "NL", "PL", "PT", "ES", "FI", "EE", "GB"]
NAMES = {"AT": "Austria", "BE": "Belgium", "BG": "Bulgaria", "HR": "Croatia", "CZ": "Czechia", "DK": "Denmark",
         "FR": "France", "DE": "Germany", "HU": "Hungary", "IE": "Ireland", "IT": "Italy", "LV": "Latvia",
         "NL": "Netherlands", "PL": "Poland", "PT": "Portugal", "RO": "Romania", "SK": "Slovakia", "ES": "Spain",
         "SE": "Sweden", "GB": "United Kingdom", "UA": "Ukraine", "GR": "Greece", "LT": "Lithuania", "FI": "Finland",
         "EE": "Estonia", "EU": "EU"}


def gie_key():
    k = os.environ.get("GIE_API_KEY") or (getattr(api_keys, "GIE_API_KEY", "") if api_keys else "")
    if not k:
        raise SystemExit("GIE_API_KEY is not set (GitHub secret, or GIE_API_KEY in api_keys.py). "
                         "Register at https://agsi.gie.eu/#/registration - the same key works for AGSI and ALSI.")
    return k


def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return float("nan")


def fetch(base, params, key, tries=4):
    """All pages of one query -> list of row dicts. Retries 429/5xx."""
    rows, page = [], 1
    while True:
        for attempt in range(tries):
            try:
                r = requests.get(base, params=dict(params, page=page, size=400), headers=dict(UA, **{"x-key": key}), timeout=(10, 120))
            except requests.RequestException as e:
                last = f"{type(e).__name__}"
                time.sleep(5 * (attempt + 1))
                continue
            if r.status_code in (429,) or r.status_code >= 500:
                last = f"HTTP {r.status_code}"
                time.sleep(30 * (attempt + 1))
                continue
            break
        else:
            raise RuntimeError(f"{base} {params}: failed after {tries} tries ({last})")
        if r.status_code != 200:
            raise RuntimeError(f"{base} {params}: HTTP {r.status_code} {r.text[:150]}")
        j = r.json()
        rows += j.get("data", [])
        if page >= int(j.get("last_page", 1) or 1):
            return rows
        page += 1
        time.sleep(0.3)


def read_existing(path):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        d = pd.read_excel(path, sheet_name="Daily")
    except Exception as e:  # noqa: BLE001
        print(f"could not read {os.path.basename(path)} ({type(e).__name__}: {e}); starting over")
        return pd.DataFrame()
    d["date"] = pd.to_datetime(d["date"], errors="coerce")
    return d.dropna(subset=["date"]).set_index("date").sort_index()


def merge(old, new):
    """New values replace old for the same date and column; old stays where new has nothing."""
    if old.empty:
        return new.sort_index()
    idx = old.index.union(new.index)
    out = old.reindex(idx)
    for c in new.columns:
        if c not in out:
            out[c] = float("nan")
        s = new[c].dropna()
        out.loc[s.index, c] = s
    return out.sort_index()


def start_for(old, cols, default):
    if old.empty:
        return default
    last = max((old[c].dropna().index.max() for c in cols if c in old and old[c].notna().any()), default=None)
    if last is None:
        return default
    return max(datetime.strptime(default, "%Y-%m-%d").date(), (last - timedelta(days=REVISION_DAYS)).date()).isoformat()


def pull_storage(out_dir, key):
    path = os.path.join(out_dir, "eu_gas_storage_daily.xlsx")
    old = read_existing(path)
    frames, status = {}, []
    for cc in ["EU"] + STORAGE:
        cols = [f"{cc}_TWh", f"{cc}_full_pct", f"{cc}_injection_GWhd", f"{cc}_withdrawal_GWhd"]
        start = start_for(old, cols, STORAGE_START)
        if len(old) and f"{cc}_injection_GWhd" not in old:   # workbook predates the injection/withdrawal columns
            start = STORAGE_START
        params = {"continent": "eu"} if cc == "EU" else {"country": cc}
        try:
            rows = fetch("https://agsi.gie.eu/api", dict(params, **{"from": start, "to": date.today().isoformat()}), key)
        except Exception as e:  # noqa: BLE001
            status.append((cc, f"FAILED {e}"))
            print(f"  {cc}: FAILED {e}", flush=True)
            continue
        recs = {}
        for r in rows:
            d = pd.Timestamp(r["gasDayStart"])
            twh, full = num(r.get("gasInStorage")), num(r.get("full"))
            if twh == twh:   # not NaN
                recs[d] = {f"{cc}_TWh": twh, f"{cc}_full_pct": full,
                           f"{cc}_injection_GWhd": num(r.get("injection")), f"{cc}_withdrawal_GWhd": num(r.get("withdrawal"))}
        if recs:
            frames[cc] = pd.DataFrame.from_dict(recs, orient="index")
        status.append((cc, f"{len(recs)} days from {start}"))
        print(f"  AGSI {cc}: {len(recs)} days from {start}", flush=True)
        time.sleep(0.3)
    new = pd.concat(frames.values(), axis=1) if frames else pd.DataFrame()
    new.index.name = "date"
    df = merge(old, new)
    if df.empty:
        print("no storage data")
        return
    df.index.name = "date"
    df = df.round(3)
    lines = ["Europe gas storage (Gas Infrastructure Europe, AGSI+)", "",
             "Source", "GIE AGSI+ (Aggregated Gas Storage Inventory), https://agsi.gie.eu/ - operators' own daily storage reports, "
             "aggregated by GIE. Free API key.",
             "", "Units and definitions",
             "<CC>_TWh = gas in storage (TWh, gas day). <CC>_full_pct = % of working gas volume. <CC>_injection_GWhd and "
             "<CC>_withdrawal_GWhd = gas injected into / withdrawn from storage (GWh per gas day). EU = GIE's EU aggregate. "
             "GB (United Kingdom) and UA (Ukraine) are included where AGSI reports them.",
             "Charts are AGSI-style water-year charts (Oct-Sep): shaded 5-year min-max band, 5-year average, previous and "
             "current water year.",
             f"Re-fetches the last {REVISION_DAYS} days each run (GIE revises); history from {STORAGE_START}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(df)} days to {df.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": df}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved eu_gas_storage_daily.xlsx: {len(df)} days x {df.shape[1]} columns")


def pull_lng(out_dir, key):
    path = os.path.join(out_dir, "eu_lng_terminals_daily.xlsx")
    old = read_existing(path)
    frames = {}
    for cc in ["EU"] + LNG:
        cols = [f"{cc}_sendout_GWhd", f"{cc}_inventory_GWh"]
        start = start_for(old, cols, LNG_START)
        params = {"continent": "eu"} if cc == "EU" else {"country": cc}
        try:
            rows = fetch("https://alsi.gie.eu/api", dict(params, **{"from": start, "to": date.today().isoformat()}), key)
        except Exception as e:  # noqa: BLE001
            print(f"  ALSI {cc}: FAILED {e}", flush=True)
            continue
        recs = {}
        for r in rows:
            d = pd.Timestamp(r["gasDayStart"])
            so = num(r.get("sendOut"))
            inv = r.get("inventory")
            inv_gwh = num(inv.get("gwh")) if isinstance(inv, dict) else float("nan")
            if so == so:
                recs[d] = {f"{cc}_sendout_GWhd": so, f"{cc}_inventory_GWh": inv_gwh}
        if recs:
            frames[cc] = pd.DataFrame.from_dict(recs, orient="index")
        print(f"  ALSI {cc}: {len(recs)} days from {start}", flush=True)
        time.sleep(0.3)
    new = pd.concat(frames.values(), axis=1) if frames else pd.DataFrame()
    new.index.name = "date"
    df = merge(old, new)
    if df.empty:
        print("no LNG data")
        return
    df.index.name = "date"
    df = df.round(3)
    # inventory is only charted for the EU total; keep the per-country send-out plus every inventory column
    lines = ["Europe LNG terminals (Gas Infrastructure Europe, ALSI)", "",
             "Source", "GIE ALSI (Aggregated LNG Storage Inventory), https://alsi.gie.eu/ - terminal operators' own daily reports, "
             "aggregated by GIE. Free API key (the same as AGSI+).",
             "", "Units and definitions",
             "<CC>_sendout_GWhd = LNG regasified and sent into the grid, GWh per gas day. <CC>_inventory_GWh = LNG in terminal tanks "
             "(GWh). EU = GIE's EU aggregate. Countries without terminals or with no ALSI data are absent.",
             f"Re-fetches the last {REVISION_DAYS} days each run; history from {LNG_START}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(df)} days to {df.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": df}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved eu_lng_terminals_daily.xlsx: {len(df)} days x {df.shape[1]} columns")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--what", default="storage,lng")
    args = ap.parse_args()
    key = gie_key()
    os.makedirs(args.out_dir, exist_ok=True)
    what = [w.strip() for w in args.what.split(",")]
    if "storage" in what:
        pull_storage(args.out_dir, key)
    if "lng" in what:
        pull_lng(args.out_dir, key)


if __name__ == "__main__":
    main()
