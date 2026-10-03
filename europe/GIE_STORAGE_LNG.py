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
(storage) / 2019-01-01 (LNG) on the first run; the workbook is saved after every country, so a timeout loses nothing.

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
STORAGE_START = "2018-10-01"   # 8 water years: the current + previous + the 5 before it for the 5-year band
LNG_START = "2019-01-01"

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


def fill_eu(df):
    """AGSI+'s own EU aggregate only exists from mid-2024; before that the EU total is the sum of its member countries
    (the two agree to 0.001% where both exist). GB and UA are not EU members and are left out of the sum."""
    members = [c[:-4] for c in df if c.endswith("_TWh") and c[:-4] in STORAGE and c[:-4] not in ("GB", "UA")]
    if not members or "EU_TWh" not in df:
        return df
    df = df.copy()
    for suffix in ("TWh", "injection_GWhd", "withdrawal_GWhd"):
        cols = [f"{m}_{suffix}" for m in members if f"{m}_{suffix}" in df]
        if len(cols) == len(members):
            total = df[cols].sum(axis=1, min_count=len(cols))
            df[f"EU_{suffix}"] = df.get(f"EU_{suffix}", pd.Series(dtype=float)).combine_first(total)
    cap = sum(df[f"{m}_TWh"] / (df[f"{m}_full_pct"] / 100.0) for m in members)   # working gas volume, TWh
    df["EU_full_pct"] = df["EU_full_pct"].combine_first(df["EU_TWh"] / cap * 100.0)
    return df.round(3)


def write_storage(path, df):
    df = fill_eu(df)
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


def pull_storage(out_dir, key, deadline):
    path = os.path.join(out_dir, "eu_gas_storage_daily.xlsx")
    df = read_existing(path)
    for cc in ["EU"] + STORAGE:
        if time.time() > deadline:
            print("time budget reached; stopping (countries done so far are saved)")
            break
        cols = [f"{cc}_TWh", f"{cc}_full_pct", f"{cc}_injection_GWhd", f"{cc}_withdrawal_GWhd"]
        start = start_for(df, cols, STORAGE_START)
        if len(df) and f"{cc}_injection_GWhd" not in df:   # workbook predates the injection/withdrawal columns
            start = STORAGE_START
        params = {"continent": "eu"} if cc == "EU" else {"country": cc}
        t0 = time.time()
        try:
            rows = fetch("https://agsi.gie.eu/api", dict(params, **{"from": start, "to": date.today().isoformat()}), key)
        except Exception as e:  # noqa: BLE001
            print(f"  AGSI {cc}: FAILED {e}", flush=True)
            continue
        recs = {}
        for r in rows:
            d = pd.Timestamp(r["gasDayStart"])
            twh, full = num(r.get("gasInStorage")), num(r.get("full"))
            if twh == twh:   # not NaN
                recs[d] = {f"{cc}_TWh": twh, f"{cc}_full_pct": full,
                           f"{cc}_injection_GWhd": num(r.get("injection")), f"{cc}_withdrawal_GWhd": num(r.get("withdrawal"))}
        print(f"  AGSI {cc}: {len(recs)} days from {start} in {time.time() - t0:.0f}s", flush=True)
        if recs:
            new = pd.DataFrame.from_dict(recs, orient="index")
            df = merge(df, new).round(3)
            df.index.name = "date"
            write_storage(path, df)   # checkpoint
        time.sleep(0.3)
    if df.empty:
        print("no storage data")
        return
    print(f"saved eu_gas_storage_daily.xlsx: {len(df)} days x {df.shape[1]} columns")


def write_lng(path, df):
    lines = ["Europe LNG terminals (Gas Infrastructure Europe, ALSI)", "",
             "Source", "GIE ALSI (Aggregated LNG Storage Inventory), https://alsi.gie.eu/ - terminal operators' own daily reports, "
             "aggregated by GIE. Free API key (the same as AGSI+).",
             "", "Units and definitions",
             "<CC>_sendout_GWhd = LNG regasified and sent into the grid, GWh per gas day. <CC>_inventory_GWh = LNG in terminal tanks "
             "(GWh). EU = GIE's EU aggregate. Countries without terminals or with no ALSI data are absent.",
             f"Re-fetches the last {REVISION_DAYS} days each run; history from {LNG_START}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(df)} days to {df.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": df}, lines, {"Source", "Units and definitions", "Last pull"})


def pull_lng(out_dir, key, deadline):
    path = os.path.join(out_dir, "eu_lng_terminals_daily.xlsx")
    df = read_existing(path)
    for cc in ["EU"] + LNG:
        if time.time() > deadline:
            print("time budget reached; stopping (countries done so far are saved)")
            break
        cols = [f"{cc}_sendout_GWhd", f"{cc}_inventory_GWh"]
        start = start_for(df, cols, LNG_START)
        params = {"continent": "eu"} if cc == "EU" else {"country": cc}
        t0 = time.time()
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
        print(f"  ALSI {cc}: {len(recs)} days from {start} in {time.time() - t0:.0f}s", flush=True)
        if recs:
            df = merge(df, pd.DataFrame.from_dict(recs, orient="index")).round(3)
            df.index.name = "date"
            write_lng(path, df)   # checkpoint
        time.sleep(0.3)
    if df.empty:
        print("no LNG data")
        return
    print(f"saved eu_lng_terminals_daily.xlsx: {len(df)} days x {df.shape[1]} columns")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--what", default="storage,lng")
    ap.add_argument("--max-minutes", type=float, default=100.0)
    args = ap.parse_args()
    key = gie_key()
    os.makedirs(args.out_dir, exist_ok=True)
    what = [w.strip() for w in args.what.split(",")]
    deadline = time.time() + args.max_minutes * 60
    if "storage" in what:
        pull_storage(args.out_dir, key, deadline)
    if "lng" in what:
        pull_lng(args.out_dir, key, deadline)


if __name__ == "__main__":
    main()
