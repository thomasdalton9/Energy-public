"""
Great Britain gas storage by site from National Gas Transmission's Data Portal ("Find gas data"), one workbook:

  output/Data and Chart Outputs/gb_storage_sites_daily.xlsx
    sheet "Daily": date (gas day), GWh. Per site (Aldbrough, Hatfield Moor, Hill Top, Holehouse Farm, Holford, Hornsea,
        Humbly Grove, Rough, Stublach): <site>_stock (the portal's 'Opening Stock'), <site>_inflow, <site>_outflow;
        totals: total_stock (sum of sites), total_inflow, total_outflow, portal_total_stock ('Storage, Medium and Long Range,
        Stock Levels').
    sheet "Units": source and definitions

The data items are medium- and long-range storage (the LNG terminals' stock items are ignored). The portal's item ids are looked
up from its catalogue (GET /api/find-gas-data-folders) by item name on every run, so a renumbering on National Gas's side does
not break the pull. Values are kWh per gas day in the source, converted to GWh.

Incremental: the committed workbook is the history store; each run re-fetches the last 14 days plus any gap.

Usage: python3 GB_STORAGE_SITES_DAILY.py [--out-dir DIR] [--start 2021-01-01]
"""
import argparse
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone

import numpy as np
import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "gb_storage_sites_daily.xlsx"
BASE = "https://data.nationalgas.com"
REVISION_DAYS = 14
WINDOW_DAYS = 60
SITES = ["Aldbrough", "Hatfield Moor", "Hill Top", "Holehouse Farm", "Holford", "Hornsea", "Humbly Grove", "Rough", "Stublach"]
KIND_COL = {"Opening Stock": "stock", "Inflow": "inflow", "Outflow": "outflow"}
TOTAL_ITEM = "Storage, Medium and Long Range, Stock Levels"
SITE_COLS = [f"{s}_{k}" for s in SITES for k in ("stock", "inflow", "outflow")]
COLUMNS = SITE_COLS + ["total_stock", "total_inflow", "total_outflow", "portal_total_stock"]
NAME_RE = re.compile(r"^(Opening Stock|Inflow|Outflow), (.+?), (?:Medium|Long) Range Storage$")
HEADERS = {"Content-Type": "application/json", "Accept": "application/json, text/plain, */*",
           "Referer": "https://data.nationalgas.com/find-gas-data/view",
           "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"}


def catalogue_ids():
    txt = None
    for i in range(4):
        try:
            r = requests.get(BASE + "/api/find-gas-data-folders", headers=HEADERS, timeout=90)
            if r.ok:
                txt = r.text
                break
        except requests.RequestException:
            pass
        time.sleep(5 * (i + 1))
    if txt is None:
        raise RuntimeError("could not read the portal catalogue")
    ids = {}
    for m in re.finditer(r'"name":\s*"([^"]*)",\s*"description":\s*"((?:PUBOB?J?\d+)[^"]*)"', txt):
        name = m.group(1)
        pid = re.match(r"(PUBOB?J?\d+)", m.group(2)).group(1)
        mm = NAME_RE.match(name)
        if (mm and mm.group(2) in SITES) or name == TOTAL_ITEM:
            ids[pid] = name
    print(f"{len(ids)} storage items found in the catalogue")
    missing = [f"{k}, {s}" for s in SITES for k in KIND_COL if not any(n.startswith(f"{k}, {s},") for n in ids.values())]
    if missing:
        print("not in the catalogue:", missing)
    return list(ids)


def fetch(ids, d0, d1, tries=5):
    body = {"latestFlag": "Y", "applicableFor": "Y", "dateFrom": d0.isoformat(), "dateTo": d1.isoformat(),
            "dateType": "GASDAY", "ids": ",".join(ids)}
    for i in range(tries):
        try:
            r = requests.post(BASE + "/api/find-gas-data", json=body, headers=HEADERS, timeout=180)
            if r.ok:
                return r.json().get("data", [])
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(5 * (i + 1))
                continue
            raise RuntimeError(f"HTTP {r.status_code}: {r.text[:300]}")
        except requests.RequestException:
            time.sleep(5 * (i + 1))
    raise RuntimeError(f"gave up on {d0}..{d1}")


def to_frame(items):
    recs = []
    for it in items:
        name = it.get("itemName", "")
        if name == TOTAL_ITEM:
            col, div = "portal_total_stock", 1.0   # already GWh
        else:
            m = NAME_RE.match(name)
            if not m or m.group(2) not in SITES:
                continue
            col, div = f"{m.group(2)}_{KIND_COL[m.group(1)]}", 1e6   # kWh -> GWh
        try:
            recs.append((pd.to_datetime(it["applicableFor"], format="%d/%m/%Y"), col, float(it["value"]) / div))
        except (KeyError, ValueError, TypeError):
            continue
    if not recs:
        return pd.DataFrame(columns=COLUMNS)
    d = pd.DataFrame(recs, columns=["date", "col", "gwh"])
    out = d.pivot_table(index="date", columns="col", values="gwh", aggfunc="last").sort_index()
    return add_totals(out)


def add_totals(df):
    for c in COLUMNS:
        if c not in df:
            df[c] = np.nan
    for k in ("stock", "inflow", "outflow"):
        cols = [f"{s}_{k}" for s in SITES]
        df[f"total_{k}"] = df[cols].sum(axis=1, min_count=len(SITES))   # only when every site reported that day
    return df[COLUMNS]


def read_existing(path):
    if not os.path.exists(path):
        return pd.DataFrame(columns=COLUMNS)
    try:
        d = pd.read_excel(path, sheet_name="Daily")
    except Exception as e:  # noqa: BLE001
        print(f"could not read {FILE} ({type(e).__name__}: {e}); starting over")
        return pd.DataFrame(columns=COLUMNS)
    d["date"] = pd.to_datetime(d["date"], errors="coerce")
    d = d.dropna(subset=["date"]).set_index("date").sort_index()
    for c in COLUMNS:
        if c not in d:
            d[c] = np.nan
    return d[COLUMNS]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--start", default="2021-01-01")
    args = ap.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)
    path = os.path.join(args.out_dir, FILE)
    start = datetime.strptime(args.start, "%Y-%m-%d").date()
    today = datetime.now(timezone.utc).date()
    old = read_existing(path)
    have = set(old.dropna(how="all").index.date) if len(old) else set()
    fs = start
    if have:
        fs = max(start, max(have) - timedelta(days=REVISION_DAYS))
        gaps = [start + timedelta(days=i) for i in range((today - start).days)]
        gaps = [g for g in gaps if g >= today - timedelta(days=120) and g not in have and g < fs]
        if gaps:
            fs = min(fs, gaps[0])
    ids = catalogue_ids()
    parts, d0 = [], fs
    while d0 <= today:
        d1 = min(d0 + timedelta(days=WINDOW_DAYS - 1), today)
        items = fetch(ids, d0, d1)
        f = to_frame(items)
        print(f"  {d0} .. {d1}: {len(items)} items, {len(f)} days", flush=True)
        parts.append(f)
        d0 = d1 + timedelta(days=1)
        time.sleep(0.5)
    new = pd.concat(parts) if parts else pd.DataFrame(columns=COLUMNS)
    combined = old[~old.index.isin(new.index)] if len(old) else old
    combined = pd.concat([combined, new]) if len(combined) else new
    combined = combined.sort_index().astype(float).round(3).dropna(how="all")
    combined.index.name = "date"
    if combined.empty:
        print("no GB storage data")
        return
    print("days per column:", combined.notna().sum().to_dict())
    lines = ["Great Britain - gas storage by site (National Gas Transmission Data Portal)", "",
             "Source", "National Gas Transmission Data Portal, Find gas data: https://data.nationalgas.com/find-gas-data . "
             "Free, no key. Operator's own data (storage operators' daily submissions).",
             "", "Units and definitions",
             "GWh per gas day (the source is kWh). Per site: <site>_stock = the portal's 'Opening Stock', <site>_inflow = injection, "
             "<site>_outflow = withdrawal. Sites: " + ", ".join(SITES) + " (Rough is the long-range site, the others medium range). "
             "total_stock/inflow/outflow = sum of the nine sites (only on days every site reported); portal_total_stock = the portal's "
             "'Storage, Medium and Long Range, Stock Levels'. LNG terminal stocks are not included.",
             f"Re-fetches the last {REVISION_DAYS} days each run plus gaps within 120 days; history from {args.start}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(combined)} days, "
             f"{combined.index.min():%Y-%m-%d} to {combined.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": combined}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}: {len(combined)} days x {combined.shape[1]} columns")


if __name__ == "__main__":
    main()
