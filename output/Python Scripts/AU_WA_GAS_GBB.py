"""
Western Australia gas from AEMO's WA Gas Bulletin Board (public API, no key), one report per gas day:

  https://gbbwa.aemo.com.au/api/v1/report/actualFlow/YYYY-MM-DD          production / pipeline / storage flows, TJ
  https://gbbwa.aemo.com.au/api/v1/report/endUserConsumption/YYYY-MM-DD  consumption by zone: large users,
                                                                         distribution, other, TJ

Writes au_wa_gas.xlsx:
  Production    TJ/day by production facility (receipts of facilityType 'production'), Total
  Storage       TJ/day storage flows (receipt = withdrawal from storage, delivery = injection) by facility
  Consumption   TJ/day by end-user type (large users, distribution, other), summed over WA zones, Total
  By zone       TJ/day total consumption by zone

WA's LNG plants (North West Shelf, Gorgon, Wheatstone, Pluto) export directly and are not on the WA GBB except
their domestic gas plants. Incremental: fetches gas days not saved yet plus the last 7 saved (AEMO revises
recent days); a new file backfills from 2021. Saves every 90 days fetched.

Usage: python3 AU_WA_GAS_GBB.py [--out "output/Data and Chart Outputs/au_wa_gas.xlsx"]
"""
import argparse
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import date

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

API = "https://gbbwa.aemo.com.au/api/v1/report/{report}/{day}"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0",
           "Accept": "application/json"}
HISTORY_START = "2021-01-01"
REFRESH_DAYS = 7
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "au_wa_gas.xlsx")
SHEETS = ["Production", "Storage", "Consumption", "By zone"]


def get_json(report, day):
    for attempt in range(4):
        try:
            r = requests.get(API.format(report=report, day=day), headers=HEADERS, timeout=(10, 60))
            if r.status_code == 404:
                return None
            r.raise_for_status()
            return r.json()
        except (requests.RequestException, ValueError) as e:
            if attempt == 3:
                print(f"    {report} {day}: {type(e).__name__}: {str(e)[:120]}", flush=True)
                return None
            time.sleep(3 * (attempt + 1))


def fetch_day(day):
    """One gas day -> dict of single-row Series per sheet."""
    ds = day.strftime("%Y-%m-%d")
    out = {}
    flow = get_json("actualFlow", ds)
    if flow and flow.get("rows"):
        f = pd.DataFrame(flow["rows"])
        for c in ("receipt", "delivery"):
            f[c] = pd.to_numeric(f.get(c), errors="coerce")
        typ = f["facilityType"].astype(str).str.lower()
        prod = f[typ.eq("production")].groupby("facilityName")["receipt"].sum(min_count=1)
        out["Production"] = prod
        st = f[typ.str.contains("storage")]
        if not st.empty:
            out["Storage"] = pd.concat([st.groupby("facilityName")["receipt"].sum(min_count=1).add_suffix(" withdrawal"),
                                        st.groupby("facilityName")["delivery"].sum(min_count=1).add_suffix(" injection")])
    use = get_json("endUserConsumption", ds)
    if use and use.get("rows"):
        u = pd.DataFrame(use["rows"])
        for c in ("largeUser", "distribution", "other", "total"):
            u[c] = pd.to_numeric(u.get(c), errors="coerce")
        out["Consumption"] = pd.Series({"Large users": u["largeUser"].sum(min_count=1),
                                        "Distribution (residential, commercial)": u["distribution"].sum(min_count=1),
                                        "Other": u["other"].sum(min_count=1)})
        out["By zone"] = u.set_index("zoneName")["total"]
    return day, out


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    d.index = pd.to_datetime(pd.Index(d.index).astype(str), errors="coerce")
    return d[d.index.notna()].sort_index()


def save(path, frames):
    sheets = {}
    for k in SHEETS:
        f = frames.get(k, pd.DataFrame())
        if f.empty:
            continue
        f = f.drop(columns=["Total"], errors="ignore").sort_index()
        if k in ("Production", "Consumption"):
            f["Total"] = f.sum(axis=1, min_count=1)
        f.index.name = "date"
        sheets[k] = f.round(2)
    p = sheets.get("Production", pd.DataFrame())
    notes = [
        "UNITS",
        "Terajoules per gas day (TJ/day), as reported by facility operators to AEMO's WA Gas Bulletin Board.",
        "Production: receipts into the pipeline network from each WA domestic gas production facility (incl. the "
        "domestic gas plants of the LNG projects, e.g. Gorgon, Wheatstone, Varanus, Devil Creek). Storage: withdrawals "
        "(receipt) and injections (delivery) at Mondarra / Tubridgi. Consumption: end-user consumption summed over "
        "the WA zones - large users, distribution networks (residential, commercial), other. By zone: total by zone.",
        "",
        "COVERAGE",
        (f"Western Australia, daily, {p.index.min():%d %b %Y} to {p.index.max():%d %b %Y}. " if not p.empty else "") +
        "LNG export volumes are not on the WA GBB.",
        "",
        "SOURCE",
        "AEMO WA Gas Bulletin Board API: https://gbbwa.aemo.com.au/api/v1/report/actualFlow/<gas day> and "
        "/endUserConsumption/<gas day>",
        "https://gbbwa.aemo.com.au/",
    ]
    xlsx_notes.write_workbook(path, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()

    frames = {k: load(args.out, k) for k in SHEETS}
    have = set(frames["Production"].index) if not frames["Production"].empty else set()
    days = pd.date_range(HISTORY_START, pd.Timestamp(date.today()) - pd.Timedelta(days=2))
    todo = [d for d in days if d not in have] + sorted(have)[-REFRESH_DAYS:]
    todo = sorted(set(todo))
    print(f"{len(have)} gas days saved; fetching {len(todo)}", flush=True)
    for i in range(0, len(todo), 90):
        chunk = todo[i:i + 90]
        with ThreadPoolExecutor(max_workers=6) as ex:
            results = list(ex.map(fetch_day, chunk))
        for k in SHEETS:
            rows = {d: r[k] for d, r in results if k in r}
            if not rows:
                continue
            new = pd.DataFrame(rows).T
            old = frames[k]
            old = old[~old.index.isin(new.index)] if not old.empty else old
            frames[k] = pd.concat([old, new]).sort_index()
        got = sum(1 for _, r in results if r)
        print(f"  {chunk[0]:%Y-%m-%d}..{chunk[-1]:%Y-%m-%d}: {got}/{len(chunk)} days", flush=True)
        save(args.out, frames)
    if frames["Production"].empty:
        raise SystemExit("No WA GBB data")
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
