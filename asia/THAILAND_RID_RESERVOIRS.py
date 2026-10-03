"""
Thailand large-reservoir storage from the Royal Irrigation Department (RID) reservoir database,
https://app.rid.go.th/reservoir/. Found via discovery_archive/asia/SEA_DISCOVERY_TH_VN_PH.py and
SEA_DISCOVERY2.py.

API (no key): https://app.rid.go.th/reservoir/api/dam/public/<YYYY-MM-DD> -> the 35 large reservoirs on
that date, by region: capacity, normal storage, active / dead storage, current volume (million m3),
% of storage, inflow / outflow (million m3/day). The large dams include EGAT's hydropower reservoirs
(Bhumibol, Sirikit, Srinagarind, Vajiralongkorn, Ubol Ratana, ...); total storage drives Thai hydro
output and irrigation releases.

Writes output/Data and Chart Outputs/thailand_hydro_reservoirs.xlsx:
  Daily     total of the 35 large reservoirs: volume, normal storage, % full, usable (active) volume,
            inflow, outflow
  By dam    volume (million m3) per reservoir, last 400 days
  Water year chart (Oct-Sep) of total % full, via water_year_chart (add_charts.py)

Incremental: the Daily sheet is the history store; only dates not saved yet (from WY start 2020-10-01)
are fetched, plus the last REVISION_DAYS. Runs on the 1st and 15th.

    python3 asia/THAILAND_RID_RESERVOIRS.py
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

API = "https://app.rid.go.th/reservoir/api/dam/public/{d}"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept": "application/json"}
T = (15, 60)
DATA_START = date(2020, 10, 1)
REVISION_DAYS = 3
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "thailand_hydro_reservoirs.xlsx")


def out(*a):
    print(*a, flush=True)


def fetch(d):
    for i in range(4):
        try:
            r = requests.get(API.format(d=d.isoformat()), headers=H, timeout=T)
            r.raise_for_status()
            j = r.json()
            if j.get("date") != d.isoformat():
                return None   # the API falls back to today for dates it does not hold
            rows = []
            for region in j.get("data") or []:
                for dam in region.get("dam") or []:
                    rows.append({"id": dam.get("id"), "name": dam.get("name"), "region": region.get("region"),
                                 **{k: pd.to_numeric(dam.get(k), errors="coerce")
                                    for k in ("capacity", "storage", "active_storage", "dead_storage", "volume",
                                              "percent_storage", "inflow", "outflow")}})
            return pd.DataFrame(rows)
        except (requests.RequestException, ValueError) as e:
            if i == 3:
                out(f"  {d} failed: {e}")
                return None
            time.sleep(3 * (i + 1))


def read_sheet(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    return df.sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    old, old_dam = read_sheet(args.out, "Daily"), read_sheet(args.out, "By dam")
    today = date.today()
    have = set(old.index.date) if not old.empty else set()
    revise = {today - timedelta(days=k) for k in range(REVISION_DAYS)}
    todo = [d for d in (DATA_START + timedelta(days=k) for k in range((today - DATA_START).days + 1))
            if d not in have or d in revise]
    out(f"{len(have)} days saved; fetching {len(todo)}")
    tot, per_dam, names = {}, {}, {}
    for n, d in enumerate(todo):
        df = fetch(d)
        if df is not None and not df.empty:
            usable = (df["volume"] - df["dead_storage"]).clip(lower=0)
            tot[pd.Timestamp(d)] = {"Volume_mcm": df["volume"].sum(), "Normal_storage_mcm": df["storage"].sum(),
                                    "Pct_full": round(100 * df["volume"].sum() / df["storage"].sum(), 2),
                                    "Usable_volume_mcm": usable.sum(),
                                    "Pct_usable": round(100 * usable.sum() / df["active_storage"].sum(), 2),
                                    "Inflow_mcm_per_day": df["inflow"].sum(), "Outflow_mcm_per_day": df["outflow"].sum(),
                                    "Reservoirs": len(df)}
            per_dam[pd.Timestamp(d)] = df.set_index("id")["volume"]
            names.update(df.set_index("id")["name"].to_dict())
        if n % 100 == 0:
            out(f"  {d}: {len(tot)} days fetched")
        time.sleep(0.2)
    new = pd.DataFrame.from_dict(tot, orient="index").round(2)
    daily = new if old.empty else (old if new.empty else
                                   pd.concat([old[~old.index.isin(new.index)], new]).sort_index())
    if daily.empty:
        raise SystemExit("No RID data")
    dam = pd.DataFrame(per_dam).T.rename(columns=lambda c: names.get(c, c)) if per_dam else pd.DataFrame()
    dam = dam if old_dam.empty else (old_dam if dam.empty else
                                     pd.concat([old_dam[~old_dam.index.isin(dam.index)], dam]).sort_index())
    dam = dam[dam.index >= dam.index.max() - pd.Timedelta(days=400)]
    daily.index.name = dam.index.name = "date"
    notes = [
        "UNITS",
        "Million m3 (mcm). Daily: totals of the RID large reservoirs - Volume_mcm (water in storage), "
        "Normal_storage_mcm (storage at normal high water level), Pct_full = Volume / Normal storage; Usable_volume_mcm "
        "= volume above dead storage, Pct_usable = usable / active storage; Inflow / Outflow mcm per day; Reservoirs = "
        "number reporting that day (35).",
        "By dam: volume per reservoir (names in Thai as published), last 400 days.",
        "",
        "COVERAGE",
        f"Daily from {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}. The 35 large reservoirs (RID and "
        "EGAT), which hold most of Thailand's stored water, including the hydropower reservoirs.",
        "",
        "SOURCE",
        "Royal Irrigation Department (RID), reservoir database: https://app.rid.go.th/reservoir/ "
        "(API https://app.rid.go.th/reservoir/api/dam/public/<date>).",
    ]
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "By dam": dam}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}: {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d}")
    out(daily.tail(3).to_string())


if __name__ == "__main__":
    main()
