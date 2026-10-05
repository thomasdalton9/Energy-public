"""
Great Britain gas supply and demand from National Gas Transmission's NTS data (Data Portal, "Find gas data"), one workbook:

  output/Data and Chart Outputs/gb_gas_nts_daily.xlsx
    sheet "Daily": date (gas day), GWh per day:
        demand : ldz_offtake, powerstations, industrial_offtake, interconnector_exports, storage_injection
        supply : bacton_ukcs, barrow, easington, st_fergus, teesside (terminal nominations; St Fergus and Easington mix
                 UKCS with Norwegian gas), storage_withdrawal, interconnector_iuk, interconnector_bbl (imports from
                 Belgium / the Netherlands)
    sheet "Units": source and definitions

LNG terminals (Isle of Grain, South Hook, Dragon) are not in these items; GIE ALSI covers them. Demand items are the NTS
"Energy Offtaken ... Total" actuals (Demand > Exit Point Actuals); supply items are "Nominations, Day Ahead Net Aggregate"
per entry terminal plus the storage-withdrawal total and the two interconnector entry points. Values are kWh per gas day
in the source, converted to GWh.

Uses the endpoint the Find gas data page itself calls (POST /api/find-gas-data). Items are matched by the item name in the
response; names that are not recognised are printed so a change on National Gas's side shows up in the log.

Incremental: the committed workbook is the history store; each run re-fetches the last 14 days plus any gap.
A one-off bad value (a spike or drop-out to near zero against the local median) is blanked and logged.

Usage: python3 GB_GAS_NTS_DAILY.py [--out-dir DIR] [--start 2021-01-01]
"""
import argparse
import os
import sys
import time
from datetime import date, datetime, timedelta, timezone

import numpy as np
import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "gb_gas_nts_daily.xlsx"
URL = "https://data.nationalgas.com/api/find-gas-data"
REVISION_DAYS = 14
WINDOW_DAYS = 60

# publication ids as listed in the portal's data item catalogue
DEMAND_IDS = ["PUBOBJ1023", "PUBOBJ1024", "PUBOBJ1025", "PUBOBJ1026", "PUBOBJ1028"]
SUPPLY_IDS = ["PUBOB3876", "PUBOB16241", "PUBOB3877", "PUBOB3881", "PUBOB3891"]
FLOW_IDS = ["PUBOB4423", "PUBOB4424", "PUBOB570"]
# System entry points that ENTSOG's UK entry list lacks or double counts (System Entry Energy, D+2 allocations, kWh): the Rough sub-terminal
# (gas from the Rough field entering at Easington, 1.0-2.0 TWh a month, in no ENTSOG row), the small beach / onshore entries, and the Rough STORAGE
# entry, which ENTSOG's Easington entry already includes until Sep 2025 (Easington - Langeled - Dimlington = Rough storage withdrawals to 0.01 TWh).
ENTRY_IDS = ["PUBOBJ11191", "PUBOB424", "PUBOBJ2251", "PUBOB19002", "PUBOBJ2854", "PUBOB394"]
ENTRY_NAMES = {"systementryenergy,roughsubterminal,d+2": "entry_rough_subterminal", "systementryenergy,rough,d+2": "entry_rough_storage",
               "systementryenergy,saltfleetby,d+2": "entry_saltfleetby", "systementryenergy,murrow,d+2": "entry_murrow",
               "systementryenergy,glenthambiomethane,d+2": "entry_glentham", "systementryenergy,burtonpoint,d+2": "entry_burton_point"}
ENTRY_COLUMNS = list(ENTRY_NAMES.values())

NAME_TO_COLUMN = {
    "NTS Energy Offtaken, LDZ Offtake Total": "ldz_offtake",
    "NTS Energy Offtaken, Powerstations Total": "powerstations",
    "NTS Energy Offtaken, Industrial Offtake Total": "industrial_offtake",
    "NTS Energy Offtaken, Interconnector Exports Total": "interconnector_exports",
    "NTS Energy Offtaken, Storage Injection Total": "storage_injection",
    "Nominations, Day Ahead Net Aggregate, Bacton UKCS": "bacton_ukcs",
    "Nominations, Day Ahead Net Aggregate, Barrow": "barrow",
    "Nominations, Day Ahead Net Aggregate, Easington": "easington",
    "Nominations, Day Ahead Net Aggregate, StFergusTer": "st_fergus",
    "Nominations, Day Ahead Net Aggregate, TeesideTer": "teesside",
}
COLUMNS = ["ldz_offtake", "powerstations", "industrial_offtake", "interconnector_exports", "storage_injection",
           "bacton_ukcs", "barrow", "easington", "st_fergus", "teesside", "storage_withdrawal", "interconnector_iuk",
           "interconnector_bbl"] + ENTRY_COLUMNS
# columns where large day-to-day swings are normal (no spike filter)
NO_SPIKE_FILTER = {"storage_injection", "storage_withdrawal", "interconnector_exports", "interconnector_iuk", "interconnector_bbl",
                   "barrow", "teesside"} | set(ENTRY_COLUMNS)   # small terminals that sit at zero most days: a median-based filter misfires on them

HEADERS = {"Content-Type": "application/json", "Accept": "application/json, text/plain, */*",
           "Referer": "https://data.nationalgas.com/find-gas-data/view",
           "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"}


def classify_flow(name):
    n = name.lower()
    if "storage" in n and "withdraw" in n:
        return "storage_withdrawal"
    if "iuk" in n:
        return "interconnector_iuk"
    if "bbl" in n:
        return "interconnector_bbl"
    return None


def fetch(d0, d1, tries=5):
    body = {"latestFlag": "Y", "applicableFor": "Y", "dateFrom": d0.isoformat(), "dateTo": d1.isoformat(),
            "dateType": "GASDAY", "ids": ",".join(DEMAND_IDS + SUPPLY_IDS + FLOW_IDS + ENTRY_IDS)}
    for i in range(tries):
        try:
            r = requests.post(URL, json=body, headers=HEADERS, timeout=180)
            if r.ok:
                return r.json().get("data", [])
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(5 * (i + 1))
                continue
            raise RuntimeError(f"HTTP {r.status_code}: {r.text[:300]}")
        except requests.RequestException:
            time.sleep(5 * (i + 1))
    raise RuntimeError(f"gave up on {d0}..{d1}")


def to_frame(items, unmatched):
    recs = []
    for it in items:
        name = it.get("itemName", "")
        col = NAME_TO_COLUMN.get(name) or ENTRY_NAMES.get(name.replace(" ", "").lower()) or classify_flow(name)
        if col is None:
            unmatched.add(name)
            continue
        try:
            recs.append((pd.to_datetime(it["applicableFor"], format="%d/%m/%Y"), col, float(it["value"]) / 1e6))   # kWh -> GWh
        except (KeyError, ValueError, TypeError):
            continue
    if not recs:
        return pd.DataFrame(columns=COLUMNS)
    d = pd.DataFrame(recs, columns=["date", "col", "gwh"])
    out = d.pivot_table(index="date", columns="col", values="gwh", aggfunc="last").sort_index()
    if "storage_withdrawal" in out:   # the source reports withdrawals as negative injections
        out["storage_withdrawal"] = out["storage_withdrawal"].abs()
    for c in COLUMNS:
        if c not in out:
            out[c] = np.nan
    return out[COLUMNS]


def drop_spikes(df, window=15, threshold=6.0):
    """Blank one-off bad values: more than `threshold` x away from the centred rolling median, with the denominator floored
    at 5% of the column's typical size so a legitimate near-zero day does not trip it."""
    out = df.copy()
    for c in df.columns:
        if c in NO_SPIKE_FILTER:
            continue
        s = df[c].astype(float)
        med = s.rolling(window, center=True, min_periods=5).median()
        scale = s.abs().median()
        if pd.isna(scale) or scale == 0:
            continue
        ratio = (s - med).abs() / np.maximum(med.abs(), scale * 0.05)
        bad = (ratio > threshold) & med.notna() & s.notna()
        for d in s.index[bad]:
            print(f"  dropped {c} {d:%Y-%m-%d} = {s[d]:.1f} (local median {med[d]:.1f})")
        out.loc[bad, c] = np.nan
    return out


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
    if have and all(c in old and old[c].notna().any() for c in ENTRY_COLUMNS if c != "entry_burton_point"):
        fs = max(start, max(have) - timedelta(days=REVISION_DAYS))
        gaps = [start + timedelta(days=i) for i in range((today - start).days)]
        gaps = [g for g in gaps if g >= today - timedelta(days=120) and g not in have and g < fs]
        if gaps:
            fs = min(fs, gaps[0])
    unmatched, parts = set(), []
    d0 = fs
    while d0 <= today:
        d1 = min(d0 + timedelta(days=WINDOW_DAYS - 1), today)
        items = fetch(d0, d1)
        f = to_frame(items, unmatched)
        print(f"  {d0} .. {d1}: {len(items)} items, {len(f)} days", flush=True)
        parts.append(f)
        d0 = d1 + timedelta(days=1)
        time.sleep(0.5)
    if unmatched:
        print("unrecognised item names in the response:", sorted(unmatched))
    new = pd.concat(parts) if parts else pd.DataFrame(columns=COLUMNS)
    combined = old[~old.index.isin(new.index)] if len(old) else old
    combined = pd.concat([combined, new]) if len(combined) else new
    combined = drop_spikes(combined.sort_index().astype(float)).round(2).dropna(how="all")
    combined.index.name = "date"
    if combined.empty:
        print("no GB gas data")
        return
    got = combined.notna().sum()
    print("days per column:", got.to_dict())
    lines = ["Great Britain - gas supply and demand (National Gas Transmission, NTS data)", "",
             "Source", "National Gas Transmission Data Portal, Find gas data: https://data.nationalgas.com/find-gas-data . "
             "Free, no key. Operator's own data.",
             "", "Units and definitions",
             "GWh per gas day (the source is kWh). Demand (NTS Energy Offtaken actuals): ldz_offtake (local distribution zones), "
             "powerstations, industrial_offtake, interconnector_exports, storage_injection. Supply: entry terminal nominations "
             "(Day Ahead Net Aggregate) at Bacton UKCS, Barrow, Easington, St Fergus, Teesside - St Fergus and Easington also carry "
             "Norwegian pipeline gas mixed with UKCS production -, plus storage_withdrawal and the interconnector entry points "
             "interconnector_iuk (Belgium) and interconnector_bbl (Netherlands) - the portal returned no data for the IUK, BBL and Teesside items "
             "(columns stay empty), so pipeline imports from Belgium and the Netherlands are not in this workbook. LNG terminals are not "
             "either (see GIE ALSI). entry_* columns (System Entry Energy, D+2 allocations): entry_rough_subterminal = Rough sub-terminal (Easington) entry, "
             "entry_rough_storage = Rough storage withdrawals (inside ENTSOG's Easington entry until Sep 2025), entry_saltfleetby / entry_murrow / entry_glentham / "
             "entry_burton_point = small onshore and biomethane entries. The portal's history starts in Oct 2021.",
             f"Re-fetches the last {REVISION_DAYS} days each run plus gaps within 120 days; history from {args.start}. "
             "One-off bad values (more than 6x the local median) are blanked.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(combined)} days, "
             f"{combined.index.min():%Y-%m-%d} to {combined.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": combined}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}: {len(combined)} days x {combined.shape[1]} columns")


if __name__ == "__main__":
    main()
