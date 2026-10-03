"""
Installed generating capacity by fuel for Europe from the ENTSO-E Transparency Platform (Installed Capacity per
Production Type [14.1.A], API document A68), one workbook per country:

  output/Data and Chart Outputs/<country>_power_capacity.xlsx
    sheet "Monthly": date (1 January of each year), Hydro_MW, Gas_MW, Wind_MW, Solar_MW, Coal_MW, Nuclear_MW, Oil_MW,
                     Bioenergy_MW, Other_MW, Total_MW, PumpedStorage_MW, Storage_MW
    sheet "Units": source and definitions

The standard capacity layout (annual rows dated 1 January), so add_charts.power_capacity() charts it and the master's
capacity-factor tab can use it. Total_MW excludes pumped storage and batteries (same as the other capacity
workbooks). Multi-zone countries are summed; a year is saved only if every zone reports it.

Incremental: years already saved are kept; the current and previous year are re-fetched every run (ENTSO-E
capacity is revised) along with any missing year since 2021. A year with no data for a country (e.g. a zone that
started reporting later) simply stays absent.

Usage: python3 ENTSOE_CAPACITY.py [--out-dir DIR] [--countries DE,FR] [--start-year 2021]
Requires ENTSOE_API_KEY (environment / GitHub secret, or api_keys.py).
"""
import argparse
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, ROOT)
import entsoe_common as C  # noqa: E402
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FUEL_COLS = [f"{f}_MW" for f in C.CAP_FUELS]
COLS = [f"{f}_MW" for f in C.CAP_FUELS[:9]] + ["Total_MW"] + [f"{f}_MW" for f in C.CAP_FUELS[9:]]


def read_existing(path):
    if not os.path.exists(path):
        return pd.DataFrame(columns=COLS)
    try:
        d = pd.read_excel(path, sheet_name="Monthly")
    except Exception as e:  # noqa: BLE001
        print(f"  could not read {os.path.basename(path)} ({type(e).__name__}: {e}); starting over")
        return pd.DataFrame(columns=COLS)
    d["date"] = pd.to_datetime(d["date"], errors="coerce")
    d = d.dropna(subset=["date"]).set_index("date").sort_index()
    for c in COLS:
        if c not in d:
            d[c] = float("nan")
    return d[COLS]


_ZONE_OFFSET = {}   # zone EIC -> UTC offset (hours) of its local midnight that worked


def zone_year(eic, year):
    """{psr: MW} for one zone and year, or None if ENTSO-E has nothing. Annual documents are aligned to the zone's
    local midnight (23:00 UTC on 31 December for CET zones), so the window is tried at UTC+0, +1, +2 and +3 until one
    returns data, and the zone remembers which worked."""
    offsets = ([_ZONE_OFFSET[eic]] if eic in _ZONE_OFFSET else []) + [o for o in (0, 1, 2, 3) if _ZONE_OFFSET.get(eic) != o]
    for off in offsets:
        d0 = datetime(year, 1, 1, tzinfo=timezone.utc) - timedelta(hours=off)
        d1 = datetime(year + 1, 1, 1, tzinfo=timezone.utc) - timedelta(hours=off)
        res = C.fetch_split({"documentType": "A68", "processType": "A33", "in_Domain": eic}, d0, d1,
                            C.parse_capacity, C.merge_capacity, min_days=400)
        time.sleep(0.25)
        out = {psr: yrs[year] for psr, yrs in (res or {}).items() if year in yrs}
        if out:
            _ZONE_OFFSET[eic] = off
            return out
    return None


# Capacity is often reported for the whole country rather than per bidding zone (Italy, Sweden): used when a zone has
# nothing for the year.
COUNTRY_EIC = {"IT": "10YIT-GRTN-----B", "SE": "10YSE-1--------K", "NO": "10YNO-0--------C", "DK": "10Y1001A1001A65H"}


def country_year(zones, year, code=None):
    total = {}
    missing = []
    for label, eic in zones:
        z = zone_year(eic, year)
        if z is None:
            missing.append(label)
            continue
        for psr, mw in z.items():
            col = C.CAP_COLUMN.get(psr)
            if col:
                total[col] = total.get(col, 0.0) + mw
    if missing:
        whole = zone_year(COUNTRY_EIC[code], year) if code in COUNTRY_EIC else None
        print(f"  {code} {year}: no capacity for zone(s) {', '.join(missing)}; country-level "
              f"{'used' if whole else 'not available'}", flush=True)
        if not whole:
            return None
        total = {}
        for psr, mw in whole.items():
            col = C.CAP_COLUMN.get(psr)
            if col:
                total[col] = total.get(col, 0.0) + mw
    return total


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default=OUT_DEFAULT)
    ap.add_argument("--countries", default="")
    ap.add_argument("--start-year", type=int, default=2021)
    args = ap.parse_args()
    C.api_key()
    os.makedirs(args.out_dir, exist_ok=True)
    this_year = datetime.now(timezone.utc).year
    want = [c.strip().upper() for c in args.countries.split(",") if c.strip()] or list(C.COUNTRIES)
    summary = []
    for code in want:
        if code not in C.COUNTRIES:
            print(f"unknown country code {code}")
            continue
        name, slug, zones = C.COUNTRIES[code]
        path = os.path.join(args.out_dir, f"{slug}_power_capacity.xlsx")
        old = read_existing(path)
        have = {d.year for d in old.index}
        todo = sorted(({this_year, this_year - 1} | (set(range(args.start_year, this_year + 1)) - have)) & set(range(args.start_year, this_year + 1)))
        rows = {}
        for y in todo:
            try:
                tot = country_year(zones, y, code)
            except Exception as e:  # noqa: BLE001
                print(f"  {code} {y}: {type(e).__name__}: {e}")
                continue
            if tot:
                rows[pd.Timestamp(y, 1, 1)] = tot
        new = pd.DataFrame.from_dict(rows, orient="index")
        if not new.empty:
            for f in C.CAP_FUELS:
                if f not in new:
                    new[f] = 0.0
            new = new.fillna(0.0)
            new["Total"] = new[C.CAP_FUELS[:9]].sum(axis=1)
            new = new.rename(columns=lambda c: f"{c}_MW")[COLS].round(0)
            new.index.name = "date"
        combined = pd.concat([old[~old.index.isin(new.index)], new]) if not new.empty else old
        combined = combined.sort_index()
        combined.index.name = "date"
        print(f"{code} {name}: fetched years {todo} -> {sorted(y.year for y in new.index) if not new.empty else 'none'}; "
              f"saved {len(combined)} years", flush=True)
        if combined.empty:
            summary.append((code, name, 0, "NO DATA"))
            continue
        lines = [f"{name} - installed generating capacity by fuel (ENTSO-E Transparency Platform)", "",
                 "Source", "ENTSO-E Transparency Platform, Installed Capacity per Production Type [14.1.A] (API document A68, "
                 "process A33 year ahead / as reported). https://transparency.entsoe.eu/",
                 "", "Zones", ", ".join(f"{label} ({eic})" for label, eic in zones),
                 "", "Units and definitions",
                 "MW installed, one row per year dated 1 January. Hydro = run-of-river + reservoir (pumped storage separate). "
                 "Gas includes coal-derived gas; Coal = hard coal, lignite, oil shale, peat; Bioenergy = biomass; Other = "
                 "geothermal, marine, other renewable, waste, other. Total_MW excludes PumpedStorage_MW and Storage_MW.",
                 "Multi-zone countries are summed and a year is kept only if every zone reports it.",
                 "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; years {combined.index.min().year}-{combined.index.max().year}"]
        xlsx_notes.write_workbook(path, {"Monthly": combined}, lines, {"Source", "Zones", "Units and definitions", "Last pull"})
        summary.append((code, name, len(combined), f"{combined.index.min().year}-{combined.index.max().year}"))
    print("\nSUMMARY")
    for code, name, n, rng in summary:
        print(f"  {code} {name:28s} {n:3d} years  {rng}")


if __name__ == "__main__":
    main()
