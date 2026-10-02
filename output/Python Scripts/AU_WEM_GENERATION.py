"""
Western Australia (WEM, the South West Interconnected System) power
generation by fuel, daily, from AEMO's public facility SCADA:

  to Sep 2023  https://data.wa.aemo.com.au/public/public-data/datafiles/facility-scada/facility-scada-YYYY-MM.csv
               (MWh per 30-minute trading interval per facility)
  since the WEM reform (Oct 2023)
               https://data.wa.aemo.com.au/public/market-data/wemde/facilityScada/previous/FacilityScada_YYYYMMDD.zip
               (JSON, MW per 5-minute dispatch interval per facility; MWh = MW / 12)

AEMO's WEM facilities list carries no fuel type, so facilities are
classified by their registered code (FUEL_RULES below: _WF wind, _PV/_SF
solar, _ESR/_BESS battery, Collie/Muja/Bluewaters coal, landfill/biogas,
the rest gas or gas/diesel peakers). Unmatched codes count as Gas and are
listed on the Units sheet.

Writes au_wem_power_generation_daily.xlsx, sheet "Daily" in the standard
layout (date, <Fuel>_MWh, Total_MWh, Battery_discharge_MWh kept out of the
total) and "By facility" (MWh per day per facility code, latest month).

Incremental: fetches days not saved yet plus the last 3 saved days; a new
file backfills from 2021. Saves every month of data.

Usage: python3 AU_WEM_GENERATION.py [--out "output/Data and Chart Outputs/au_wem_power_generation_daily.xlsx"]
"""
import argparse
import io
import json
import os
import re
import sys
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

BASE = "https://data.wa.aemo.com.au/public"
OLD = BASE + "/public-data/datafiles/facility-scada/facility-scada-{m}.csv"
NEW = BASE + "/market-data/wemde/facilityScada/previous/FacilityScada_{d}.zip"
NEW_LIST = BASE + "/market-data/wemde/facilityScada/previous/"
REFORM = pd.Timestamp("2023-10-01")   # first full day of WEMDE data (files from 29 Sep 2023)
HISTORY_START = pd.Timestamp("2021-01-01")
REFRESH_DAYS = 3
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "au_wem_power_generation_daily.xlsx")
FUELS = ["Gas", "Wind", "Solar", "Coal", "Oil", "Bioenergy"]
FUEL_RULES = [   # (regex on facility code, fuel) - first match wins
    (r"ESR|BESS|BATT|_BS\d|STORAGE", "Battery_discharge"),
    (r"_WF|WWF|WINDFARM|WIND", "Wind"),
    (r"_PV|_SF|SOLAR", "Solar"),
    (r"^MUJA|COLLIE|BW\d_BLUEWATERS|BLUEWATERS", "Coal"),
    (r"LFG|LANDFILL|BIOGAS|_IG\d|RENEWABLE|WOOD|BIOMASS|WTE|WASTE", "Bioenergy"),
    (r"DIESEL|_DG\d", "Oil"),
]


def fuel_of(code):
    for pat, fuel in FUEL_RULES:
        if re.search(pat, str(code), re.I):
            return fuel
    return "Gas"


def get(url, ok404=False):
    for attempt in range(4):
        try:
            r = requests.get(url, headers=HEADERS, timeout=(10, 180))
            if ok404 and r.status_code == 404:
                return None
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            print(f"    attempt {attempt + 1}/4 failed: {type(e).__name__}: {str(e)[:150]}", flush=True)
            time.sleep(5 * (attempt + 1))
    raise RuntimeError(f"failed: {url}")


def old_month(m):
    """Pre-reform monthly CSV -> daily MWh by facility (date = trading interval's calendar date, AWST)."""
    r = get(OLD.format(m=m.strftime("%Y-%m")), ok404=True)
    if r is None:
        return pd.DataFrame()
    d = pd.read_csv(io.BytesIO(r.content), low_memory=False)
    d["date"] = pd.to_datetime(d["Trading Date"], errors="coerce")
    d["mwh"] = pd.to_numeric(d["Energy Generated (MWh)"], errors="coerce").clip(lower=0)
    return d.pivot_table(index="date", columns="Facility Code", values="mwh", aggfunc="sum")


def new_day(day):
    """WEMDE daily zip (JSON, 5-minute MW, 08:00 to 08:00 AWST) -> one row of MWh by facility for that trading day."""
    r = get(NEW.format(d=day.strftime("%Y%m%d")), ok404=True)
    if r is None:
        return None
    z = zipfile.ZipFile(io.BytesIO(r.content))
    rows = json.loads(z.read(z.namelist()[0]))["data"]["facilityScadaDispatchIntervals"]
    d = pd.DataFrame(rows)
    d["mwh"] = pd.to_numeric(d["quantity"], errors="coerce").clip(lower=0) / 12.0
    out = d.groupby("code")["mwh"].sum().to_frame(day).T
    return out


def new_days():
    r = get(NEW_LIST)
    return sorted(pd.Timestamp(x) for x in set(re.findall(r"FacilityScada_(\d{8})\.zip", r.text)))


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    d.index = pd.to_datetime(pd.Index(d.index).astype(str), errors="coerce")
    return d[d.index.notna()].sort_index()


def merge(base, new):
    if new is None or new.empty:
        return base
    if base.empty:
        return new
    base = base[~base.index.isin(new.index)]
    return pd.concat([base, new]).sort_index()


def save(path, fac):
    fuel_map = {c: fuel_of(c) for c in fac.columns}
    by_fuel = fac.T.groupby(fuel_map).sum(min_count=1).T
    daily = by_fuel.reindex(columns=[f for f in FUELS if f in by_fuel.columns]).add_suffix("_MWh")
    daily["Total_MWh"] = daily.sum(axis=1, min_count=1)
    if "Battery_discharge" in by_fuel:
        daily["Battery_discharge_MWh"] = by_fuel["Battery_discharge"]
    daily.index.name = "date"
    gas = sorted(c for c, f in fuel_map.items() if f == "Gas")
    notes = [
        "UNITS",
        "MWh per day by fuel, Western Australia's main grid (SWIS). Pre-reform files give MWh per 30-minute "
        "trading interval; from Oct 2023 AEMO publishes MW per 5-minute dispatch interval (MWh = MW / 12). "
        "Negative readings are clipped to 0. Day = WEM trading day, 08:00 to 08:00 AWST.",
        "Daily: Gas (incl. dual-fuel gas/diesel peakers), Wind, Solar (utility scale), Coal, Oil (diesel units), "
        "Bioenergy (landfill gas, biomass, waste to energy); Total_MWh is their sum. Battery_discharge_MWh is "
        "storage output, kept out of Total_MWh.",
        "Fuel is classified from each facility's registered code (AEMO's WEM facilities list has no fuel field): "
        + "; ".join(f"{p} -> {f}" for p, f in FUEL_RULES) + "; everything else -> Gas.",
        "Facilities counted as Gas: " + ", ".join(gas),
        "By facility: MWh per day per facility code (all days saved).",
        "",
        "COVERAGE",
        f"WEM / SWIS, daily, {daily.index.min():%d %b %Y} to {daily.index.max():%d %b %Y}. Rooftop solar and the "
        "North West Interconnected System (Pilbara) are not included.",
        "",
        "SOURCE",
        f"AEMO WA public data: {OLD.format(m='YYYY-MM')} (to Sep 2023), {NEW.format(d='YYYYMMDD')} (from Oct 2023)",
        "https://data.wa.aemo.com.au/",
    ]
    f = fac.copy()
    f.index.name = "date"
    xlsx_notes.write_workbook(path, {"Daily": daily.round(1), "By facility": f.round(2)}, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()

    fac = load(args.out, "By facility")
    have = set(fac.index) if not fac.empty else set()
    # pre-reform: whole months
    for m in pd.date_range(HISTORY_START, REFORM - pd.Timedelta(days=1), freq="MS"):
        days = pd.date_range(m, m + pd.offsets.MonthEnd(0))
        if all(d in have for d in days):
            continue
        new = old_month(m)
        new = new[(new.index >= m) & (new.index < min(REFORM, m + pd.offsets.MonthBegin(1)))] if not new.empty else new
        fac = merge(fac, new)
        print(f"  {m:%Y-%m} (monthly CSV): {len(new)} days", flush=True)
        save(args.out, fac)
    # post-reform: one zip per day
    avail = [d for d in new_days() if d >= REFORM]
    have = set(fac.index) if not fac.empty else set()
    todo = [d for d in avail if d not in have] + sorted(d for d in have if d >= REFORM)[-REFRESH_DAYS:]
    todo = sorted(set(todo))
    print(f"WEMDE files {avail[0]:%Y-%m-%d}..{avail[-1]:%Y-%m-%d}; fetching {len(todo)} days", flush=True)
    for i in range(0, len(todo), 31):
        chunk = todo[i:i + 31]
        with ThreadPoolExecutor(max_workers=6) as ex:
            parts = [p for p in ex.map(new_day, chunk) if p is not None]
        if parts:
            new = pd.concat(parts)
            fac = merge(fac, new)
        print(f"  {chunk[0]:%Y-%m-%d}..{chunk[-1]:%Y-%m-%d}: {len(parts)} files", flush=True)
        save(args.out, fac)
    if fac.empty:
        raise SystemExit("No WEM data")
    print(f"Saved {args.out}: {len(fac)} days, {fac.shape[1]} facilities", flush=True)


if __name__ == "__main__":
    main()
