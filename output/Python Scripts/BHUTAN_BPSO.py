"""
Bhutan electricity from BPSO, the Bhutan Power System Operator (Bhutan Power Corporation),
https://www.bpso.bt/home/energy. Found via discovery_archive/subcontinent/SEA_DISCOVERY_SOUTH_ASIA.py.

BPSO's public data endpoint returns one month of daily values per series (JSON, no key):
  https://www.bpso.bt/publicdata/energy_data/<series>/<YYYY-MM-01>
Series pulled: generation_mwh (total generation; Bhutan's grid is almost entirely run-of-river hydro),
energy_met_mwh (domestic consumption met), peak_demand_mw, energy_export_mwh / energy_import_mwh
(cross-border trade with India), iex_export_mwh / iex_import_mwh (trade on India's IEX).
Daily values start in 2023.

Trade definitions (checked against the balance generation + imports - exports = energy met): energy_export_mwh is
TOTAL exports (its IEX part is iex_export_mwh; adding iex_export again over-counts, -2.7% vs +0.7% residual on IEX
export days). energy_import_mwh is bilateral imports only (non-zero on ~19 days): winter imports are bought on
IEX and appear only in iex_import_mwh. So Imports_MWh = energy_import + iex_import (total imports); the parts are
kept as Imports_bilateral_MWh and IEX_imports_MWh. With these, generation + imports - exports matches energy met
(median ratio 1.003 over all days, 1.000 on winter import days, 88% of those within 3%).

Data checks: BPSO's daily figures carry occasional keying slips (e.g. 8,200 for 82,000 MWh). A day below
30% (or above 4x) the centred 15-day median of its series is blanked and listed on the Flags sheet, not charted.

Writes output/Data and Chart Outputs/bhutan_power_generation_daily.xlsx:
  Daily   standard layout: Hydro_MWh = total generation (Total_MWh the same), plus Exports/Imports/Energy met
  Demand  daily peak demand, MW (Demand_peak_MW)
  Flags   values blanked by the check above (date, series, value, median)

Incremental: the Daily sheet is the history store; months with a missing day, plus the latest two months
(revisions), are fetched. Runs on the 1st and 15th.

    python3 asia/BHUTAN_BPSO.py
"""
import argparse
import os
import sys
import time
from datetime import date

import pandas as pd
import requests
import urllib3

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

urllib3.disable_warnings()
URL = "https://www.bpso.bt/publicdata/energy_data/{series}/{month}"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Referer": "https://www.bpso.bt/home/energy", "X-Requested-With": "XMLHttpRequest",
     "Accept": "application/json, */*"}
T = (15, 90)
DATA_START = date(2023, 7, 1)   # BPSO's daily data starts in July 2023
OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "bhutan_power_generation_daily.xlsx")
SERIES = {"generation_mwh": "Generation_MWh", "energy_met_mwh": "Energy_met_MWh", "peak_demand_mw": "Demand_peak_MW",
          "energy_export_mwh": "Exports_MWh", "energy_import_mwh": "Imports_MWh",
          "iex_export_mwh": "IEX_exports_MWh", "iex_import_mwh": "IEX_imports_MWh"}
CHECK = ["Generation_MWh", "Energy_met_MWh", "Demand_peak_MW", "Exports_MWh"]
SMOOTH = ["Generation_MWh", "Energy_met_MWh", "Demand_peak_MW"]


def out(*a):
    print(*a, flush=True)


def month_series(series, month):
    for i in range(4):
        try:
            r = requests.get(URL.format(series=series, month=month.strftime("%Y-%m-01")), headers=H, timeout=T,
                             verify=False)   # BPSO serves an incomplete certificate chain
            r.raise_for_status()
            rows = r.json().get("data") or []
            s = pd.Series({pd.Timestamp(x["r_date"]): pd.to_numeric(x["r_data"], errors="coerce") for x in rows},
                          dtype=float)
            return s[s.index.to_period("M") == pd.Period(month, "M")]   # the 2023 placeholder rows are monthly
        except (requests.RequestException, ValueError) as e:
            if i == 3:
                out(f"  {series} {month:%Y-%m} failed: {e}")
                return pd.Series(dtype=float)
            time.sleep(10 * (i + 1))


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

    old = read_sheet(args.out, "Raw")
    today = date.today()
    months = pd.period_range(DATA_START, today, freq="M")
    have = set(old.dropna(subset=["Generation_MWh"]).index.normalize()) if not old.empty else set()
    todo = [m for m in months
            if m >= months[-2] or any(pd.Timestamp(d) not in have for d in pd.date_range(m.start_time, min(
                m.end_time.normalize(), pd.Timestamp(today) - pd.Timedelta(days=1))))]
    out(f"{len(have)} days saved; fetching {len(todo)} months")
    new = {}
    for m in todo:
        cols = {}
        for series, col in SERIES.items():
            cols[col] = month_series(series, m.start_time.date())
            time.sleep(1)
        df = pd.DataFrame(cols)
        if not df.empty:
            new[m] = df
        out(f"  {m}: {len(df)} days")
    raw = old.copy()
    if new:
        add = pd.concat(new.values())
        # a series that failed to download this run (empty) keeps its saved values: new values win only where present
        raw = add.combine_first(raw).sort_index() if not raw.empty else add.sort_index()
    raw = raw[raw["Generation_MWh"].fillna(0) > 0] if "Generation_MWh" in raw else raw
    if raw.empty:
        raise SystemExit("No BPSO data")
    raw.index.name = "date"

    clean, flags = raw.copy(), []
    num = clean.select_dtypes("number").columns
    clean[num] = clean[num].mask(clean[num] >= 99999).mask(clean[num] < 0)   # BPSO's 99999.999 placeholder; negatives
    for c in CHECK:
        if c not in clean:
            continue
        med = clean[c].rolling(15, center=True, min_periods=5).median()
        bad = clean[c] < 0.3 * med
        if c in SMOOTH:   # trade jumps with the season and imports are mostly zero: only the smooth series get an upper bound
            bad |= clean[c] > 4 * med
        for d in clean.index[bad]:
            flags.append({"date": d, "series": c, "value": clean.at[d, c], "median_15d": round(med[d], 1)})
        clean.loc[bad, c] = None
    flags = pd.DataFrame(flags, columns=["date", "series", "value", "median_15d"]).set_index("date")
    out(f"{len(flags)} values blanked by the check")

    daily = pd.DataFrame({"Hydro_MWh": clean["Generation_MWh"]})
    daily["Total_MWh"] = daily["Hydro_MWh"]
    for c in ("Energy_met_MWh", "Exports_MWh"):
        if c in clean:
            daily[c] = clean[c]
    # total imports = bilateral (energy_import, rare) + IEX (where nearly all winter imports appear)
    bil, iex = clean.get("Imports_MWh"), clean.get("IEX_imports_MWh")
    parts = pd.DataFrame({"b": bil, "i": iex}, index=clean.index)
    daily["Imports_MWh"] = parts.sum(axis=1, min_count=1)
    daily["Imports_bilateral_MWh"] = parts["b"]
    daily["IEX_imports_MWh"] = parts["i"]
    if "IEX_exports_MWh" in clean:
        daily["IEX_exports_MWh"] = clean["IEX_exports_MWh"]   # part of Exports_MWh, not additional
    daily["Net_imports_MWh"] = daily["Imports_MWh"].fillna(0) - daily.get("Exports_MWh", 0).fillna(0)
    if "Energy_met_MWh" in daily:   # generation + imports - exports - energy met (losses, auxiliary, slips)
        daily["Balance_residual_MWh"] = (daily["Hydro_MWh"] + daily["Net_imports_MWh"] - daily["Energy_met_MWh"]).round(1)
        q = (daily["Hydro_MWh"] + daily["Net_imports_MWh"]) / daily["Energy_met_MWh"]
        out(f"balance check (generation + imports - exports) / energy met: median {q.median():.4f}, "
            f"{((q - 1).abs() <= 0.03).mean():.0%} of days within 3%")
    demand = clean[["Demand_peak_MW"]] if "Demand_peak_MW" in clean else pd.DataFrame(index=clean.index)
    notes = [
        "UNITS",
        "Daily: MWh per day. Hydro_MWh = BPSO total generation (Bhutan's generation is run-of-river hydro apart from "
        "a few MW of solar, which BPSO does not report separately); Total_MWh the same. Energy_met_MWh = domestic "
        "energy met. Exports_MWh = total exports to India (BPSO energy_export; IEX_exports_MWh is the part sold on "
        "India's power exchange, already inside Exports_MWh). Imports_MWh = TOTAL imports = Imports_bilateral_MWh "
        "(BPSO energy_import, rarely non-zero) + IEX_imports_MWh (bought on IEX: nearly all winter imports). "
        "Net_imports_MWh = Imports - Exports. Balance_residual_MWh = generation + imports - exports - energy met "
        "(transmission losses, auxiliary use and keying slips; median ratio 1.003, 1.000 on winter import days). "
        "Before this fix Imports_MWh held only the bilateral part.",
        "Demand: Demand_peak_MW = daily peak demand met, MW.",
        "Raw: every series as published (before the check below).",
        "Flags: values blanked because they fall below 30% (generation, energy met, peak: or above 4x) the centred 15-day median of their series - "
        "BPSO's daily sheets carry occasional keying slips (e.g. 8,200 for 82,000 MWh). BPSO's placeholder 99999.999 "
        "and negative energy values are blanked too (not listed).",
        "",
        "COVERAGE",
        f"Daily from {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}.",
        "",
        "SOURCE",
        "BPSO (Bhutan Power System Operator, Bhutan Power Corporation), energy data: https://www.bpso.bt/home/energy "
        "(JSON at https://www.bpso.bt/publicdata/energy_data/<series>/<YYYY-MM-01>).",
    ]
    xlsx_notes.write_workbook(args.out, {"Daily": daily, "Demand": demand, "Flags": flags, "Raw": raw}, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})
    out(f"Saved {args.out}: {len(daily)} days {daily.index.min():%Y-%m-%d}..{daily.index.max():%Y-%m-%d}")
    out(daily.tail(3).to_string())


if __name__ == "__main__":
    main()
