"""
Colombia daily power generation by fuel from XM, in the standard grid-operator
layout shared by every country (scheduled in GitHub Actions; the owner's
local COLOMBIA_XM_GENERATION.py, whose XM calls this reuses, is unchanged).

Source: XM's free public API (servapibi.xm.com.co):
  - POST https://servapibi.xm.com.co/lists  {"MetricId": "ListadoRecursos", "Entity": "Sistema"}
    every plant's code and fuel ('EnerSource': AGUA, GAS, CARBON, ...)
  - POST https://servapibi.xm.com.co/hourly {"MetricId": "Gene", "Entity": "Recurso", "StartDate", "EndDate"}
    each plant's hourly generation in kWh (Hour01..Hour24); XM runs about two weeks behind.
Each plant-day is summed (kWh -> MWh) and grouped by the plant's fuel ('By XM fuel' tab),
then mapped onto the standard fuels ('Daily').

Incremental: the existing workbook is read back; only missing days since
START plus the last REFRESH_DAYS days (XM fills in and revises late) are
fetched. The workbook is checkpointed after each block. --full rebuilds.

Output (sheet 'Daily'): date, Hydro_MWh, Gas_MWh, Wind_MWh, Solar_MWh,
Coal_MWh, Oil_MWh, Bioenergy_MWh, Other_MWh, Total_MWh (no nuclear in Colombia).

Usage: python3 south_america/COLOMBIA_XM_GENERATION_DAILY.py [--out PATH] [--full]
"""

import argparse
import os
import sys
from datetime import date, timedelta

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, REPO)  # xlsx_notes
sys.path.insert(0, HERE)  # COLOMBIA_XM_GENERATION

import xlsx_notes  # noqa: E402
import COLOMBIA_XM_GENERATION as xm  # noqa: E402  (plant list + hourly fetch)

START = date(2021, 1, 1)
REFRESH_DAYS = 21
BLOCK_DAYS = 84  # checkpoint the workbook every 6 XM requests (14 days each)
DEFAULT_OUT = os.path.join(REPO, "output", "Data and Chart Outputs", "colombia_power_generation_daily.xlsx")

FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Oil", "Bioenergy", "Other"]
# XM fuel label (EnerSource) -> standard fuel
FUEL_MAP = {
    "AGUA": "Hydro",
    "GAS": "Gas", "GLP": "Gas",
    "VIENTO": "Wind",
    "RAD SOLAR": "Solar",
    "CARBON": "Coal",
    "COMBUSTOLEO": "Oil", "ACPM": "Oil", "JET-A1": "Oil",
    "BAGAZO": "Bioenergy", "BIOGAS": "Bioenergy", "BIOMASA": "Bioenergy",
}

NOTES = [
    "UNITS",
    "MWh per day (energy generated that day): XM's hourly kWh per plant (Hour01..Hour24) summed over the day "
    "and divided by 1,000. Total_MWh = sum of the fuel columns.",
    "",
    "TABS",
    "'Daily': standard layout shared by every country's grid-operator workbook. "
    "'By XM fuel': the same days in MWh per XM fuel label (each plant's EnerSource), before mapping. "
    "'Plant fuels': plant code -> XM fuel label, kept across runs so a plant XM drops from its list keeps its fuel; "
    "the gas combined cycles Termosierra, Termovalle and Termoemcali (listed by XM under backup ACPM) count as GAS.",
    "",
    "CATEGORY MAPPING (XM EnerSource -> column)",
    "Hydro = AGUA. Gas = GAS, GLP (domestic gas and imported LNG via SPEC Cartagena are not split by XM). "
    "Wind = VIENTO. Solar = RAD SOLAR. Coal = CARBON. Oil = COMBUSTOLEO (fuel oil), ACPM (diesel), JET-A1. "
    "Bioenergy = BAGAZO (bagasse), BIOGAS, BIOMASA. Other = plants missing from XM's current plant list "
    "(UNKNOWN - e.g. retired units) and any new label not yet mapped. Colombia has no nuclear, so there is no "
    "Nuclear_MWh column.",
    "",
    "COVERAGE",
    "Every resource XM reports generation for (centrally dispatched plants, smaller non-dispatched plants, "
    "cogenerators and self-generators that deliver to the grid). Rooftop/behind-the-meter solar is not included. "
    "Fuels are each plant's fuel in XM's current plant list, applied to the whole history.",
    "",
    "SOURCE",
    "XM S.A. E.S.P. public API (servapibi.xm.com.co): /hourly MetricId=Gene Entity=Recurso (kWh per plant per "
    "hour) and /lists MetricId=ListadoRecursos Entity=Sistema (plant fuel). History from 2021-01-01; XM "
    "publishes about two weeks behind and the last 21 days are re-fetched on every run.",
]


def load_existing(path):
    try:
        df = pd.read_excel(path, sheet_name="By XM fuel", index_col=0)
    except (FileNotFoundError, ValueError) as e:
        print(f"No existing workbook to extend ({type(e).__name__}) - full history", flush=True)
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index)
    return df


def ranges_to_fetch(existing, today, full):
    """Contiguous (start, end) date ranges: gaps in the history plus the refresh window."""
    if full or existing.empty:
        return [(START, today)]
    have = set(existing.index.date)
    last = max(have)
    refresh_from = max(START, last - timedelta(days=REFRESH_DAYS))
    days = [d.date() for d in pd.date_range(START, refresh_from - timedelta(days=1)) if d.date() not in have]
    days += [d.date() for d in pd.date_range(refresh_from, today)]
    out = []
    for d in days:
        if out and d == out[-1][1] + timedelta(days=1):
            out[-1] = (out[-1][0], d)
        else:
            out.append((d, d))
    return out


_WARNED = set()


def to_standard(by_fuel):
    unknown = [c for c in by_fuel.columns if c not in FUEL_MAP and c not in _WARNED]
    if unknown:
        _WARNED.update(unknown)
        print(f"Fuel labels counted as Other: {unknown}", flush=True)
    out = pd.DataFrame(index=by_fuel.index)
    for f in FUELS:
        cols = [c for c in by_fuel.columns if FUEL_MAP.get(c, "Other") == f]
        out[f"{f}_MWh"] = by_fuel[cols].sum(axis=1, min_count=1) if cols else 0.0
    out = out.fillna(0.0)
    out["Total_MWh"] = out.sum(axis=1)
    out.index.name = "date"
    return out.round(1)


def load_plant_fuels(path):
    """Saved plant code -> XM fuel label (None when the workbook predates the saved table)."""
    try:
        d = pd.read_excel(path, sheet_name="Plant fuels", dtype=str)
    except (FileNotFoundError, ValueError):
        return None
    return dict(zip(d["code"], d["xm_fuel"]))


PLANT_FUELS = {}


def save(path, by_fuel):
    by_fuel = by_fuel.sort_index().round(1)
    by_fuel.index.name = "date"
    daily = to_standard(by_fuel)
    pf = pd.DataFrame(sorted(PLANT_FUELS.items()), columns=["code", "xm_fuel"]).set_index("code")
    xlsx_notes.write_workbook(path, {"Daily": daily, "By XM fuel": by_fuel, "Plant fuels": pf}, NOTES,
                              {"UNITS", "TABS", "CATEGORY MAPPING (XM EnerSource -> column)", "COVERAGE", "SOURCE"})
    return daily


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=DEFAULT_OUT, help="workbook to write (and extend)")
    ap.add_argument("--full", action="store_true", help="ignore the existing workbook and rebuild from 2021")
    args = ap.parse_args()
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)

    saved_fuels = load_plant_fuels(args.out)
    if saved_fuels is None and not args.full:   # first run with the saved plant table (and the gas-CC fix): rebuild
        print("No saved plant-fuel table yet: rebuilding the history once with the current mapping", flush=True)
        args.full = True
    by_fuel = pd.DataFrame() if args.full else load_existing(args.out)
    today = date.today()
    ranges = ranges_to_fetch(by_fuel, today, args.full)
    print(f"Existing days: {len(by_fuel):,}; ranges to fetch: {[(str(a), str(b)) for a, b in ranges]}", flush=True)
    print("Plant list...", flush=True)
    fuels = xm.plant_fuels(saved_fuels)
    PLANT_FUELS.update(fuels)
    print("Fuel labels in the plant list:", sorted(set(fuels.values())), flush=True)

    for a, b in ranges:
        cur = a
        while cur <= b:
            block_end = min(cur + timedelta(days=BLOCK_DAYS - 1), b)
            print(f"Hourly generation {cur}..{block_end}", flush=True)
            fresh = xm.fetch_generation(cur, block_end, fuels)
            if not fresh.empty:
                fresh = fresh[fresh.index >= pd.Timestamp(START)]
                # fetched days replace what was there (a fuel missing from a fetched day is really absent)
                by_fuel = pd.concat([by_fuel.drop(index=fresh.index, errors="ignore"), fresh]) if not by_fuel.empty else fresh
                save(args.out, by_fuel)
            cur = block_end + timedelta(days=1)

    if xm.UNKNOWN_MWH:
        top = sorted(xm.UNKNOWN_MWH.items(), key=lambda kv: -kv[1])[:15]
        print(f"Generation from plants with no fuel (counted as Other), GWh: {[(c, round(m / 1000, 1)) for c, m in top]}",
              flush=True)
    if by_fuel.empty:
        sys.exit("No XM data")
    daily = save(args.out, by_fuel)
    gaps = pd.date_range(daily.index.min(), daily.index.max()).difference(daily.index)
    print(f"\nSaved {args.out}: {len(daily):,} days, {daily.index.min():%Y-%m-%d} to {daily.index.max():%Y-%m-%d}; "
          f"{len(gaps)} missing day(s){': ' + ', '.join(f'{d:%Y-%m-%d}' for d in gaps[:20]) if len(gaps) else ''}")
    print("\nLatest days (MWh):")
    print(daily.tail(5).to_string())
    print("\nLatest months (GWh):")
    print((daily.resample("MS").sum() / 1000).round(0).tail(6).to_string())


if __name__ == "__main__":
    main()
