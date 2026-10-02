"""
Australia NEM (east coast + Tasmania) power generation by fuel, daily, from
AEMO's 5-minute unit SCADA (DISPATCH_UNIT_SCADA, via the nemosis package -
public NEMWEB / MMSDM data, no key). Units are mapped to region and fuel with
AEMO's NEM Registration and Exemption List (the mapping in
rest_of_world/aemo_nemweb_power_mix.py).

Writes au_nem_power_generation_daily.xlsx:
  Daily    standard layout: date, <Fuel>_MWh (Hydro, Gas, Wind, Solar, Coal,
           Oil, Bioenergy), Total_MWh, plus Battery_discharge_MWh and
           Pumped_hydro_MWh (storage, kept out of Total_MWh)
  States   MWh per day by NEM region (NSW, QLD, SA, TAS, VIC), generation only
Energy = sum of 5-minute MW / 12 (negative readings clipped to 0).

Incremental: fetches months not saved yet plus the latest saved month (the
month in progress); a new file backfills from 2021, --max-months per run
(the MMSDM monthly SCADA archive is large). Saves after every month.

Usage: python3 AU_NEM_GENERATION.py [--out "output/Data and Chart Outputs/au_nem_power_generation_daily.xlsx"]
"""
import argparse
import os
import shutil
import sys
from datetime import date

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
sys.path.insert(0, os.path.join(ROOT, "rest_of_world"))
import xlsx_notes  # noqa: E402
import aemo_nemweb_power_mix as nem  # noqa: E402

HISTORY_START = "2021-01-01"
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "au_nem_power_generation_daily.xlsx")
CACHE = os.path.join(ROOT, "nemosis_cache")
CATEGORY = {"coal": "Coal", "gas": "Gas", "oil": "Oil", "hydro": "Hydro", "wind": "Wind", "solar": "Solar",
            "biomass": "Bioenergy", "nuclear": "Nuclear", "battery": "Battery_discharge",
            "hydro_pumped": "Pumped_hydro"}
FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Oil", "Bioenergy"]
STORAGE = ["Battery_discharge", "Pumped_hydro"]
STATES = {"NSW1": "NSW", "QLD1": "QLD", "SA1": "SA", "TAS1": "TAS", "VIC1": "VIC"}


def fetch_month(duid_map, start, end):
    """[start, end) -> (daily MWh by fuel, daily generation MWh by state)."""
    from nemosis import dynamic_data_compiler

    os.makedirs(CACHE, exist_ok=True)
    try:
        df = dynamic_data_compiler(start.strftime("%Y/%m/%d %H:%M:%S"), end.strftime("%Y/%m/%d %H:%M:%S"),
                                   "DISPATCH_UNIT_SCADA", CACHE)
    finally:
        shutil.rmtree(CACHE, ignore_errors=True)   # the monthly archive is large; keep the runner's disk free
    if df is None or df.empty:
        return None, None
    m = df["DUID"].map(duid_map)
    df = df[m.notna()].assign(region=m.dropna().str[0], fuel=m.dropna().str[1].map(CATEGORY))
    # SETTLEMENTDATE is the interval END: 00:05..24:00 make up one day
    t = pd.to_datetime(df["SETTLEMENTDATE"]) - pd.Timedelta(minutes=5)
    df = df.assign(date=t.dt.normalize(),
                   mwh=pd.to_numeric(df["SCADAVALUE"], errors="coerce").clip(lower=0) / 12.0)
    fuel = df.pivot_table(index="date", columns="fuel", values="mwh", aggfunc="sum")
    gen = df[~df["fuel"].isin(STORAGE)]
    states = gen.pivot_table(index="date", columns="region", values="mwh", aggfunc="sum").rename(columns=STATES)
    return fuel, states


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    d.index = pd.to_datetime(pd.Index(d.index).astype(str), errors="coerce")
    return d[d.index.notna()].sort_index()


def save(path, fuel, states):
    daily = fuel.reindex(columns=[f for f in FUELS + ["Nuclear"] if f in fuel.columns]).add_suffix("_MWh")
    daily["Total_MWh"] = daily.sum(axis=1, min_count=1)
    for s in STORAGE:
        if s in fuel:
            daily[f"{s}_MWh"] = fuel[s]
    st = states.reindex(columns=[s for s in STATES.values() if s in states.columns]).add_suffix("_MWh")
    for x in (daily, st):
        x.index.name = "date"
    notes = [
        "UNITS",
        "MWh per day = sum of AEMO's 5-minute unit SCADA MW / 12 (negative readings clipped to 0), day = market "
        "day in AEST (intervals ending 00:05 to 24:00).",
        "Daily: Hydro (excl. pumped storage), Gas, Wind, Solar (utility scale), Coal, Oil (diesel/kerosene), "
        "Bioenergy; Total_MWh is their sum. Battery_discharge_MWh and Pumped_hydro_MWh are storage output, kept "
        "out of Total_MWh (they shift energy rather than generate it).",
        "States: generation (excl. storage) by NEM region, MWh per day.",
        "",
        "COVERAGE",
        f"National Electricity Market (QLD, NSW incl. ACT, VIC, SA, TAS), daily, {daily.index.min():%d %b %Y} to "
        f"{daily.index.max():%d %b %Y}. Not covered: Western Australia (separate WEM market - see "
        "au_wem_power_generation_daily.xlsx), Northern Territory, rooftop solar (not metered in unit SCADA). "
        "Units retired before the current registration list are not mapped to a fuel and are left out.",
        "",
        "SOURCE",
        "AEMO NEMWEB / MMSDM DISPATCH_UNIT_SCADA (via nemosis), unit fuel types from the NEM Registration and "
        f"Exemption List: {nem.REGISTRATION_LIST_URL}",
        "https://nemweb.com.au/",
    ]
    xlsx_notes.write_workbook(path, {"Daily": daily.round(1), "States": st.round(1)}, notes,
                              {"UNITS", "COVERAGE", "SOURCE"})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--max-months", type=int, default=12)
    args = ap.parse_args()

    raw = load(args.out, "Daily")
    fuel = raw[[c for c in raw.columns if c != "Total_MWh"]].rename(columns=lambda c: c[:-4]) if not raw.empty else raw
    states = load(args.out, "States")
    states = states.rename(columns=lambda c: c[:-4]) if not states.empty else states
    today = pd.Timestamp(date.today())
    months = pd.date_range(HISTORY_START, today, freq="MS")
    have = sorted(set(fuel.index.to_period("M").to_timestamp())) if not fuel.empty else []
    todo = [m for m in months if m not in have] + have[-1:]
    todo = sorted(set(todo))[:args.max_months]
    print(f"{len(have)} months saved; fetching {len(todo)}: {[f'{m:%Y-%m}' for m in todo]}", flush=True)
    duid_map = nem.fetch_duid_category_map()
    print(f"{len(duid_map)} NEM units mapped", flush=True)
    for m in todo:
        end = min(m + pd.offsets.MonthBegin(1), today)
        try:
            f, s = fetch_month(duid_map, m, end)
        except Exception as e:  # noqa: BLE001 - keep the months already done
            print(f"  {m:%Y-%m}: failed {type(e).__name__}: {str(e)[:200]}", flush=True)
            continue
        if f is None:
            print(f"  {m:%Y-%m}: no rows", flush=True)
            continue
        keep = lambda d: d[d.index.to_period("M") != m.to_period("M")] if not d.empty else d  # noqa: E731
        fuel = pd.concat([keep(fuel), f]).sort_index()
        states = pd.concat([keep(states), s]).sort_index()
        print(f"  {m:%Y-%m}: {len(f)} days, {f.drop(columns=STORAGE, errors='ignore').sum().sum() / 1e3:,.0f} GWh",
              flush=True)
        save(args.out, fuel, states)
    if fuel.empty:
        raise SystemExit("No NEM generation data")
    print(f"Saved {args.out}: {len(fuel)} days", flush=True)


if __name__ == "__main__":
    main()
