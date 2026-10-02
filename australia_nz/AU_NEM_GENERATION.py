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
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_notes  # noqa: E402
import aemo_registration  # noqa: E402

HISTORY_START = "2021-01-01"
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "au_nem_power_generation_daily.xlsx")
CACHE = os.path.join(ROOT, "nemosis_cache")
FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Oil", "Bioenergy"]
STORAGE = ["Battery_discharge", "Pumped_hydro"]
RENAME = {"Battery_storage": "Battery_discharge", "Pumped_storage": "Pumped_hydro"}
FETCH_VERSION = 4   # bump when a month's saved content changes (2: state balance, 3: storage charging,
#   4: units that left the registration list - retired plant, pre-2024 battery/pump load DUIDs) - refetches once
CHARGE = {"Battery_discharge": "Battery_charge", "Pumped_hydro": "Pumped_hydro_pumping"}   # storage consumption
STATES = ["NSW", "QLD", "SA", "TAS", "VIC"]


def _table(name, start, end):
    from nemosis import dynamic_data_compiler

    os.makedirs(CACHE, exist_ok=True)
    try:
        return dynamic_data_compiler(start.strftime("%Y/%m/%d %H:%M:%S"), end.strftime("%Y/%m/%d %H:%M:%S"),
                                     name, CACHE)
    finally:
        shutil.rmtree(CACHE, ignore_errors=True)   # the monthly archives are large; keep the runner's disk free


def _day(ts, minutes):
    """Interval-END timestamps -> market day (intervals ending 00:05..24:00, or 00:30..24:00, make up one day)."""
    return (pd.to_datetime(ts) - pd.Timedelta(minutes=minutes)).dt.normalize()


def region_balance(start, end):
    """Per region and day: operational demand and net imports (MWh, from DISPATCHREGIONSUM) and rooftop solar
    (MWh, ROOFTOP_PV_ACTUAL) where available. Columns <STATE>_<item>."""
    out = []
    rs = _table("DISPATCHREGIONSUM", start, end)
    if rs is not None and not rs.empty:
        if "INTERVENTION" in rs:
            rs = rs[pd.to_numeric(rs["INTERVENTION"], errors="coerce").fillna(0) == 0]
        rs = rs.assign(date=_day(rs["SETTLEMENTDATE"], 5), state=rs["REGIONID"].str[:-1],
                       demand=pd.to_numeric(rs["TOTALDEMAND"], errors="coerce") / 12.0,
                       # NETINTERCHANGE > 0 is a net EXPORT from the region
                       imports=-pd.to_numeric(rs["NETINTERCHANGE"], errors="coerce") / 12.0)
        g = rs.groupby(["date", "state"])[["demand", "imports"]].sum().unstack("state")
        g.columns = [f"{st}_{'Operational_demand' if k == 'demand' else 'Net_imports'}" for k, st in g.columns]
        out.append(g)
    try:
        pv = _table("ROOFTOP_PV_ACTUAL", start, end)
        if pv is not None and not pv.empty:
            pv = pv[pv["REGIONID"].isin([f"{st}1" for st in STATES])]
            if "QI" in pv:   # one estimate per half-hour: the highest-quality one
                pv = pv.sort_values("QI", ascending=False)
            pv = pv.drop_duplicates(["INTERVAL_DATETIME", "REGIONID"])
            pv = pv.assign(date=_day(pv["INTERVAL_DATETIME"], 30), state=pv["REGIONID"].str[:-1],
                           mwh=pd.to_numeric(pv["POWER"], errors="coerce") / 2.0)
            g = pv.groupby(["date", "state"])["mwh"].sum().unstack("state")
            g.columns = [f"{st}_Rooftop_solar" for st in g.columns]
            out.append(g)
    except Exception as e:  # noqa: BLE001 - the balance still works without rooftop solar
        print(f"    rooftop PV not available: {type(e).__name__}: {str(e)[:150]}", flush=True)
    return pd.concat(out, axis=1) if out else pd.DataFrame()


def fetch_month(duid_map, start, end):
    """[start, end) -> (daily MWh by fuel, daily generation MWh by state, daily state power balance)."""
    df = _table("DISPATCH_UNIT_SCADA", start, end)
    if df is None or df.empty:
        return None, None, None
    unk = df[~df["DUID"].isin(duid_map["duid"])]
    if not unk.empty:   # SCADA from units with no region/fuel (left out of every total) - largest first
        top = (pd.to_numeric(unk["SCADAVALUE"], errors="coerce").abs() / 12e3).groupby(unk["DUID"]).sum()
        print(f"    unmapped DUIDs (GWh): {top.sort_values(ascending=False).head(12).round(1).to_dict()}", flush=True)
    df = df.merge(duid_map, left_on="DUID", right_on="duid", how="inner")
    raw = pd.to_numeric(df["SCADAVALUE"], errors="coerce") / 12.0
    df = df.assign(date=_day(df["SETTLEMENTDATE"], 5))
    # storage charging: a scheduled-load DUID's consumption, or a bidirectional unit's negative output
    store = df["fuel"].isin(STORAGE)
    charge = raw.where(df["role"].eq("load"), (-raw).clip(lower=0)).where(store).clip(lower=0)
    ch = df.assign(mwh=charge)[store].copy()
    ch["fuel"] = ch["fuel"].map(CHARGE)
    df = df[df["role"].ne("load")].assign(mwh=raw.clip(lower=0))
    fuel = df.pivot_table(index="date", columns="fuel", values="mwh", aggfunc="sum")
    gen = df[~df["fuel"].isin(STORAGE)]
    states = gen.pivot_table(index="date", columns="region", values="mwh", aggfunc="sum")
    fuel = fuel.join(ch.pivot_table(index="date", columns="fuel", values="mwh", aggfunc="sum"), how="left")
    sf = pd.concat([df, ch]).pivot_table(index="date", columns=["region", "fuel"], values="mwh", aggfunc="sum")
    sf.columns = [f"{st}_{f}" for st, f in sf.columns]
    try:
        bal = pd.concat([sf, region_balance(start, end)], axis=1)
    except Exception as e:  # noqa: BLE001 - generation by state and fuel is still saved
        print(f"    region demand/interchange failed: {type(e).__name__}: {str(e)[:150]}", flush=True)
        bal = sf
    bal = bal[(bal.index >= start) & (bal.index < end)].copy()
    bal["Fetch_version"] = FETCH_VERSION
    return fuel, states, bal


def load(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    d.index = pd.to_datetime(pd.Index(d.index).astype(str), errors="coerce")
    return d[d.index.notna()].sort_index()


def save(path, fuel, states, balance):
    daily = fuel.reindex(columns=[f for f in FUELS + ["Nuclear"] if f in fuel.columns]).add_suffix("_MWh")
    daily["Total_MWh"] = daily.sum(axis=1, min_count=1)
    for s in STORAGE + list(CHARGE.values()):
        if s in fuel:
            daily[f"{s}_MWh"] = fuel[s]
    st = states.reindex(columns=[s for s in STATES if s in states.columns]).add_suffix("_MWh")
    for x in (daily, st):
        x.index.name = "date"
    notes = [
        "UNITS",
        "MWh per day = sum of AEMO's 5-minute unit SCADA MW / 12 (negative readings clipped to 0), day = market "
        "day in AEST (intervals ending 00:05 to 24:00).",
        "Daily: Hydro (excl. pumped storage), Gas, Wind, Solar (utility scale), Coal, Oil (diesel/kerosene), "
        "Bioenergy; Total_MWh is their sum. Battery_discharge_MWh and Pumped_hydro_MWh are storage output, kept "
        "out of Total_MWh (they shift energy rather than generate it). Battery_charge_MWh and "
        "Pumped_hydro_pumping_MWh are the energy they draw (scheduled-load DUIDs, or a bidirectional unit's negative "
        "output).",
        "States: generation (excl. storage) by NEM region, MWh per day.",
        "State balance: per NEM region and day, MWh - <STATE>_<Fuel> generation by fuel (incl. Battery_discharge, "
        "Pumped_hydro), <STATE>_Rooftop_solar (AEMO ROOFTOP_PV_ACTUAL estimate, half-hourly MW / 2), "
        "<STATE>_Net_imports (interconnector flow into the region; negative = net export; DISPATCHREGIONSUM "
        "NETINTERCHANGE with the sign reversed) and <STATE>_Operational_demand (TOTALDEMAND: demand met by "
        "scheduled, semi-scheduled and significant non-scheduled generation, excl. rooftop solar).",
        "",
        "COVERAGE",
        f"National Electricity Market (QLD, NSW incl. ACT, VIC, SA, TAS), daily, {daily.index.min():%d %b %Y} to "
        f"{daily.index.max():%d %b %Y}. Not covered: Western Australia (separate WEM market - see "
        "au_wem_power_generation_daily.xlsx), Northern Territory, rooftop solar (not metered in unit SCADA). "
        "Units retired before the current registration list are not mapped to a fuel and are left out.",
        "",
        "SOURCE",
        "AEMO NEMWEB / MMSDM DISPATCH_UNIT_SCADA (via nemosis), unit fuel types from the NEM Registration and "
        f"Exemption List: {aemo_registration.URL}",
        "https://nemweb.com.au/",
    ]
    bal = balance.reindex(columns=sorted(c for c in balance.columns if c != "Fetch_version")).add_suffix("_MWh")
    if "Fetch_version" in balance:
        bal["Fetch_version"] = balance["Fetch_version"]
    bal.index.name = "date"
    xlsx_notes.write_workbook(path, {"Daily": daily.round(1), "States": st.round(1), "State balance": bal.round(1)},
                              notes,
                              {"UNITS", "COVERAGE", "SOURCE"})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--max-months", type=int, default=36)
    args = ap.parse_args()

    raw = load(args.out, "Daily")
    fuel = raw[[c for c in raw.columns if c != "Total_MWh"]].rename(columns=lambda c: c[:-4]) if not raw.empty else raw
    states = load(args.out, "States")
    states = states.rename(columns=lambda c: c[:-4]) if not states.empty else states
    balance = load(args.out, "State balance")
    balance = balance.rename(columns=lambda c: c[:-4] if c.endswith("_MWh") else c) if not balance.empty else balance
    today = pd.Timestamp(date.today())
    months = pd.date_range(HISTORY_START, today, freq="MS")
    # a month counts as saved once its state balance is saved too (the balance sheet was added after the first
    # backfill, so earlier months are fetched once more to fill it)
    # (and once more for storage CHARGING, added later: a month counts once SA - which has had batteries throughout -
    # has its battery charging saved)
    done = balance[balance["Fetch_version"] >= FETCH_VERSION] if "Fetch_version" in balance else balance.iloc[0:0]
    have = sorted(set(done.index.to_period("M").to_timestamp()))
    todo = [m for m in months if m not in have] + have[-1:]
    todo = sorted(set(todo))[:args.max_months]
    print(f"{len(have)} months saved; fetching {len(todo)}: {[f'{m:%Y-%m}' for m in todo]}", flush=True)
    duid_map = aemo_registration.units(include_loads=True)
    duid_map = aemo_registration.with_history(duid_map, pd.Timestamp(HISTORY_START), today, _table)
    duid_map = duid_map[["duid", "region", "fuel", "role"]]
    duid_map["fuel"] = duid_map["fuel"].replace(RENAME)
    for m in todo:
        end = min(m + pd.offsets.MonthBegin(1), today)
        try:
            f, s, b = fetch_month(duid_map, m, end)
        except Exception as e:  # noqa: BLE001 - keep the months already done
            print(f"  {m:%Y-%m}: failed {type(e).__name__}: {str(e)[:200]}", flush=True)
            continue
        if f is None:
            print(f"  {m:%Y-%m}: no rows", flush=True)
            continue
        keep = lambda d: d[d.index.to_period("M") != m.to_period("M")] if not d.empty else d  # noqa: E731
        fuel = pd.concat([keep(fuel), f]).sort_index()
        states = pd.concat([keep(states), s]).sort_index()
        balance = pd.concat([keep(balance), b]).sort_index()
        print(f"  {m:%Y-%m}: {len(f)} days, {f.drop(columns=STORAGE, errors='ignore').sum().sum() / 1e3:,.0f} GWh",
              flush=True)
        save(args.out, fuel, states, balance)
    if fuel.empty:
        raise SystemExit("No NEM generation data")
    print(f"Saved {args.out}: {len(fuel)} days", flush=True)


if __name__ == "__main__":
    main()
