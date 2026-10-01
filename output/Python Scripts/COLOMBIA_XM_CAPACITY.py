"""
Colombia installed generation capacity by technology, monthly from January
2021, from XM's public API (servapibi.xm.com.co, no key).

  /daily  MetricId=CapEfecNeta, Entity=Recurso: each plant's net effective
          capacity ("Capacidad Efectiva Neta", kW) for each day. XM: "the
          maximum net power a generating unit can deliver under normal
          operating conditions; includes small plants (menores) and
          cogenerators".
  /lists  MetricId=ListadoRecursos, Entity=Sistema: every plant's code,
          name, type and fuel (EnerSource), used to map plants to fuels.

Found via discovery_archive/south_america/POWER_CAPACITY_PROBE.py.

Monthly value = the capacity XM reports for the last day of each month
(month-end); the current month uses the latest day XM has published.
Incremental: months already saved are kept; only missing months plus the
latest two are fetched (XM can revise recent days).

Usage: python3 COLOMBIA_XM_CAPACITY.py [--out <xlsx>] [--months N]
  --months N  fetch at most N missing months this run (test runs)
"""
import argparse
import datetime as dt
import os
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import power_capacity_std as std  # noqa: E402

OUT = "output/Data and Chart Outputs/colombia_power_capacity.xlsx"
DAILY_URL = "https://servapibi.xm.com.co/daily"
LISTS_URL = "https://servapibi.xm.com.co/lists"
HEADERS = {"Connection": "close"}
REFRESH_MONTHS = 2

# XM EnerSource -> standard fuel
FUEL_MAP = {
    "AGUA": "Hydro",
    "GAS": "Gas", "GLP": "Gas",
    "CARBON": "Coal",
    "COMBUSTOLEO": "Oil", "ACPM": "Oil", "JET-A1": "Oil",
    "VIENTO": "Wind",
    "RAD SOLAR": "Solar",
    "BAGAZO": "Bioenergy", "BIOGAS": "Bioenergy", "BIOMASA": "Bioenergy",
}


def post(url, body, tries=4):
    for attempt in range(tries):
        try:
            r = requests.post(url, json=body, headers=HEADERS, timeout=(20, 300))
            r.raise_for_status()
            return r.json()
        except (requests.RequestException, ValueError) as e:
            if attempt == tries - 1:
                raise
            print(f"  retrying after {type(e).__name__}: {e}", flush=True)
            time.sleep(10 * (attempt + 1))


def plant_list():
    payload = post(LISTS_URL, {"MetricId": "ListadoRecursos", "Entity": "Sistema"})
    rows = [e.get("Values", {}) for it in payload.get("Items", []) for e in it.get("ListEntities", [])]
    df = pd.DataFrame(rows)
    df = df[df["Code"].notna()].drop_duplicates("Code").set_index("Code")
    df["EnerSource"] = df["EnerSource"].fillna("").astype(str).str.strip().str.upper()
    print(f"  {len(df):,} resources in XM's plant list", flush=True)
    return df


def capacity_on(day):
    """{plant code: MW} for one day; empty if XM has nothing for that day."""
    payload = post(DAILY_URL, {"MetricId": "CapEfecNeta", "Entity": "Recurso",
                               "StartDate": day.isoformat(), "EndDate": day.isoformat()})
    out = {}
    for it in payload.get("Items", []):
        for e in it.get("DailyEntities", []):
            try:
                out[e["Code"]] = float(e["Value"]) / 1000.0  # kW -> MW
            except (KeyError, TypeError, ValueError):
                continue
    return out


def month_end_value(month, today):
    """(day used, {code: MW}) for the month's last day, stepping back if XM hasn't published it yet."""
    last = (month + pd.offsets.MonthEnd(0)).date()
    day = min(last, today - dt.timedelta(days=1))
    for _ in range(12):
        if day < month.date():
            break
        vals = capacity_on(day)
        if vals:
            return day, vals
        day -= dt.timedelta(days=1)
    return None, {}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--months", type=int, default=0)
    args = ap.parse_args()

    today = dt.date.today()
    plants = plant_list()
    existing = std.load_sheet(args.out, "By plant monthly", index_col=0)
    if not existing.empty:
        existing.index = pd.to_datetime(existing.index)
        existing.columns = existing.columns.astype(str)
    months = pd.date_range(std.START, pd.Timestamp(today).replace(day=1), freq="MS")
    have = set(existing.index) if not existing.empty else set()
    todo = [m for m in months if m not in have]
    todo += [m for m in sorted(have)[-REFRESH_MONTHS:] if m not in todo]
    todo = sorted(set(todo))
    if args.months:
        todo = todo[-args.months:]
    print(f"{len(have)} month(s) saved; fetching {len(todo)}: "
          f"{', '.join(m.strftime('%Y-%m') for m in todo)}", flush=True)

    rows, days_used = {}, {}
    for m in todo:
        day, vals = month_end_value(m, today)
        if not vals:
            print(f"  {m:%Y-%m}: no data", flush=True)
            continue
        rows[m] = vals
        days_used[m] = day
        print(f"  {m:%Y-%m}: {day} {len(vals)} plants {sum(vals.values()):,.0f} MW", flush=True)
        time.sleep(0.5)
    new = pd.DataFrame.from_dict(rows, orient="index")
    by_plant = existing.copy() if not existing.empty else pd.DataFrame()
    if not new.empty:
        by_plant = by_plant.drop(index=[i for i in new.index if i in by_plant.index], errors="ignore")
        by_plant = pd.concat([by_plant, new]).sort_index()
    if by_plant.empty:
        raise SystemExit("No capacity data from XM")
    if "day_used" in by_plant.columns:
        by_plant = by_plant.drop(columns=["day_used"])
    by_plant.index.name = "date"

    fuel_of = {c: FUEL_MAP.get(plants["EnerSource"].get(c, ""), "Other") for c in by_plant.columns}
    unmapped = sorted({plants["EnerSource"].get(c, "(not in XM plant list)") for c in by_plant.columns
                       if fuel_of[c] == "Other"})
    if unmapped:
        print(f"  counted as Other: {unmapped}", flush=True)
    by_fuel = by_plant.T.groupby(pd.Series(fuel_of)).sum().T
    monthly = std.standard(by_fuel)

    by_label = by_plant.T.groupby(pd.Series({c: plants["EnerSource"].get(c, "(not listed)") for c in by_plant.columns})
                                  ).sum().T.round(1)
    by_label.index.name = "date"
    latest = by_plant.iloc[-1].dropna()
    latest = latest[latest > 0]
    plant_tab = pd.DataFrame({
        "name": [plants["Name"].get(c) for c in latest.index],
        "xm_type": [plants["Type"].get(c) for c in latest.index],
        "xm_fuel": [plants["EnerSource"].get(c) for c in latest.index],
        "fuel": [fuel_of[c] for c in latest.index],
        "dispatch": [plants["Disp"].get(c) for c in latest.index],
        "class": [plants["RecType"].get(c) for c in latest.index],
        "start": [plants["OperStartdate"].get(c) for c in latest.index],
        "capacity_MW": latest.round(3).values,
    }, index=pd.Index(latest.index, name="code")).sort_values(["fuel", "capacity_MW"], ascending=[True, False])
    check = std.ember_check("Colombia", monthly)
    if not check.empty:
        check = check.set_index("year")

    last_month = monthly.index.max()
    notes = [
        "UNITS",
        "MW of net effective capacity (Capacidad Efectiva Neta) at month end. 'date' is the 1st of the month; the value "
        "is the capacity XM reports for that month's last day (or, for the current month, the latest day published).",
        "",
        "COVERAGE",
        f"Monthly, {monthly.index.min():%b %Y} to {last_month:%b %Y}. National interconnected system (SIN) plants that "
        "XM registers with a net effective capacity, including small plants and cogenerators. Behind-the-meter and "
        "small self-generation without an XM capacity record (most rooftop solar) is not included.",
        "",
        "SOURCE",
        "XM (Colombian system operator), public API https://servapibi.xm.com.co/daily with MetricId=CapEfecNeta, "
        "Entity=Recurso (kW per plant per day); plant fuels from https://servapibi.xm.com.co/lists, "
        "MetricId=ListadoRecursos. Updated weekly by GitHub Actions (colombia_power_capacity.yml); months already "
        "saved are kept and only missing months plus the latest two are re-fetched.",
        "",
        "MAPPING (XM EnerSource -> column)",
        "Hydro_MW = AGUA (reservoir, run-of-river and small hydro; Colombia has no pumped storage). Gas_MW = GAS, GLP. "
        "Coal_MW = CARBON. Oil_MW = COMBUSTOLEO (fuel oil), ACPM (diesel), JET-A1. Wind_MW = VIENTO. "
        "Solar_MW = RAD SOLAR. Bioenergy_MW = BAGAZO, BIOGAS, BIOMASA. Nuclear_MW = 0.",
        "Other_MW = plants whose code is missing from XM's plant list (unclassified); Colombia reports no geothermal "
        "or storage capacity in this metric. Total_MW = sum of the fuel columns.",
        "Dual-fuel thermal plants are counted under the fuel XM lists for them (e.g. gas plants that can burn "
        "liquids stay in Gas_MW).",
        "",
        "SHEETS",
        "Monthly: standard capacity table. By XM fuel: the same month-end MW by XM's own fuel label. By plant monthly: "
        "MW per plant code per month (the incremental archive). Latest by plant: plant detail for the last month. "
        "Ember check: this table's December value against Ember's yearly capacity (GW, % difference).",
    ]
    std.write(args.out, monthly, {"By XM fuel": by_label, "Latest by plant": plant_tab,
                                  "By plant monthly": by_plant.round(3), "Ember check": check},
              notes, {"UNITS", "COVERAGE", "SOURCE", "MAPPING (XM EnerSource -> column)", "SHEETS"})


if __name__ == "__main__":
    main()
