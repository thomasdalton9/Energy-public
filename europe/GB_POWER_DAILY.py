"""
Great Britain power generation, load and interconnector flows from Elexon (BMRS) and NESO, one workbook in the same
daily layout as the ENTSO-E country workbooks:

  output/Data and Chart Outputs/great_britain_power_generation_daily.xlsx
    sheet "Daily": date, <Fuel>_MWh (Hydro, PumpedStorage, Gas, Coal, Oil, Nuclear, Wind, Solar, Bioenergy, Other, Storage),
                   Total_MWh, PumpedStorageConsumption_MWh, StorageCharging_MWh, Load_MWh, NetImports_MWh
    sheet "Units": source and definitions

GB left the ENTSO-E Transparency Platform's generation reporting after Brexit, so it is pulled from the system operators:
  - Elexon BMRS FUELHH (half-hourly metered generation by fuel, MW, incl. the interconnector flows INT*; free, no key)
  - NESO "Historic demand data" (half-hourly national demand ND, embedded wind and solar generation, pumped-storage
    pumping; free, no key) - the embedded wind/solar are not in FUELHH and are large (about 20 GW solar capacity)
Definitions
  - Hydro = NPSHYD (non-pumped hydro); PumpedStorage = PS output; Gas = CCGT + OCGT; Coal; Oil; Nuclear;
    Bioenergy = BIOMASS; Other = OTHER; Wind = metered WIND + NESO embedded wind; Solar = NESO embedded solar
    (GB solar is almost all embedded, so it is not in FUELHH).
  - Load_MWh = national demand (ND) + embedded wind + embedded solar (the demand GB generators and imports meet before
    station load); PumpedStorageConsumption_MWh = NESO pumping.
  - NetImports_MWh = sum of the interconnector flows (INTFR, INTIRL, INTNED, INTEW, INTNEM, INTELEC, INTIFA2, INTNSL,
    INTVKL, INTGRNL), positive = import into GB.
  - MWh = MW x 0.5 h per half hour; days are UTC days; a day is kept only if both feeds cover >= 46 of 48 half hours.

Incremental: reads the committed workbook, re-fetches the last 10 days plus any gaps; first run backfills from 2021-01-01.

Usage: python3 GB_POWER_DAILY.py [--out-dir DIR] [--start 2021-01-01]
"""
import argparse
import io
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DEFAULT = os.path.join(ROOT, "output", "Data and Chart Outputs")
FILE = "great_britain_power_generation_daily.xlsx"
ELEXON = "https://data.elexon.co.uk/bmrs/api/v1"
NESO = "https://api.neso.energy/api/3/action"
REVISION_DAYS = 10
MIN_PERIODS = 46
WINDOW_DAYS = 30

GEN_FUELS = ["Hydro", "PumpedStorage", "Gas", "Coal", "Oil", "Nuclear", "Wind", "Solar", "Bioenergy", "Other", "Storage"]
GEN_COLS = [f"{f}_MWh" for f in GEN_FUELS]
ALL_COLS = GEN_COLS + ["Total_MWh", "PumpedStorageConsumption_MWh", "StorageCharging_MWh", "Load_MWh", "NetImports_MWh"]
FUEL_MAP = {"NPSHYD": "Hydro", "PS": "PumpedStorage", "CCGT": "Gas", "OCGT": "Gas", "COAL": "Coal", "OIL": "Oil",
            "NUCLEAR": "Nuclear", "WIND": "Wind", "BIOMASS": "Bioenergy", "OTHER": "Other"}
INTERCONNECTORS = {"INTFR", "INTIRL", "INTNED", "INTEW", "INTNEM", "INTELEC", "INTIFA2", "INTNSL", "INTVKL", "INTGRNL"}

S = requests.Session()
S.headers["Accept"] = "application/json"


def get(url, tries=5, **params):
    for i in range(tries):
        try:
            r = S.get(url, params=params, timeout=120)
            if r.status_code == 200:
                return r
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(3 * (i + 1))
                continue
            raise RuntimeError(f"{r.status_code} {r.text[:200]} for {r.url}")
        except requests.RequestException:
            time.sleep(3 * (i + 1))
    raise RuntimeError(f"gave up on {url} {params}")


def fuelhh(d0, d1):
    """Half-hourly FUELHH for settlement dates d0..d1 -> frame indexed by UTC start time, columns = fuel groups (MW)."""
    rows = get(f"{ELEXON}/datasets/FUELHH/stream", settlementDateFrom=d0.isoformat(), settlementDateTo=d1.isoformat(),
               format="json").json()
    if not rows:
        return pd.DataFrame()
    d = pd.DataFrame(rows)[["startTime", "fuelType", "generation", "publishTime"]]
    d["startTime"] = pd.to_datetime(d["startTime"], utc=True)
    d = d.sort_values("publishTime").drop_duplicates(["startTime", "fuelType"], keep="last")
    d["generation"] = pd.to_numeric(d["generation"], errors="coerce")
    mapped = d[d.fuelType.isin(FUEL_MAP)].assign(group=lambda x: x.fuelType.map(FUEL_MAP))
    out = mapped.pivot_table(index="startTime", columns="group", values="generation", aggfunc="sum")
    net = d[d.fuelType.isin(INTERCONNECTORS)].pivot_table(index="startTime", values="generation", aggfunc="sum")
    out["NetImports"] = net["generation"]
    # periods where a mapped fuel is entirely absent (e.g. a half hour not yet published) are not complete
    out["_n"] = d.groupby("startTime")["fuelType"].nunique()
    return out


_resources = {}


def neso_resource(year):
    """Resource id of NESO's "Historic Demand Data <year> CSV" (resolved at run time, ids change when NESO re-publishes)."""
    if not _resources:
        for r in get(f"{NESO}/package_show", id="historic-demand-data").json()["result"]["resources"]:
            parts = r.get("name", "").split()
            if parts[:3] == ["Historic", "Demand", "Data"] and len(parts) >= 4 and parts[3].isdigit():
                _resources[int(parts[3])] = r["id"]
    return _resources.get(year)


def neso_year(year):
    """Half-hourly NESO demand for a calendar year -> frame by UTC start time (MW)."""
    rid = neso_resource(year)
    if not rid:
        return pd.DataFrame()
    rows, offset = [], 0
    while True:
        j = get(f"{NESO}/datastore_search", resource_id=rid, limit=32000, offset=offset).json()["result"]
        rows += j["records"]
        offset += len(j["records"])
        if not j["records"] or offset >= j.get("total", 0):
            break
    if not rows:
        return pd.DataFrame()
    d = pd.DataFrame(rows)
    local_midnight = pd.to_datetime(d["SETTLEMENT_DATE"]).dt.tz_localize("Europe/London", ambiguous="NaT", nonexistent="NaT")
    d["startTime"] = (local_midnight.dt.tz_convert("UTC") + pd.to_timedelta((d["SETTLEMENT_PERIOD"] - 1) * 30, unit="m"))
    d = d.dropna(subset=["startTime"]).drop_duplicates("startTime", keep="last").set_index("startTime")
    cols = ["ND", "EMBEDDED_WIND_GENERATION", "EMBEDDED_SOLAR_GENERATION", "PUMP_STORAGE_PUMPING"]
    return d[cols].apply(pd.to_numeric, errors="coerce")


def daily(f, n):
    """Half-hourly MW frames -> UTC-day MWh in the standard columns; days with < MIN_PERIODS periods dropped."""
    if f.empty or n.empty:
        return pd.DataFrame(columns=ALL_COLS)
    j = f.join(n, how="inner")
    j = j[j.index.notna()]
    wind = j.get("Wind", 0).fillna(0) + j["EMBEDDED_WIND_GENERATION"].fillna(0)
    out = pd.DataFrame(index=j.index)
    for fuel in GEN_FUELS:
        out[f"{fuel}_MWh"] = j[fuel] if fuel in j else 0.0
    out["Wind_MWh"] = wind
    if "PumpedStorage" in j:   # FUELHH "PS" goes negative while pumping; output only here, pumping comes from NESO
        out["PumpedStorage_MWh"] = j["PumpedStorage"].clip(lower=0)
    out["Solar_MWh"] = j["EMBEDDED_SOLAR_GENERATION"]
    out["Storage_MWh"] = 0.0
    out["Load_MWh"] = j["ND"] + j["EMBEDDED_WIND_GENERATION"] + j["EMBEDDED_SOLAR_GENERATION"]
    out["PumpedStorageConsumption_MWh"] = j["PUMP_STORAGE_PUMPING"]
    out["StorageCharging_MWh"] = 0.0
    out["NetImports_MWh"] = j["NetImports"]
    out["Total_MWh"] = out[GEN_COLS].sum(axis=1)
    out = out.fillna(0.0) * 0.5   # MW per half hour -> MWh
    days = out.groupby(out.index.floor("D")).sum()
    n_per = j["ND"].notna().groupby(j.index.floor("D")).sum()
    days = days[n_per.reindex(days.index) >= MIN_PERIODS]
    days.index = pd.DatetimeIndex(days.index.tz_localize(None), name="date")
    return days[ALL_COLS]


def read_existing(path):
    if not os.path.exists(path):
        return pd.DataFrame(columns=ALL_COLS)
    try:
        d = pd.read_excel(path, sheet_name="Daily")
    except Exception as e:  # noqa: BLE001
        print(f"could not read {FILE} ({type(e).__name__}: {e}); starting over")
        return pd.DataFrame(columns=ALL_COLS)
    d["date"] = pd.to_datetime(d["date"], errors="coerce")
    d = d.dropna(subset=["date"]).set_index("date").sort_index()
    for c in ALL_COLS:
        if c not in d:
            d[c] = float("nan")
    return d[ALL_COLS]


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
    have = set(old.index.date) if len(old) else set()
    fs = start
    if have:
        fs = max(start, max(have) - timedelta(days=REVISION_DAYS))
        gaps = [start + timedelta(days=i) for i in range((today - start).days)]
        gaps = [g for g in gaps if g >= today - timedelta(days=120) and g not in have and g < fs]
        if gaps:
            fs = min(fs, gaps[0])
    print(f"fetching {fs} .. {today - timedelta(days=1)}", flush=True)
    neso = pd.concat([neso_year(y) for y in range(fs.year, today.year + 1)]) if fs <= today else pd.DataFrame()
    neso = neso[~neso.index.duplicated(keep="last")] if not neso.empty else neso
    parts, d0 = [], fs
    while d0 < today:
        d1 = min(d0 + timedelta(days=WINDOW_DAYS - 1), today - timedelta(days=1))
        f = fuelhh(d0 - timedelta(days=1), d1 + timedelta(days=1))   # +-1 day: settlement dates are local, days here are UTC
        new = daily(f, neso)
        new = new[(new.index.date >= d0) & (new.index.date <= d1)]
        parts.append(new)
        print(f"  {d0} .. {d1}: {len(new)} days", flush=True)
        d0 = d1 + timedelta(days=1)
        time.sleep(0.3)
    new = pd.concat(parts) if parts else pd.DataFrame(columns=ALL_COLS)
    combined = old[~old.index.isin(new.index)]
    combined = pd.concat([combined, new]) if len(combined) else new
    combined = combined.sort_index().round(1)
    combined.index.name = "date"
    if combined.empty:
        print("no GB data")
        return
    lines = ["Great Britain - power generation, load and interconnector flows (Elexon BMRS + NESO)", "",
             "Source", "Elexon BMRS FUELHH (half-hourly metered generation by fuel and interconnector flows), "
             "https://bmrs.elexon.co.uk/ ; NESO Historic demand data (national demand, embedded wind and solar, pumping), "
             "https://www.neso.energy/data-portal/historic-demand-data . Free, no key.",
             "", "Units and definitions",
             "MWh per UTC day (MW x 0.5 h per half hour). Hydro = non-pumped hydro (NPSHYD); PumpedStorage = pumped storage output (PS, positive half hours only); "
             "Gas = CCGT + OCGT; Bioenergy = BIOMASS; Wind = metered wind + NESO embedded wind; Solar = NESO embedded solar "
             "(almost all GB solar is embedded and not in FUELHH).",
             "Load_MWh = national demand (ND) + embedded wind + embedded solar. PumpedStorageConsumption_MWh = pumping. "
             "NetImports_MWh = sum of interconnector flows, positive = import into GB. Total_MWh = all generation columns.",
             f"A day is kept only if both feeds cover >= {MIN_PERIODS} of 48 half hours. Re-fetches the last {REVISION_DAYS} days each run plus gaps "
             f"within 120 days; history from {args.start}.",
             "", "Last pull", f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC; {len(combined)} days, "
             f"{combined.index.min():%Y-%m-%d} to {combined.index.max():%Y-%m-%d}"]
    xlsx_notes.write_workbook(path, {"Daily": combined}, lines, {"Source", "Units and definitions", "Last pull"})
    print(f"saved {FILE}: {len(combined)} days")


if __name__ == "__main__":
    main()
