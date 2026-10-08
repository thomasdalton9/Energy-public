"""
Daily rainfall at one representative point per key South / Central American hydro catchment, from NASA POWER.

    python3 RAINFALL_NASA_POWER.py [--out "output/Data and Chart Outputs/south_america_rainfall_daily.xlsx"]

Source: NASA POWER daily point API (https://power.larc.nasa.gov/api/temporal/daily/point), parameter PRECTOTCORR
(bias-corrected precipitation, mm/day; NASA's blend of MERRA-2 reanalysis and IMERG satellite rainfall, 0.5 x 0.625 degree
cell). Free, no key. It is a satellite/reanalysis-derived value for the grid cell around ONE point of each catchment -
NOT a rain-gauge record and NOT a basin average. NASA POWER lags real time by a few days.

The points are the reservoirs / dams the South and Central America master already charts for hydro (Brazil ONS subsystems,
Colombia, Ecuador Paute, Peru Mantaro, Chile, Argentina Comahue, Uruguay Rio Negro, Parana at Itaipu, Panama Gatun); a
catchment the master has no hydro data for (e.g. Venezuela Guri/Caroni) is not included. Latitude/longitude are the
approximate site of the dam or lake rounded to 0.1 degree (the NASA cell is far larger, so the rounding does not matter);
the basis of each point is on the 'Points' sheet.

Sheets: Units (notes), Daily (mm per day per point, from 2018-10-01), Monthly (mm per month, complete months only),
Points (catchment, latitude, longitude, basis). add_charts.py adds one 'Rainfall since 1 Oct' water-year sheet per point
(cumulative mm since 1 October: 5-year min-max band, 5-year average, previous and current water year).

Incremental: the workbook is the history store. Each run re-fetches each point from (last saved day - 10 days) to today
(NASA revises recent days); a point not yet in the workbook is fetched in full from 2018-10-01. Days NASA has not yet
published (-999) are left blank, never filled.
"""
import argparse
import datetime as dt
import os
import sys
import time

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

OUT = os.path.join(os.path.dirname(HERE), "output", "Data and Chart Outputs", "south_america_rainfall_daily.xlsx")
URL = "https://power.larc.nasa.gov/api/temporal/daily/point"
START = dt.date(2018, 10, 1)
REVISION_DAYS = 10
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept": "application/json"}

# key, country, catchment label, latitude, longitude, basis (the reservoir workbook / dashboard chart it matches)
POINTS = [
    ("BR_SECO", "Brazil", "SE/CO - Grande (Furnas)", -20.7, -46.3,
     "Furnas dam/lake, Rio Grande, Minas Gerais (Grande-Paranaiba basin, the core of the ONS SE/CO subsystem)"),
    ("BR_S", "Brazil", "South - Iguacu (Foz do Areia)", -26.0, -51.7,
     "Foz do Areia dam, Rio Iguacu, Parana (the largest storage of the ONS South subsystem)"),
    ("BR_NE", "Brazil", "Northeast - Sao Francisco (Tres Marias)", -18.2, -45.3,
     "Tres Marias dam, upper Rio Sao Francisco, Minas Gerais: the basin's rain falls in its upper reaches, which feed "
     "Sobradinho (the semi-arid Sobradinho site itself is not representative of the inflow)"),
    ("BR_N", "Brazil", "North - Tocantins (Tucurui)", -3.8, -49.7,
     "Tucurui dam, Rio Tocantins, Para (the main storage of the ONS North subsystem)"),
    ("CO", "Colombia", "Antioquia (Guatape)", 6.3, -75.2,
     "Guatape / El Penol lake, Antioquia (San Carlos - Guatape cascade; Colombia national reservoir chart)"),
    ("EC", "Ecuador", "Paute (Mazar)", -2.6, -78.6,
     "Mazar dam, Rio Paute, Azuay (the Mazar / Amaluza reservoirs charted on the Ecuador hydro workbook)"),
    ("PE", "Peru", "Mantaro (Lake Junin)", -11.0, -76.1,
     "Lake Junin (Chinchaycocha), head of the Rio Mantaro (Junin and Mantaro lagoons in the Peru COES reservoir series)"),
    ("CL_LAJA", "Chile", "Laja (Lago Laja)", -37.4, -71.3,
     "Laguna del Laja, Biobio region (Lago Laja volume, Chile DGA reservoir series)"),
    ("CL_MAULE", "Chile", "Maule (Laguna del Maule)", -36.1, -70.5,
     "Laguna del Maule, Maule region (Laguna del Maule volume, Chile DGA reservoir series)"),
    ("AR_COMAHUE", "Argentina", "Comahue - Limay (Alicura)", -40.6, -70.8,
     "Alicura dam, Rio Limay, Neuquen/Rio Negro (Comahue lakes charted on the Argentina hydro workbook; the Andean "
     "headwaters, not the arid lower valley where El Chocon sits, carry the rain)"),
    ("UY", "Uruguay", "Rio Negro (Bonete)", -32.8, -56.4,
     "Rincon del Bonete (Gabriel Terra) lake, Rio Negro (Bonete level, Uruguay ADME series)"),
    ("PY_PARANA", "Paraguay-Brazil", "Parana (Itaipu)", -25.4, -54.6,
     "Itaipu dam, Rio Parana on the Paraguay-Brazil border (Paraguay power, Parana flow at Yacyreta on the Argentina "
     "workbook); a point at the dam, the Parana catchment above it is far larger"),
    ("PA", "Panama", "Gatun / Canal watershed", 9.2, -79.9,
     "Gatun Lake, Panama Canal (Gatun Lake level chart)"),
]


def fetch_point(lat, lon, start, end, tries=3):
    """Daily PRECTOTCORR (mm/day) as a Series; NASA's -999 (not yet available) -> NaN."""
    params = {"parameters": "PRECTOTCORR", "community": "AG", "latitude": lat, "longitude": lon,
              "start": start.strftime("%Y%m%d"), "end": end.strftime("%Y%m%d"), "format": "JSON"}
    last = None
    for i in range(tries):
        try:
            r = requests.get(URL, params=params, headers=UA, timeout=(15, 90))
            if r.status_code == 429:
                time.sleep(30 * (i + 1))
                continue
            r.raise_for_status()
            p = r.json()["properties"]["parameter"]["PRECTOTCORR"]
            s = pd.Series({pd.Timestamp(k): float(v) for k, v in p.items()})
            return s.where(s > -998)
        except Exception as exc:  # noqa: BLE001
            last = exc
            print(f"    attempt {i + 1} failed: {type(exc).__name__}: {exc}", flush=True)
            time.sleep(5 * (i + 1))
    raise RuntimeError(f"{lat},{lon}: {last}")


def load_daily(path):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        d = pd.read_excel(path, sheet_name="Daily")
    except Exception as exc:  # noqa: BLE001
        print(f"existing workbook unreadable ({exc}); rebuilding in full", flush=True)
        return pd.DataFrame()
    d = d.drop(columns=[c for c in d.columns if str(c).startswith("Unnamed")])
    d["date"] = pd.to_datetime(d["date"])
    return d.set_index("date").sort_index()


def monthly_totals(daily):
    """mm per month for complete months only (every day present)."""
    m = daily.resample("MS").sum(min_count=1)
    n = daily.resample("MS").count()
    days = pd.Series(m.index.days_in_month, index=m.index)
    return m.where(n.eq(days, axis=0))


def write(path, daily, fetched_to):
    pts = pd.DataFrame([{"column": k, "country": c, "catchment": n, "latitude": la, "longitude": lo, "basis": b}
                        for k, c, n, la, lo, b in POINTS]).set_index("column")
    last = {k: (daily[k].dropna().index.max().date().isoformat() if daily[k].notna().any() else "none") for k in daily}
    notes = [
        "UNITS",
        "Daily: mm of rain per day at one point of each catchment (NASA POWER PRECTOTCORR, bias-corrected precipitation). "
        "Blank = NASA has not published that day yet; nothing is filled.",
        "Monthly: mm per month, only months in which every day is present.",
        "Water-year sheets (added by add_charts.py): CUMULATIVE mm since 1 October (Oct-Sep water year); 5-year band/average "
        "= the five water years before the current one; blank from the first missing day of a water year.",
        "",
        "COVERAGE",
        f"Daily from {daily.index.min().date()} to {daily.dropna(how='all').index.max().date()} "
        f"(NASA POWER lags real time by a few days; fetched on {fetched_to}).",
        "Points: " + "; ".join(f"{k} = {n} ({la}, {lo})" for k, _, n, la, lo, _ in POINTS),
        "",
        "READ THIS BEFORE USING",
        "NASA POWER is a satellite / reanalysis-derived product (MERRA-2 and IMERG, bias-corrected) on a 0.5 x 0.625 degree "
        "grid. Each column is the value for the grid cell around ONE point, NOT a rain gauge and NOT an average over the "
        "catchment. Mountain catchments (Andes) are poorly resolved and totals can differ markedly from gauges. Use it for "
        "'wetter or drier than the last five years', not for absolute basin rainfall.",
        "Latitude/longitude are the approximate site of each dam / lake rounded to 0.1 degree (see the Points sheet for the "
        "basis of each); they are not surveyed coordinates.",
        "",
        "SOURCE",
        "NASA Langley Research Center POWER Project, daily point API https://power.larc.nasa.gov/api/temporal/daily/point "
        "(parameter PRECTOTCORR, community AG). Free, no key.",
        "Script: south_america/RAINFALL_NASA_POWER.py (scheduled by .github/workflows/south_america_rainfall.yml, 1st and 15th).",
        "",
        "METHOD",
        f"Incremental: the workbook is the history store. Each run re-fetches every point from its last saved day minus "
        f"{REVISION_DAYS} days (NASA revises recent days); a point not yet saved is fetched in full from {START}.",
        "",
        "LATEST DAY SAVED PER POINT",
    ] + [f"  {k}: {v}" for k, v in last.items()]
    out = daily.copy()
    out.index.name = "date"
    mon = monthly_totals(daily).round(1)
    mon.index.name = "month"
    xlsx_notes.write_workbook(path, {"Daily": out.round(2), "Monthly": mon, "Points": pts}, notes,
                              {"UNITS", "COVERAGE", "READ THIS BEFORE USING", "SOURCE", "METHOD", "LATEST DAY SAVED PER POINT"})


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start", type=dt.date.fromisoformat, default=START)
    ap.add_argument("--budget-min", type=float, default=12, help="stop fetching further points after this many minutes")
    args = ap.parse_args()
    today = dt.date.today()
    daily = load_daily(args.out)
    failed = []
    t0 = time.time()
    for key, country, name, lat, lon, _ in POINTS:
        have = daily[key].dropna() if key in daily else pd.Series(dtype=float)
        start = (have.index.max().date() - dt.timedelta(days=REVISION_DAYS)) if len(have) else args.start
        start = max(start, args.start)
        if time.time() - t0 > args.budget_min * 60:
            print(f"  {key}: skipped, time budget used up", flush=True)
            failed.append(key)
            continue
        print(f"  {key}: fetching from {start} ...", flush=True)
        try:
            s = fetch_point(lat, lon, start, today).dropna()
        except Exception as exc:  # noqa: BLE001
            print(f"  {key}: FAILED {exc}", flush=True)
            failed.append(key)
            continue
        if key not in daily:
            daily[key] = float("nan")
        daily = daily.reindex(daily.index.union(s.index))
        daily.loc[s.index, key] = s.values
        print(f"  {key} {name}: {len(s)} days {s.index.min().date()} to {s.index.max().date()} "
              f"(saved {daily[key].notna().sum()})", flush=True)
        time.sleep(1)
    if daily.empty:
        sys.exit("no data and no saved workbook")
    daily = daily[[k for k, *_ in POINTS if k in daily]].sort_index()
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    write(args.out, daily, today)
    print(f"wrote {args.out}: {daily.index.min().date()} to {daily.dropna(how='all').index.max().date()}, "
          f"{len(failed)} point(s) failed {failed}")
    if len(failed) == len(POINTS):
        sys.exit(1)


if __name__ == "__main__":
    main()
