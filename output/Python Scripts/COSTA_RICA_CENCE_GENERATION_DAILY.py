"""
Costa Rica daily power generation by type from ICE / CENCE (Centro Nacional de
Control de Energia, DOCSE) - raw grid-operator pull in the standard layout.

Source: CENCE web service "EnergiaHorariaFuente" (the data behind
https://apps.grupoice.com/CenceWeb/paginas/GeneracionReal.html), one day per call:
  https://apps.grupoice.com/CenceWeb/data/sen/json/EnergiaHorariaFuente?anno=YYYY&mes=M&dia=D
-> hourly energy (MWh) by source: Hidro, Eolico, Geotermico, Solar, Bagazo, Bunker,
Diesel and Intercambio (net regional interchange). SCADA real-time data (CENCE:
"no son obtenidos de medicion comercial"; the current and previous month are
still being revised). History from 2021-01-01. Found via
discovery_archive/south_america/CENTRAL_AMERICA_POWER_DISCOVERY3.py (the
parameters are anno/mes/dia; inicio/fin are silently ignored and return today).

Each run fetches only days missing from the workbook plus the last 14 days
(revisions), within a time budget, saving as it goes.

Usage: python3 COSTA_RICA_CENCE_GENERATION_DAILY.py [--out PATH] [--start YYYY-MM-DD]
"""

print("STARTING", flush=True)

import argparse
import datetime as dt
import os
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import central_america_power_common as C  # noqa: E402

OUT = "output/Data and Chart Outputs/costa_rica_power_generation_daily.xlsx"
URL = "https://apps.grupoice.com/CenceWeb/data/sen/json/EnergiaHorariaFuente"
PAGE = "https://apps.grupoice.com/CenceWeb/paginas/GeneracionReal.html"

# CENCE 'Fuente' (accents/case ignored) -> standard fuel. None = not generation.
MAPPING = {
    "Hidro": "Hydro",
    "Eolico": "Wind",
    "Solar": "Solar",
    "Geotermico": "Other",      # geothermal (ICE Miravalles, Las Pailas, ...)
    "Bagazo": "Bioenergy",      # sugar-cane bagasse
    "Biomasa": "Bioenergy",
    "Bunker": "Oil",
    "Diesel": "Oil",
    "Termico": "Oil",
    "Gas": "Gas",
    "Intercambio": None,        # net regional (MER) interchange, not generation
}
KEY = {C.norm(k): v for k, v in MAPPING.items()}

NOTES = [
    "UNITS",
    "MWh per day: the sum of CENCE's 24 hourly energy values (MWh) for each source. Days with fewer "
    "than 23 hourly values are left out and fetched again on the next run.",
    "",
    "COVERAGE",
    "@COVERAGE",
    "",
    "SOURCE",
    "ICE / CENCE (DOCSE), Costa Rica - 'Generacion Real, MWh': " + PAGE,
    "Web service: " + URL + "?anno=YYYY&mes=M&dia=D (JSON, one day per call).",
    "SCADA real-time data, not commercial metering; CENCE revises the current and previous month. "
    "Updated daily by GitHub Actions (costa_rica_power_generation.yml): only missing days plus the last 14.",
    "",
    "MAPPING",
    "CENCE source -> column in sheet 'Daily':",
    "Hidro -> Hydro_MWh",
    "Eolico -> Wind_MWh",
    "Solar -> Solar_MWh",
    "Geotermico -> Other_MWh (geothermal is put in Other_MWh; Costa Rica has no other 'other' source)",
    "Bagazo (sugar-cane bagasse) -> Bioenergy_MWh",
    "Bunker, Diesel (ICE thermal plants) -> Oil_MWh",
    "Intercambio (net regional interchange; negative = net export) -> not generation, left out of "
    "Total_MWh; shown in sheet 'Detail'.",
    "Costa Rica has no gas, coal or nuclear generation: those columns are 0.",
    "Total_MWh = sum of the fuel columns (domestic generation, interchange excluded).",
    "Sheet 'Detail': MWh per day for each CENCE source as published (incl. Intercambio).",
]


def fetch_day(session, day):
    r = session.get(URL, params={"anno": day.year, "mes": day.month, "dia": day.day}, timeout=(15, 60))
    r.raise_for_status()
    rows = r.json().get("data") or []
    df = pd.DataFrame(rows)
    if df.empty:
        return None
    df["Fecha"] = pd.to_datetime(df["Fecha"])
    df = df[df["Fecha"].dt.date == day]  # the service answers today's data if the date is not understood
    if df.empty or df["Fecha"].dt.hour.nunique() < 23:
        return None
    df["Dato"] = pd.to_numeric(df["Dato"], errors="coerce")
    return df.groupby("Fuente")["Dato"].sum(min_count=1)


def to_frames(per_day):
    detail = pd.DataFrame(per_day).T
    detail.index = pd.to_datetime(detail.index)
    detail.index.name = "date"
    fuels = pd.DataFrame(index=detail.index)
    unknown = []
    for src in detail.columns:
        k = C.norm(src)
        if k not in KEY:
            unknown.append(src)
            fuel = "Other"
        else:
            fuel = KEY[k]
        if fuel is None:
            continue
        fuels[fuel] = fuels.get(fuel, 0) + detail[src].fillna(0)
    if unknown:
        print(f"  WARNING: CENCE sources not in MAPPING, counted as Other: {unknown}", flush=True)
    return C.standardise(fuels), detail.round(1)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start", type=dt.date.fromisoformat, default=C.HISTORY_START)
    ap.add_argument("--refresh-days", type=int, default=14)
    ap.add_argument("--budget-min", type=float, default=40, help="stop fetching after this many minutes")
    ap.add_argument("--save-every", type=int, default=150, help="write the workbook every N days fetched")
    args = ap.parse_args()
    t0 = time.time()
    C.print_mapping(MAPPING, "CENCE Fuente -> standard fuel; None = excluded")

    daily = C.load_sheet(args.out, "Daily")
    detail = C.load_sheet(args.out, "Detail")
    end = dt.date.today() - dt.timedelta(days=1)
    days = sorted(C.days_to_fetch(daily, args.start, end, args.refresh_days), reverse=True)  # newest first
    print(f"{len(daily):,} days saved; fetching {len(days):,}", flush=True)

    s = requests.Session()
    s.headers.update({"User-Agent": C.USER_AGENT, "Referer": PAGE})
    per_day, failed = {}, []

    def flush():
        nonlocal daily, detail, per_day
        if per_day:
            new, new_detail = to_frames(per_day)
            daily, detail = C.merge(new, daily), C.merge(new_detail, detail)
            per_day = {}
        if not daily.empty:
            notes = [x for line in NOTES for x in (C.coverage_lines(daily) if line == "@COVERAGE" else [line])]
            C.write(args.out, daily, notes, detail)

    for k, day in enumerate(days, 1):
        if time.time() - t0 > args.budget_min * 60:
            print(f"time budget reached; {len(days) - k + 1} days left for later runs", flush=True)
            break
        try:
            got = fetch_day(s, day)
        except (requests.RequestException, ValueError) as e:
            print(f"  {day}: {type(e).__name__}: {e}", flush=True)
            got = None
            time.sleep(3)
        if got is None:
            failed.append(day)
        else:
            per_day[pd.Timestamp(day)] = got
        if k % 50 == 0:
            print(f"  {k}/{len(days)} ({day})", flush=True)
        if k % args.save_every == 0:
            flush()
    flush()
    if failed:
        print(f"{len(failed)} days without a complete answer: {[str(d) for d in failed[:20]]}", flush=True)
    if daily.empty:
        print("No data returned.", flush=True)
        sys.exit(1)
    C.print_monthly(daily)


if __name__ == "__main__":
    main()
