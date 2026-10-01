"""
Dominican Republic daily power generation by fuel from the Organismo
Coordinador del SENI (OC, the national grid operator) - raw grid-operator
pull in the standard layout.

Source: the web service behind OC's public "Reportes" page
(https://www.oc.org.do/Servicios/Reporte, chart "Generacion acumulada"):
  https://apps.oc.org.do/wsOCWebsiteChart/Service.asmx/GetGeneracionAcumuladaJSon
      ?Filtro=Ax&Desde=YYYY-MM-DD&Hasta=YYYY-MM-DD
returns one row per day from `Desde` with the SENI's generation by fuel in
MWh: Solar, Eolica, Biomasa, Carbon, FuelOilNo.2, FuelOilNo.2yNo.6,
FuelOilNo.6, GasNatural, Hidroelectrica. (Filtro=Dx returns the hours of one
day.) History from 2021-01-01. Found via
discovery_archive/south_america/CARIBBEAN_POWER_GAS_PROBE2.py.

Incremental: each run asks only from the first day missing from the workbook
(or 14 days before the last saved day, as OC revises recent days) and keeps
complete days only (before today, Santo Domingo time).

Usage: python3 DOMINICAN_REPUBLIC_OC_POWER_DAILY.py [--out PATH] [--start YYYY-MM-DD]
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

OUT = "output/Data and Chart Outputs/dominican_republic_power_generation_daily.xlsx"
URL = "https://apps.oc.org.do/wsOCWebsiteChart/Service.asmx/GetGeneracionAcumuladaJSon"
PAGE = "https://www.oc.org.do/Servicios/Reporte"
REFRESH_DAYS = 14

MAPPING = {
    "Hidroelectrica": "Hydro",
    "GasNatural": "Gas",
    "Eolica": "Wind",
    "Solar": "Solar",
    "Carbon": "Coal",
    "FuelOilNo.2": "Oil",
    "FuelOilNo.6": "Oil",
    "FuelOilNo.2yNo.6": "Oil",
    "Biomasa": "Bioenergy",
}


def today_do():
    return (dt.datetime.utcnow() - dt.timedelta(hours=4)).date()


def fetch(start, end):
    params = {"Filtro": "Ax", "Desde": f"{start:%Y-%m-%d}", "Hasta": f"{end:%Y-%m-%d}"}
    for attempt in range(4):
        try:
            r = requests.get(URL, params=params, headers={"User-Agent": C.USER_AGENT, "Referer": PAGE},
                             timeout=(15, 180))
            r.raise_for_status()
            rows = r.json()["GetGeneracionAcumulada"]
            break
        except (requests.RequestException, ValueError, KeyError) as e:
            print(f"  attempt {attempt + 1}: {type(e).__name__}: {e}", flush=True)
            time.sleep(10 * (attempt + 1))
    else:
        raise RuntimeError("OC web service did not answer")
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["date"] = pd.to_datetime(df["FECHA"]).dt.normalize()
    num = [c for c in df.columns if c not in ("FECHA", "date", "ACTUALIZADO", "TOTAL")]
    detail = df.groupby("date")[num].sum(min_count=1).astype(float)
    return detail[(detail.index >= pd.Timestamp(start)) & (detail.index <= pd.Timestamp(end))]


def to_standard(detail):
    unknown = [c for c in detail.columns if c not in MAPPING]
    if unknown:
        print(f"  new OC categories counted as Other: {unknown}", flush=True)
    fuels = detail.T.groupby(lambda c: MAPPING.get(c, "Other")).sum(min_count=1).T
    out = C.standardise(fuels)
    return out[detail.notna().any(axis=1).values]


def notes(daily):
    return [
        "UNITS",
        "MWh of generation per day (Santo Domingo calendar day). Total_MWh = sum of the fuel columns.",
        "Detail: OC's own fuel categories, MWh per day.",
        "",
        "COVERAGE",
        *C.coverage_lines(daily),
        "Days are kept only once complete (before today, Santo Domingo time); the last 14 days are re-read each "
        "run because OC revises them.",
        "",
        "SOURCE",
        "Organismo Coordinador del SENI (OC), Dominican Republic grid operator - 'Generacion acumulada' chart on "
        f"{PAGE}, served by {URL}?Filtro=Ax&Desde=YYYY-MM-DD&Hasta=YYYY-MM-DD (daily MWh by fuel, SENI).",
        "",
        "MAPPING",
        *[f"{k} -> {v}" for k, v in MAPPING.items()],
        "Isolated systems outside the SENI and self-generation are not included.",
    ]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start", default=None, help="YYYY-MM-DD (default: first missing day)")
    args = ap.parse_args()

    end = today_do() - dt.timedelta(days=1)
    existing = C.load_sheet(args.out, "Daily")
    old_detail = C.load_sheet(args.out, "Detail")
    if args.start:
        start = dt.date.fromisoformat(args.start)
    else:
        todo = C.days_to_fetch(existing, C.HISTORY_START, end, REFRESH_DAYS)
        if not todo:
            print("Nothing to fetch", flush=True)
            return
        start = todo[0]
    print(f"Fetching OC daily generation {start} .. {end}", flush=True)
    detail = fetch(start, end)
    if detail.empty and existing.empty:
        sys.exit("OC returned no data")
    new = to_standard(detail) if not detail.empty else pd.DataFrame()
    daily = C.merge(new, existing)
    detail = C.merge(detail, old_detail)
    C.print_mapping(MAPPING, "OC category")
    C.write(args.out, daily, notes(daily), detail)
    C.print_monthly(daily)


if __name__ == "__main__":
    main()
