"""
Peru daily power generation by type from COES (Comite de Operacion
Economica del Sistema Interconectado Nacional) - free, public, no key.

Source: COES Portal de Informacion, 'Generacion' page
(www.coes.org.pe/Portal/portalinformacion/generacion). The page POSTs
fechaInicial / fechaFinal (dd/mm/yyyy) and indicador=0 (SCADA) to the same
URL and gets JSON back:
  GraficoTipoCombustible  half-hourly MW by fuel (AGUA/HIDRICO, GAS,
                          DIESEL, RESIDUAL, CARBON, NAFTA & GAS REFINERIA,
                          BIOGAS, BAGAZO, EOLICA, SOLAR) - for the FIRST
                          day of the range only, whatever the range
  GraficoPorEmpresa       MWh per company by technology (HIDROELECTRICA,
                          TERMOELECTRICA, SOLAR, EOLICA, _NO DEFINIDO)
So this asks for one day at a time (works back to 2021-01-01 and beyond).
Found via discovery_archive/south_america/PERU_COES_ECUADOR_CENACE_DISCOVERY.py.

Daily MWh = sum of the 48 half-hourly MW values x 0.5 h (COES' own
per-company MWh totals are the same sum, checked).

COES labels two fuel series the wrong way round: the one called 'BAGAZO'
has the solar daytime shape and equals the SOLAR technology total, the one
called 'SOLAR' is the small flat bagasse output. To stay right if COES
fixes the labels, Solar comes from the technology split and bagasse is
taken as (BAGAZO + SOLAR fuel series) - Solar technology.

Incremental: keeps every day in the workbook's 'COES raw' sheet, fetches
only days missing since 2021-01-01, and re-fetches the last REFRESH_DAYS
days (SCADA values are preliminary and the latest day may be partial).

Usage: python3 PERU_COES_GENERATION.py [--out PATH] [--start YYYY-MM-DD]
"""

print("STARTING", flush=True)

import argparse
import os
import sys
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

URL = "https://www.coes.org.pe/Portal/portalinformacion/generacion"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36", "X-Requested-With": "XMLHttpRequest"}
DEFAULT_OUT = os.path.join("output", "Data and Chart Outputs", "peru_power_generation_daily.xlsx")
START = date(2021, 1, 1)
REFRESH_DAYS = 5
WORKERS = 4
CHECKPOINT = 150  # days fetched between workbook checkpoints
LIMA = timezone(timedelta(hours=-5))

STANDARD = ["Hydro_MWh", "Gas_MWh", "Wind_MWh", "Solar_MWh", "Coal_MWh", "Oil_MWh", "Bioenergy_MWh", "Other_MWh"]


def key(name):
    """'HÍDRICO' -> 'HIDRICO' (COES' accents and spacing vary)."""
    s = unicodedata.normalize("NFKD", str(name)).encode("ascii", "ignore").decode().upper().strip()
    return " ".join(s.split())


def fetch_day(day, session):
    d = day.strftime("%d/%m/%Y")
    for attempt in range(4):
        try:
            r = session.post(URL, data={"fechaInicial": d, "fechaFinal": d, "indicador": 0}, timeout=120)
            r.raise_for_status()
            j = r.json()
            break
        except (requests.RequestException, ValueError) as e:
            if attempt == 3:
                print(f"  {day}: FAILED {type(e).__name__}: {str(e)[:150]}", flush=True)
                return None
            time.sleep(5 * (attempt + 1))
    row, points = {}, 0
    for s in (j.get("GraficoTipoCombustible") or {}).get("Series") or []:
        vals = [p for p in s.get("Data") or [] if str(p.get("Nombre", "")).startswith(day.strftime("%Y/%m/%d"))
                or str(p.get("Nombre", "")) == (day + timedelta(days=1)).strftime("%Y/%m/%d 00:00:00")]
        points = max(points, sum(p.get("Valor") is not None for p in vals))
        row["fuel: " + key(s.get("Name"))] = sum(float(p.get("Valor") or 0) for p in vals) / 2
    for s in (j.get("GraficoPorEmpresa") or {}).get("Series") or []:
        row["tech: " + key(s.get("Name"))] = sum(float(v or 0) for v in s.get("Data") or [])
    if not row:
        return None
    row["halfhours"] = points
    return row


def to_standard(raw):
    f = lambda c: raw[c].fillna(0) if c in raw else pd.Series(0.0, index=raw.index)  # noqa: E731
    fuel_total = raw[[c for c in raw.columns if c.startswith("fuel: ")]].fillna(0).sum(axis=1)
    out = pd.DataFrame(index=raw.index)
    out["Hydro_MWh"] = f("tech: HIDROELECTRICA")
    out["Gas_MWh"] = f("fuel: GAS")
    out["Wind_MWh"] = f("tech: EOLICA")
    out["Solar_MWh"] = f("tech: SOLAR")
    out["Coal_MWh"] = f("fuel: CARBON")
    out["Oil_MWh"] = f("fuel: DIESEL") + f("fuel: RESIDUAL") + f("fuel: NAFTA & GAS REFINERIA")
    bagasse = (f("fuel: BAGAZO") + f("fuel: SOLAR") - f("tech: SOLAR")).clip(lower=0)
    out["Bioenergy_MWh"] = f("fuel: BIOGAS") + bagasse
    known = out[STANDARD[:-1]].sum(axis=1)
    out["Other_MWh"] = (fuel_total - known).clip(lower=0)  # '_NO DEFINIDO' and any new fuel label
    out["Total_MWh"] = fuel_total
    out.index.name = "date"
    return out.round(1)


def load(path):
    try:
        raw = pd.read_excel(path, sheet_name="COES raw", index_col=0)
    except (FileNotFoundError, ValueError) as e:
        print(f"No archive yet ({type(e).__name__}) - full backfill from {START}", flush=True)
        return pd.DataFrame()
    raw.index = pd.to_datetime(raw.index).date
    raw.index.name = "date"
    return raw


def save(path, raw):
    raw = raw.sort_index()
    daily = to_standard(raw)
    first, last = daily.index.min(), daily.index.max()
    notes = [
        "UNITS",
        "'Daily': energy generated per day in MWh by type (sum of COES' 48 half-hourly SCADA MW values x 0.5 h). "
        "Total_MWh = all generation COES reports for the SEIN (national grid); no imports/exports.",
        "'COES raw': the same days as COES reports them - 'fuel: ...' = MWh by fuel from the fuel-type chart, "
        "'tech: ...' = MWh by technology summed over companies; halfhours = half-hour points received.",
        "",
        "CATEGORY MAPPING",
        "Hydro_MWh = technology HIDROELECTRICA (= fuel HIDRICO/AGUA)",
        "Gas_MWh = fuel GAS (Camisea, Aguaytia, Malacas etc. natural gas)",
        "Wind_MWh = technology EOLICA; Solar_MWh = technology SOLAR",
        "Coal_MWh = fuel CARBON",
        "Oil_MWh = fuels DIESEL + RESIDUAL + NAFTA & GAS REFINERIA (charted as 'Other Fossil')",
        "Bioenergy_MWh = fuel BIOGAS + bagasse. COES labels its 'BAGAZO' and 'SOLAR' fuel series the wrong way "
        "round (the 'BAGAZO' series has the solar daytime shape and equals the SOLAR technology total), so bagasse "
        "= ('BAGAZO' + 'SOLAR' fuel series) - Solar technology total.",
        "Other_MWh = anything COES reports outside those fuels (technology '_NO DEFINIDO').",
        "No nuclear in Peru, so no Nuclear_MWh column.",
        "",
        "COVERAGE",
        f"{first} to {last} ({len(daily)} days). History from 2021-01-01 (COES serves earlier days too). "
        f"Days with fewer than 48 half-hour points: {int((raw.get('halfhours', pd.Series(48)) < 48).sum())}.",
        "Values are COES' preliminary SCADA data (indicador=0); the last few days are re-fetched every run.",
        "",
        "SOURCE",
        "COES Portal de Informacion - Generacion: POST https://www.coes.org.pe/Portal/portalinformacion/generacion "
        "with fechaInicial=fechaFinal=dd/mm/yyyy, indicador=0 (keyless). One request per day.",
        "Script: south_america/PERU_COES_GENERATION.py (scheduled by .github/workflows/peru_power_generation.yml).",
    ]
    xlsx_notes.write_workbook(path, {"Daily": daily, "COES raw": raw}, notes,
                              {"UNITS", "CATEGORY MAPPING", "COVERAGE", "SOURCE"})
    return daily


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--start", default=str(START))
    args = ap.parse_args()
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)

    raw = load(args.out)
    yesterday = datetime.now(LIMA).date() - timedelta(days=1)
    start = date.fromisoformat(args.start)
    have = set(raw.index) if not raw.empty else set()
    refresh = {yesterday - timedelta(days=i) for i in range(REFRESH_DAYS)}
    wanted = [start + timedelta(days=i) for i in range((yesterday - start).days + 1)]
    todo = [d for d in wanted if d not in have or d in refresh]
    print(f"Archive: {len(have)} days; fetching {len(todo)} day(s) ({todo[0] if todo else '-'} .. "
          f"{todo[-1] if todo else '-'})", flush=True)

    rows, failed, done = {}, [], 0
    session = requests.Session()
    session.headers.update(HEADERS)
    with ThreadPoolExecutor(WORKERS) as pool:
        futures = {pool.submit(fetch_day, d, session): d for d in todo}
        for fut in as_completed(futures):
            d = futures[fut]
            row = fut.result()
            if row is None:
                failed.append(d)
            else:
                rows[d] = row
            done += 1
            if done % CHECKPOINT == 0:
                print(f"  {done}/{len(todo)} fetched", flush=True)
                new = pd.DataFrame.from_dict(rows, orient="index")
                save(args.out, new.combine_first(raw) if not raw.empty else new)
    if rows:
        new = pd.DataFrame.from_dict(rows, orient="index")
        new.index.name = "date"
        raw = new.combine_first(raw) if not raw.empty else new
    if raw.empty:
        sys.exit("No data fetched and no archive - nothing to write")
    daily = save(args.out, raw)
    tech = raw[[c for c in raw.columns if c.startswith("tech: ")]].fillna(0).sum(axis=1)
    gap = (tech - daily["Total_MWh"]).abs()
    print(f"Saved {args.out}: {len(daily)} days {daily.index.min()} .. {daily.index.max()}; "
          f"failed {len(failed)} {sorted(failed)[:10]}; max |tech - fuel total| {gap.max():,.1f} MWh", flush=True)
    print(daily.tail(7).to_string(), flush=True)
    if failed and len(failed) > max(10, len(todo) // 4):
        sys.exit(f"{len(failed)} of {len(todo)} days failed")


if __name__ == "__main__":
    main()
