"""
Uruguay generation by source from ADME (Administracion del Mercado
Electrico, the market administrator) - free, public, no key.

Source: ADME's "Generacion por fuente" page (pronos.adme.com.uy/gpf.php,
SCADA data). Asking it for a date range (fecha_ini / fecha_fin,
dd/mm/yyyy) makes it build an hourly .ods file, linked from the page as
/cache/gpf_<a>_<b>_horario.ods. Its 'GPF' sheet has MW per hour for the
four hydro plants (Salto Grande - Uruguay's half of the shared dam -
Bonete, Baygorria, Palmar), wind, solar, thermal, biomass, imports and
demand; the 'Intercambios.' sheet has exports. The .ods is read with
the standard library (zip + XML), so no extra package is needed.
Confirmed against live responses, Sep-2026 (URUGUAY_ADME_DISCOVERY.py).

Cache: the hourly rows are kept in uy_generation_hourly_cache.csv; later
runs only fetch the last few days.

Outputs - uruguay_generation.xlsx:
  'Daily'    mean MW by source per day (hydro, wind, solar, thermal,
             biomass), plus demand, imports, exports
  'Monthly'  GWh by source per month, plus mean MW
  'Hydro by plant'  daily mean MW per hydro plant
  'Reservoir levels'  daily mean lake level (m above sea level) behind
             the Rio Negro dams - Bonete (Rincon del Bonete, Uruguay's main
             storage lake), Baygorria, Palmar - from ADME's per-plant SCADA
             series (seriescentralhidro.cgi, 'hToma'); also turbined flow
  'Salto Grande flows'  daily m3/s turbined, spilled and inflow (Aportes)
             at the binational Salto Grande dam, since 1996 - ADME's
             pronos.adme.com.uy/scripts/saltogrande_xls.php, one download

Usage: python3 URUGUAY_ADME.py [--full]
"""

print("STARTING", flush=True)

import argparse
import datetime as dt
import io
import os
import re
import sys
import time
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

GPF_URL = "https://pronos.adme.com.uy/gpf.php"
SALTO_GRANDE_URL = "https://pronos.adme.com.uy/scripts/saltogrande_xls.php"
PLANT_SERIES_URL = "https://pronos.adme.com.uy/cgi-bin/seriescentralhidro.cgi"
RIO_NEGRO = {"bon": "Bonete", "bay": "Baygorria", "pal": "Palmar"}
LEVEL_START = dt.date(2020, 1, 1)
LEVEL_CACHE_CSV = "uy_reservoir_hourly_cache.csv"
BASE_URL = "https://pronos.adme.com.uy"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36"}
START_DATE = dt.date(2019, 1, 1)  # gpf.php's own history starts in 2019 (older years: gpf_historico.php)
CHUNK_DAYS = 31
REFRESH_DAYS = 3
CACHE_CSV = "uy_generation_hourly_cache.csv"
OUT_FILE = "uruguay_generation.xlsx"

HYDRO_PLANTS = ["Salto Grande", "Bonete", "Baygorria", "Palmar"]
SOURCES = {"hydro": HYDRO_PLANTS, "wind": ["Eólica"], "solar": ["Solar"], "thermal": ["Térmica"], "biomass": ["Biomasa"]}
IMPORTS = ["Imp.Arg", "Imp.Br.Riv", "Imp.Br.Mel"]

# ------------------------------------------------------------------
# .ods reading (zip + XML) - no odfpy needed
# ------------------------------------------------------------------
_T = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
_O = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"


def read_ods(content):
    """{sheet name: list of rows (lists of str/None)}."""
    root = ET.fromstring(zipfile.ZipFile(io.BytesIO(content)).read("content.xml"))
    sheets = {}
    for table in root.iter(_T + "table"):
        rows = []
        for r in table.iter(_T + "table-row"):
            row = []
            for c in r:
                if c.tag not in (_T + "table-cell", _T + "covered-table-cell"):
                    continue
                rep = min(int(c.get(_T + "number-columns-repeated", "1")), 500)
                v = c.get(_O + "value") or c.get(_O + "date-value")
                if v is None:
                    v = "".join(c.itertext()) or None
                row += [v] * rep
            while row and row[-1] is None:
                row.pop()
            if row:
                rows.append(row)
        sheets[table.get(_T + "name")] = rows
    return sheets


def sheet_frame(rows):
    """Rows under the 'Fecha' header -> DataFrame indexed by timestamp."""
    head = next(i for i, r in enumerate(rows) if r and str(r[0]).strip() == "Fecha")
    cols = [str(c).strip() for c in rows[head]]
    body = [r + [None] * (len(cols) - len(r)) for r in rows[head + 1:] if r and r[0]]
    df = pd.DataFrame([r[:len(cols)] for r in body], columns=cols)
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    df = df.dropna(subset=["Fecha"]).set_index("Fecha")
    return df.apply(pd.to_numeric, errors="coerce")


# ------------------------------------------------------------------
# fetch
# ------------------------------------------------------------------

def fetch_chunk(session, start, end):
    """Hourly MW for start..end (inclusive): the GPF summary plus total exports."""
    page = session.get(GPF_URL, params={"fecha_ini": start.strftime("%d/%m/%Y"), "fecha_fin": end.strftime("%d/%m/%Y"),
                                        "send": "MOSTRAR"}, timeout=180)
    page.raise_for_status()
    link = re.search(r"""href=['"]([^'"]*_horario\.ods)['"]""", page.text)
    if not link:
        raise RuntimeError("no hourly .ods link on the page (ADME may not have that range)")
    ods = session.get(BASE_URL + link.group(1) if link.group(1).startswith("/") else link.group(1), timeout=180)
    ods.raise_for_status()
    sheets = read_ods(ods.content)
    gpf = sheet_frame(sheets["GPF"])
    exch_name = next((n for n in sheets if n.lower().startswith("intercambio")), None)
    if exch_name:
        exch = sheet_frame(sheets[exch_name])
        gpf["Exports"] = exch[[c for c in exch.columns if c.lower().startswith("exp")]].sum(axis=1)
    return gpf


def fetch_range(start, end):
    session = requests.Session()
    session.headers.update(HEADERS)
    frames = []
    cur = start
    while cur <= end:
        chunk_end = min(cur + dt.timedelta(days=CHUNK_DAYS - 1), end)
        for attempt in range(3):
            try:
                df = fetch_chunk(session, cur, chunk_end)
                frames.append(df)
                print(f"  {cur}..{chunk_end}: {len(df)} hours", flush=True)
                break
            except (requests.RequestException, RuntimeError, zipfile.BadZipFile, KeyError, StopIteration) as e:
                if attempt == 2:
                    print(f"  {cur}..{chunk_end}: FAILED ({type(e).__name__}: {e})", flush=True)
                else:
                    time.sleep(5 * (attempt + 1))
        cur = chunk_end + dt.timedelta(days=1)
        time.sleep(0.5)
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames)
    return out[~out.index.duplicated(keep="last")].sort_index()


def excel_serial(d):
    return (d - dt.date(1899, 12, 30)).days


def plant_series(session, plant, start, end):
    """Hourly lake level (hToma, m) and turbined flow (m3/s) for one Rio
    Negro plant. ADME times out on long ranges, so callers ask a month
    at a time."""
    r = session.get(PLANT_SERIES_URL, params={"idCentral": plant, "ts": "ods", "dtIni": excel_serial(start),
                                              "dtFin": excel_serial(end + dt.timedelta(days=1))}, timeout=180)
    r.raise_for_status()
    rows = next(iter(read_ods(r.content).values()))
    head = next(i for i, row in enumerate(rows) if len(row) > 2 and str(row[1]).startswith("Pot_MW"))
    cols = {name: j for j, name in enumerate(rows[head]) if name}
    out = []
    for row in rows[head + 1:]:
        if not row or not row[0]:
            continue
        out.append({"time": row[0], "level_m": row[cols["hToma_m_"]], "turbined_m3s": row[cols.get("QTurbinado_m3_s_", 1)]})
    df = pd.DataFrame(out)
    if df.empty:
        return df
    df["time"] = pd.to_datetime(df["time"], errors="coerce")
    df = df.dropna(subset=["time"]).set_index("time").apply(pd.to_numeric, errors="coerce")
    df.columns = [f"{RIO_NEGRO[plant]} {c}" for c in df.columns]
    return df


def fetch_levels(start, end):
    session = requests.Session()
    session.headers.update(HEADERS)
    frames = []
    for plant in RIO_NEGRO:
        parts = []
        cur = start
        while cur <= end:
            chunk_end = min(cur + dt.timedelta(days=CHUNK_DAYS - 1), end)
            for attempt in range(3):
                try:
                    df = plant_series(session, plant, cur, chunk_end)
                    parts.append(df)
                    break
                except (requests.RequestException, zipfile.BadZipFile, StopIteration, KeyError) as e:
                    if attempt == 2:
                        print(f"  {RIO_NEGRO[plant]} {cur}..{chunk_end}: FAILED ({type(e).__name__})", flush=True)
                    else:
                        time.sleep(5 * (attempt + 1))
            cur = chunk_end + dt.timedelta(days=1)
            time.sleep(0.5)
        if parts:
            df = pd.concat(parts)
            frames.append(df[~df.index.duplicated(keep="last")])
            print(f"  {RIO_NEGRO[plant]}: {len(df):,} hourly readings", flush=True)
    return pd.concat(frames, axis=1).sort_index() if frames else pd.DataFrame()


def reservoir_levels(full):
    """Daily mean level and turbined flow per Rio Negro lake, cached."""
    cached = None
    if not full and os.path.exists(LEVEL_CACHE_CSV):
        cached = pd.read_csv(LEVEL_CACHE_CSV, index_col=0, parse_dates=True)
    start = LEVEL_START if cached is None or cached.empty else \
        max(LEVEL_START, (cached.index.max() - pd.Timedelta(days=REFRESH_DAYS)).date())
    print(f"Rio Negro lake levels {start}..{dt.date.today()}...", flush=True)
    fresh = fetch_levels(start, dt.date.today())
    hourly = fresh if cached is None or cached.empty else (fresh.combine_first(cached) if not fresh.empty else cached)
    if hourly is None or hourly.empty:
        return pd.DataFrame()
    hourly = hourly.sort_index()
    hourly.index.name = "time"
    hourly.to_csv(LEVEL_CACHE_CSV)
    daily = hourly.groupby(hourly.index.date).mean().round(3)
    daily.index = pd.to_datetime(daily.index)
    daily.index.name = "date"
    return daily


def fetch_salto_grande():
    """Daily flows at Salto Grande (m3/s): Turbinado, Vertido, Aportes.
    The whole history comes in one .xlsx (served as .xls)."""
    r = requests.get(SALTO_GRANDE_URL, headers=HEADERS, timeout=180)
    r.raise_for_status()
    df = pd.read_excel(io.BytesIO(r.content), header=1, engine="openpyxl")
    df = df.rename(columns={df.columns[0]: "date"})
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"]).set_index("date").sort_index()
    df = df.apply(pd.to_numeric, errors="coerce").round(1)
    df.columns = [f"{c} (m3/s)" for c in df.columns]
    return df


# ------------------------------------------------------------------
# shape
# ------------------------------------------------------------------

def by_source(hourly):
    out = pd.DataFrame(index=hourly.index)
    for name, cols in SOURCES.items():
        have = [c for c in cols if c in hourly.columns]
        out[name] = hourly[have].sum(axis=1, min_count=1) if have else float("nan")
    out["demand"] = hourly.get("Demanda")
    out["imports"] = hourly[[c for c in IMPORTS if c in hourly.columns]].sum(axis=1, min_count=1)
    out["exports"] = hourly.get("Exports")
    return out


def complete_days(hourly):
    """Readings are on the hour, midnight to midnight: a day's own 24 are
    00:00..23:00. Drop a trailing day that hasn't got them all yet."""
    counts = hourly.groupby(hourly.index.date).size()
    last_full = max((d for d, n in counts.items() if n >= 24), default=None)
    return hourly[hourly.index.date <= last_full] if last_full else hourly


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--full", action="store_true", help="ignore the cache and fetch everything from START_DATE")
    args = parser.parse_args()

    cached = None
    if not args.full and os.path.exists(CACHE_CSV):
        cached = pd.read_csv(CACHE_CSV, index_col=0, parse_dates=True)
    start = START_DATE if cached is None or cached.empty else \
        max(START_DATE, (cached.index.max() - pd.Timedelta(days=REFRESH_DAYS)).date())
    end = dt.date.today()
    print(f"ADME generation by source {start}..{end}...", flush=True)
    fresh = fetch_range(start, end)
    hourly = fresh if cached is None or cached.empty else \
        (fresh.combine_first(cached) if not fresh.empty else cached)
    if hourly is None or hourly.empty:
        print("No data returned.", flush=True)
        sys.exit(1)
    hourly = hourly.sort_index()
    hourly.index.name = "Fecha"
    hourly.to_csv(CACHE_CSV)

    hourly = complete_days(hourly)
    src = by_source(hourly)
    daily = src.groupby(src.index.date).mean().round(1)
    daily.index = pd.to_datetime(daily.index)
    daily.index.name = "date"
    hydro_plants = hourly[[c for c in HYDRO_PLANTS if c in hourly.columns]]
    hydro_daily = hydro_plants.groupby(hydro_plants.index.date).mean().round(1)
    hydro_daily.index = pd.to_datetime(hydro_daily.index)
    hydro_daily.index.name = "date"

    month = src.index.to_period("M")
    hours = src.groupby(month).size()
    monthly = pd.DataFrame(index=hours.index)
    mean_mw = src.groupby(month).mean()
    for col in list(SOURCES) + ["demand", "imports", "exports"]:
        monthly[f"{col}_gwh"] = (mean_mw[col] * hours / 1000).round(1)
    for col in SOURCES:
        monthly[f"{col}_avg_mw"] = mean_mw[col].round(1)
    monthly["hours_of_data"] = hours
    monthly.index = monthly.index.to_timestamp()
    monthly.index.name = "month"

    notes = [
        "UNITS",
        "'Daily' and 'Hydro by plant': MW, the day's mean of ADME's hourly readings. "
        "'Monthly': GWh (mean MW x hours of data / 1000); the latest month is partial - see hours_of_data.",
        "",
        "SOURCES",
        "hydro = Salto Grande (Uruguay's half of the binational dam, as ADME reports it) + Bonete + Baygorria + "
        "Palmar. thermal, wind, solar, biomass as ADME groups them. imports = Argentina + Brazil (Rivera, Melo); "
        "exports from ADME's interchange sheet. Generation - exports + imports = demand.",
        "",
        "SOURCE",
        "ADME 'Generacion por fuente' (pronos.adme.com.uy/gpf.php) - SCADA values, hourly .ods files. "
        "History from 2019 (earlier years are on ADME's gpf_historico.php page, not pulled).",
        "'Reservoir levels': daily mean lake level upstream of each Rio Negro dam (hToma, metres above sea "
        "level) and turbined flow (m3/s), from ADME's hourly per-plant SCADA series, since 2020. Rincon del "
        "Bonete is the big storage lake; ADME doesn't publish stored volume, so these are levels, not totals.",
        "'Salto Grande flows': ADME's daily Salto Grande file - Turbinado (through the turbines), Vertido "
        "(spilled), Aportes (inflow), m3/s, since 1996. The dam is shared with Argentina.",
    ]
    sheets = {"Daily": daily, "Monthly": monthly, "Hydro by plant": hydro_daily}
    try:
        levels = reservoir_levels(args.full)
        if not levels.empty:
            sheets["Reservoir levels"] = levels
    except Exception as e:  # keep the generation output even if ADME's series page misbehaves
        print(f"  lake levels FAILED ({type(e).__name__}: {e})", flush=True)
    try:
        sheets["Salto Grande flows"] = fetch_salto_grande()
        print(f"  Salto Grande flows: {len(sheets['Salto Grande flows']):,} days", flush=True)
    except (requests.RequestException, ValueError, KeyError) as e:
        print(f"  Salto Grande flows FAILED ({type(e).__name__}: {e})", flush=True)
        try:  # keep the last good copy
            sheets["Salto Grande flows"] = pd.read_excel(OUT_FILE, sheet_name="Salto Grande flows", index_col=0)
        except (FileNotFoundError, ValueError):
            pass
    xlsx_notes.write_workbook(OUT_FILE, sheets,
                              notes, {"UNITS", "SOURCES", "SOURCE"})
    print(f"\nSaved {OUT_FILE}: {len(daily):,} days ({daily.index.min():%d-%b-%Y} to {daily.index.max():%d-%b-%Y})")
    print(daily.tail(5).to_string())


if __name__ == "__main__":
    main()
