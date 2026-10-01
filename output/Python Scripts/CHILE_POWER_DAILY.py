"""
Chile daily gross power generation by technology (SEN, the national grid),
in the standard layout shared by the South America raw grid-operator pulls
(see power_daily_std.py). No API key needed.

Source: CNE (Comision Nacional de Energia) statistics workbook "Generacion
Bruta" (https://www.cne.cl/wp-content/uploads/YYYY/MM/Generacion_Bruta.xlsx,
linked from https://www.cne.cl/estadisticas/electricidad/). One sheet per
year ("21sen" ... "26sen") with DAILY gross generation (MWh) for every SEN
plant/unit-fuel, compiled by CNE from Coordinador Electrico Nacional (CEN)
data, each column tagged with its technology ("TECNOLOGIA" row). CNE
re-uploads the workbook about once a month (e.g. data to 31-Aug-2026 was
uploaded 10-Sep-2026); the newest upload is found through cne.cl's public
WordPress media search. (CEN's own coordinador.cl downloads sit behind a
Cloudflare browser challenge that rejects GitHub runners, and its SIP API
needs a user key - see discovery_archive/south_america/CHILE_POWER_HYDRO_*.py.)

Incremental: the saved workbook is the archive. Each run asks cne.cl for the
newest upload; when it is the file already processed and no day is missing,
nothing is downloaded. Otherwise the file is downloaded and only the year
sheets holding missing days, or the last REVISION_DAYS saved days, are read.

Usage: python3 CHILE_POWER_DAILY.py [--out PATH] [--start YYYY-MM-DD] [--force]
"""

print("STARTING", flush=True)

import argparse
import datetime as dt
import io
import os
import re
import sys

import openpyxl
import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import power_daily_std as std  # noqa: E402

OUT = "output/Data and Chart Outputs/chile_power_generation_daily.xlsx"
MEDIA_API = "https://www.cne.cl/wp-json/wp/v2/media"
STATS_PAGE = "https://www.cne.cl/estadisticas/electricidad/"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
REVISION_DAYS = 62   # CNE's monthly re-upload can revise the latest month or two

# CNE 'TECNOLOGIA' label -> standard fuel
MAPPING = {
    "Hidráulica Pasada": "Hydro", "Hidráulica Embalse": "Hydro",
    "Gas Natural": "Gas", "GNL": "Gas",
    "Eólica": "Wind",
    "Solar": "Solar", "Concentración Solar": "Solar",
    "Carbón": "Coal", "Carbón + Petcoke": "Coal", "Petcoke": "Coal",
    "Petróleo Diesel": "Oil", "Fuel Oil": "Oil", "Petróleo Diesel + Fuel Oil": "Oil",
    "Biomasa": "Bioenergy",
    "Geotérmica": "Other", "Cogeneración": "Other",
}
SKIP = {"TOTAL", "Máxima (MW)"}

NOTES = [
    "UNITS",
    "MWh per day (gross generation at the plant terminals, as published by CNE from Coordinador Eléctrico "
    "Nacional data). Sistema Eléctrico Nacional (SEN) only - the small Aysén and Magallanes systems are not "
    "included.",
    "",
    "COVERAGE",
    "Daily from 2021-01-01. CNE re-publishes the workbook about once a month, roughly 10 days after month end, "
    "so the latest data is usually 5-6 weeks old. Days CNE has not filled yet (all plants zero or blank) are "
    "left out.",
    "",
    "CATEGORY MAPPING (CNE 'TECNOLOGÍA' -> sheet 'Daily')",
    "Hydro_MWh = Hidráulica Pasada (run-of-river, including mini hydro) + Hidráulica Embalse (reservoir).",
    "Gas_MWh = Gas Natural + GNL. Wind_MWh = Eólica. Solar_MWh = Solar (PV) + Concentración Solar (CSP).",
    "Coal_MWh = Carbón + Carbón + Petcoke + Petcoke. Oil_MWh = Petróleo Diesel + Fuel Oil + Petróleo Diesel + "
    "Fuel Oil (diesel engines/turbines and fuel-oil units).",
    "Bioenergy_MWh = Biomasa. Other_MWh = Geotérmica (Cerro Pabellón) + Cogeneración (sulphuric-acid plant heat "
    "recovery and refinery cogeneration). Nuclear_MWh = 0 (Chile has no nuclear). Batteries are not in the CNE "
    "table.",
    "Total_MWh = sum of the fuel columns; it equals CNE's own 'TOTAL' column.",
    "Sheet 'By technology': the same daily MWh for each CNE technology label before mapping.",
    "",
    "VALIDATION",
    "Against Ember monthly data for Chile (Jan-2021 to Apr-2026, the months both cover): Hydro -0.1%, Gas 0.0%, "
    "Wind -1.2%, Solar -0.4%, Coal +1.4%, Bioenergy +9.0%, Oil vs Ember 'Other Fossil' -41%, Other vs Ember "
    "'Other Renewables' +87%, total -0.2%. Ember counts some units as other fossil that CNE tags as biomass or "
    "cogeneration; the small Oil/Other categories differ by under 0.1 TWh a month.",
    "",
    "SOURCE",
    "CNE, Estadísticas > Electricidad > 'Generación Bruta' workbook (daily generation by plant, SEN sheets "
    "18sen..26sen) - https://www.cne.cl/estadisticas/electricidad/ ; file e.g. "
    "https://www.cne.cl/wp-content/uploads/2026/09/Generacion_Bruta.xlsx. Data compiled by CNE from Coordinador "
    "Eléctrico Nacional. Updated by GitHub Actions (chile_power_generation_daily.yml), which checks daily for a "
    "new upload and only reads the year sheets with missing or recent days.",
]


def latest_upload():
    """(url, modified) of the newest SEN 'Generacion_Bruta*.xlsx' in cne.cl's media library."""
    r = requests.get(MEDIA_API, params={"search": "Generacion_Bruta", "per_page": 50,
                                        "_fields": "date,modified,source_url"}, headers=HEADERS, timeout=(10, 60))
    r.raise_for_status()
    items = [i for i in r.json() if re.search(r"/Generacion_Bruta(-\d+)?\.xlsx$", i.get("source_url", ""), re.I)]
    if items:
        best = max(items, key=lambda i: (i.get("modified") or i.get("date") or "", i.get("date") or ""))
        return best["source_url"], best.get("modified") or best.get("date")
    # fallback: the statistics page link
    r = requests.get(STATS_PAGE, headers=HEADERS, timeout=(10, 60))
    m = re.findall(r"https://www\.cne\.cl/wp-content/uploads/\d{4}/\d{2}/Generacion_Bruta(?:-\d+)?\.xlsx", r.text)
    if not m:
        raise RuntimeError("No Generacion_Bruta.xlsx found on cne.cl")
    return sorted(m)[-1], ""


def download(url):
    for attempt in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=(10, 300))
            r.raise_for_status()
            if r.content[:2] != b"PK":
                raise RuntimeError(f"not an xlsx ({r.headers.get('content-type')})")
            print(f"Downloaded {url}: {len(r.content) / 1e6:.1f} MB", flush=True)
            return r.content
        except Exception as e:  # noqa: BLE001
            print(f"  download attempt {attempt + 1} failed: {e}", flush=True)
    raise RuntimeError(f"Could not download {url}")


def parse_year(wb, year):
    """Daily MWh by CNE technology label for one '{yy}sen' sheet; days with no data are dropped."""
    name = f"{year % 100:02d}sen"
    if name not in wb.sheetnames:
        print(f"  sheet {name} not in workbook", flush=True)
        return pd.DataFrame()
    rows = wb[name].iter_rows(values_only=True)
    tech = None
    recs = []
    for i, r in enumerate(rows):
        if tech is None:
            if len(r) > 1 and isinstance(r[1], str) and r[1].strip().upper().startswith("TECNOLOG"):
                tech = [str(x).strip() if x is not None else None for x in r]
            continue
        if len(r) < 2 or not isinstance(r[1], dt.datetime):
            if recs and r[1] is not None and str(r[1]).strip().lower().startswith("total"):
                break
            continue
        rec = {}
        for j, t in enumerate(tech):
            if j < 2 or t is None or t in SKIP:
                continue
            v = r[j] if j < len(r) else None
            if isinstance(v, (int, float)):
                rec[t] = rec.get(t, 0.0) + float(v)
        if sum(rec.values()) > 0:
            recs.append((pd.Timestamp(r[1]).normalize(), rec))
    if tech is None:
        raise RuntimeError(f"{name}: no TECNOLOGÍA header row - CNE layout changed")
    unknown = sorted({t for _, rec in recs for t in rec} - set(MAPPING))
    if unknown:
        print(f"  WARNING {name}: technology labels not mapped (counted as Other): {unknown}", flush=True)
    df = pd.DataFrame([rec for _, rec in recs], index=[d for d, _ in recs]).fillna(0.0)
    df.index.name = "date"
    print(f"  {name}: {len(df)} days with data ({df.index.min():%d-%b-%Y} to {df.index.max():%d-%b-%Y})"
          if len(df) else f"  {name}: no data", flush=True)
    return df


def to_fuels(by_tech):
    fuels = pd.DataFrame(index=by_tech.index)
    for f in std.FUELS:
        cols = [c for c in by_tech.columns if MAPPING.get(c, "Other") == f]
        fuels[f] = by_tech[cols].sum(axis=1) if cols else 0.0
    return std.standardise(fuels)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--start", type=dt.date.fromisoformat, default=std.HISTORY_START)
    ap.add_argument("--force", action="store_true", help="download and re-read even if the upload is unchanged")
    args = ap.parse_args()

    daily = std.load_sheet(args.out, "Daily")
    by_tech = std.load_sheet(args.out, "By technology")
    try:
        meta = pd.read_excel(args.out, sheet_name="Source file", index_col=0)["value"].to_dict()
    except Exception:  # noqa: BLE001 - first run / older layout
        meta = {}

    url, modified = latest_upload()
    print(f"Newest CNE upload: {url} (modified {modified}); last processed: {meta.get('url')} "
          f"({meta.get('modified')})", flush=True)
    have = set(daily.index[daily["Total_MWh"].notna()].date) if not daily.empty else set()
    last = max(have) if have else None
    gaps = [args.start + dt.timedelta(days=i) for i in range(((last or args.start) - args.start).days + 1)]
    gaps = [d for d in gaps if d not in have]
    if not args.force and have and not gaps and meta.get("url") == url and str(meta.get("modified")) == str(modified):
        print(f"Unchanged upload and no gaps ({len(have):,} days to {last}) - nothing to do.", flush=True)
        return

    years = {d.year for d in gaps}
    if last:
        years |= {(last - dt.timedelta(days=REVISION_DAYS)).year, last.year}
        years |= set(range(last.year, dt.date.today().year + 1))   # anything newer than the archive
    else:
        years = set(range(args.start.year, dt.date.today().year + 1))
    years = sorted(y for y in years if y >= args.start.year)
    print(f"{len(have):,} days saved (to {last}); {len(gaps)} gap days; reading years {years}", flush=True)

    wb = openpyxl.load_workbook(io.BytesIO(download(url)), read_only=True, data_only=True)
    frames = [parse_year(wb, y) for y in years]
    new_tech = pd.concat([f for f in frames if not f.empty]) if any(not f.empty for f in frames) else pd.DataFrame()
    if new_tech.empty:
        print("No data parsed.", flush=True)
        sys.exit(1)
    new_tech = new_tech[new_tech.index >= pd.Timestamp(args.start)]
    # only gap days, plus everything from the revision window on, replace saved rows
    cutoff = pd.Timestamp(last - dt.timedelta(days=REVISION_DAYS)) if last else pd.Timestamp(args.start)
    keep = new_tech.index.isin(pd.to_datetime(sorted(gaps))) | (new_tech.index > cutoff)
    new_tech = new_tech[keep].round(1)
    by_tech = std.merge(new_tech, by_tech).fillna(0.0)
    by_tech = by_tech[sorted(by_tech.columns, key=lambda c: (list(MAPPING).index(c) if c in MAPPING else 99, c))]
    daily = std.merge(to_fuels(new_tech), daily)
    daily = daily[[c for c in [f"{f}_MWh" for f in std.FUELS] + ["Total_MWh"] if c in daily]]
    src = pd.DataFrame({"value": [url, str(modified), dt.datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")]},
                       index=pd.Index(["url", "modified", "processed"], name="key"))
    print(f"{len(new_tech)} days updated", flush=True)
    std.write(args.out, daily, NOTES, {"By technology": by_tech, "Source file": src})


if __name__ == "__main__":
    main()
