"""
Bolivia natural gas demand by sector (domestic market), monthly from
2021, from the national statistics institute (INE):
  "Bolivia - Volumen Comercializado de Gas Natural segun Red de
   Distribucion segun Ano y Mes" (sheet H05), linked from
  https://www.ine.gob.bo/index.php/estadisticas-economicas/hidrocarburos-mineria/hidrocarburo-cuadros-estadisticos/

Sectors as published: GNV (vehicle gas), Comercial, Domestico
(residential), Industrial, Generadoras electricas (power), Total. Units
are millions of cubic metres per month (the sheet header states the
unit; it is printed on every run). Exports to Brazil and Argentina are
NOT included - this is the domestic market only.

The sheet lists a year row (annual total) followed by twelve month rows.
INE publishes with a lag; the latest year is marked preliminary (p).
Writes a 'Demand by sector' sheet (million m3/day) plus a native Excel
chart. Not reachable from the editing sandbox; runs in GitHub Actions.
"""
import argparse
import io
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root
import xlsx_charts
import xlsx_notes

PAGE = ("https://www.ine.gob.bo/index.php/estadisticas-economicas/hidrocarburos-mineria/"
        "hidrocarburo-cuadros-estadisticos/")
FALLBACK_URL = "https://nube.ine.gob.bo/index.php/s/Jsm9uMT7L9ALmWr/download"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (10, 120)
DATA_START = 2021
MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
         "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
RENAME = {"gas natural vehicular": "Vehicle_CNG", "comercial": "Commercial", "domestico": "Residential",
          "industrial": "Industrial", "generadoras electricas": "Power", "total": "Total"}


def out(*a):
    print(*a, flush=True)


def strip_accents(s):
    return (str(s).lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o")
            .replace("ú", "u").strip())


def source_url():
    """The download link for the 'Red de Distribucion' table, from INE's page (falls back to the known link)."""
    try:
        r = requests.get(PAGE, headers=H, timeout=T)
        for href, txt in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', r.text, re.S | re.I):
            if re.search(r"Red de Distribuci", re.sub(r"<[^>]+>", "", txt), re.I):
                return href.replace("&amp;", "&")
    except requests.RequestException as e:
        out(f"INE page unavailable ({e}); using known link")
    return FALLBACK_URL


def parse(content):
    raw = pd.read_excel(io.BytesIO(content), sheet_name=0, header=None)
    title = " ".join(str(v) for v in raw.iloc[:4, 0] if str(v) != "nan")
    out(f"sheet header: {title}")
    hdr_row = next(i for i in range(15) if strip_accents(raw.iloc[i, 0]) == "periodo")
    cols = []
    for v in raw.iloc[hdr_row, 1:]:
        name = strip_accents(v)
        cols.append(next((std for key, std in RENAME.items() if name.startswith(key)), str(v)))
    rows, year = [], None
    for _, r in raw.iloc[hdr_row + 1:].iterrows():
        label = strip_accents(r.iloc[0]).replace("(p)", "").strip()
        if re.fullmatch(r"(19|20)\d\d", label.split(" ")[0] if label else ""):
            year = int(label.split(" ")[0])
            continue
        if label in MESES and year:
            vals = pd.to_numeric(r.iloc[1:1 + len(cols)], errors="coerce").values
            rows.append({"Month": pd.Timestamp(year, MESES[label], 1), **dict(zip(cols, vals))})
    return pd.DataFrame(rows).set_index("Month").sort_index(), title


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/bolivia_gas_demand_by_sector.xlsx")
    args = ap.parse_args()

    url = source_url()
    out(f"downloading {url}")
    r = requests.get(url, headers=H, timeout=T)
    r.raise_for_status()
    monthly, title = parse(r.content)
    monthly = monthly[monthly.index.year >= DATA_START].dropna(how="all")
    sectors = [c for c in ["Power", "Industrial", "Residential", "Commercial", "Vehicle_CNG"] if c in monthly]
    days = monthly.index.days_in_month
    perday = monthly.div(days, axis=0).round(3)
    perday.columns = [f"{c}_mcm_per_day" for c in perday.columns]
    table = pd.concat([monthly.round(2).add_suffix("_mcm_month"), perday], axis=1)
    check = (monthly[sectors].sum(axis=1) - monthly["Total"]).abs().max() if "Total" in monthly else None
    out(f"{len(monthly)} months {monthly.index.min():%Y-%m}..{monthly.index.max():%Y-%m}; "
        f"max |sum of sectors - total| = {check}")
    out(perday.tail(6).to_string())

    notes = [
        "UNITS",
        "*_mcm_month: millions of cubic metres per month, as published by INE. *_mcm_per_day: the same divided by "
        "days in the month.",
        f"Source sheet header: {title}",
        "",
        "SECTORS",
        "Power (generadoras electricas), Industrial, Residential (domestico), Commercial, Vehicle_CNG (GNV). "
        "Domestic market only - exports to Brazil and Argentina are not included.",
        "",
        "COVERAGE",
        f"Monthly from {DATA_START}. INE publishes with a lag; the latest year is preliminary (p) and may be revised. "
        "Each run re-reads the whole table, which is small.",
        "",
        "SOURCE",
        f"Instituto Nacional de Estadistica (INE), 'Volumen Comercializado de Gas Natural segun Red de Distribucion': "
        f"{PAGE}",
    ]
    table.index = table.index.strftime("%Y-%m")
    table.index.name = "Month"
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Demand by sector": table}, notes, {"UNITS", "SECTORS", "COVERAGE", "SOURCE"})
    chart_df = monthly[sectors].div(days, axis=0)
    chart_df.columns = [c.replace("_", " ") for c in sectors]
    xlsx_charts.add_chart_sheet(args.out, chart_df, "Bolivia gas demand by sector", "million m3/day",
                                kind="stacked_bar")
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
