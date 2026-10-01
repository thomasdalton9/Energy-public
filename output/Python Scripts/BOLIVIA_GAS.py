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
import datetime
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
MES3 = {k[:3]: v for k, v in MESES.items()}
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


def _cell(v):
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else re.sub(r"\s+", " ", strip_accents(v))


def _sector(name):
    if name.startswith("gnv") or "vehicular" in name:
        return "Vehicle_CNG"
    return next((std for key, std in RENAME.items() if name.startswith(key)), None)


def _year_month(v):
    """(year, month) from a label cell: a year ('2023', '2023(p)'), a month name/abbreviation, or a real date."""
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        return v.year, v.month
    if isinstance(v, (int, float)) and not pd.isna(v):
        return (int(v), None) if float(v).is_integer() and 1990 <= v <= 2100 else (None, None)
    s = _cell(v).replace("(p)", "").replace("p/", "").replace("(*)", "").strip(" .*")
    m = re.fullmatch(r"((?:19|20)\d\d)(?:\.0)?\s*p?", s)
    if m:
        return int(m.group(1)), None
    return None, MES3.get(s[:3]) if len(s) >= 3 and s.isalpha() else None


def dump(raw, name, n=30):
    out(f"--- sheet {name!r} {raw.shape}: first {n} rows ---")
    for i in range(min(n, len(raw))):
        out(i, raw.iloc[i].tolist()[:10])


def parse(content):
    """Find the sheet and row holding the sector headers, then read the year / month rows below it.
    (The first published layout broke a fixed 'Periodo in column A' assumption, so this searches.)"""
    sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
    out(f"sheets: {list(sheets)}")
    for name in sorted(sheets, key=lambda s: 0 if s.strip().upper() == "H05" else 1):
        raw = sheets[name]
        for i in range(min(40, len(raw))):
            sec = {}
            for j, v in enumerate(raw.iloc[i]):
                std = _sector(_cell(v))
                if std and std not in sec.values():
                    sec[j] = std
            if len(set(sec.values()) - {"Total"}) < 3:
                continue
            title = " | ".join(re.sub(r"\s+", " ", v).strip() for r in range(i) for v in raw.iloc[r]
                               if isinstance(v, str) and v.strip())
            dump(raw, name, i + 4)
            out(f"sheet {name!r}, header row {i}: {sec}")
            out(f"sheet header: {title}")
            first = min(sec)
            rows, year = [], None
            for k in range(i + 1, len(raw)):
                r = raw.iloc[k]
                month = None
                for j in range(first):
                    y, m = _year_month(r.iloc[j])
                    year, month = y or year, m or month
                if month and year:
                    rows.append({"Month": pd.Timestamp(year, month, 1),
                                 **{std: pd.to_numeric(r.iloc[j], errors="coerce") for j, std in sec.items()}})
            if rows:
                return pd.DataFrame(rows).groupby("Month").last().sort_index(), title
    for name, raw in sheets.items():
        dump(raw, name)
    raise SystemExit("could not find the sector header row (Comercial / Industrial / ...) in any sheet")


def unit_scale(title):
    """Factor to million m3 from the unit stated in the sheet header."""
    t = strip_accents(title)
    if re.search(r"miles de metros|mm?3 ?\(miles\)|miles de m", t):
        return 1e-3, "thousand m3"
    if re.search(r"millones de pies|mmpc|mmcf", t):
        return 0.0283168, "million cubic feet"
    if re.search(r"millones de metros|mmm3|mmmc|millones de m", t):
        return 1.0, "million m3"
    return 1.0, "not stated in header - assumed million m3"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/bolivia_gas_demand_by_sector.xlsx")
    args = ap.parse_args()

    url = source_url()
    out(f"downloading {url}")
    r = requests.get(url, headers=H, timeout=T)
    r.raise_for_status()
    monthly, title = parse(r.content)
    scale, unit = unit_scale(title)
    out(f"unit from sheet header: {unit} (x{scale} -> million m3)")
    monthly = monthly * scale
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
        f"Unit read from that header: {unit}" + ("" if scale == 1 else f" - converted to million m3 (x{scale})"),
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
