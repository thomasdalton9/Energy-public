"""
Mexico natural gas demand by sector (power, industrial, residential, commercial/services, transport CNG,
Pemex own use / other), monthly -> mexico_gas_demand_by_sector.xlsx.

STATUS: written WITHOUT being able to reach any Mexican source from the editing sandbox (sie.energia.gob.mx,
datos.gob.mx, pemex.com, gob.mx/cenagas, ri.pemex.com all returned a proxy 403). The parser is therefore
deliberately defensive and DATA-DRIVEN: it never writes a number it cannot tie to a recognised sector label and
a recognised unit. Run discovery_archive/americas/MEXICO_GAS_SECTOR_DISCOVERY.py (workflow
discovery_archive/workflows/mexico_gas_sector_discovery.yml) once in GitHub Actions, read its log, and pin the
confirmed resource(s) in PINNED below; until then the script searches the open-data catalogue.

Raw sources, in order of preference (all Mexican government / state company):
  1. Pemex (Pemex Transformacion Industrial / Gas y Petroquimica Basica) "Ventas internas de gas natural por
     sector" in the Base de Datos Institucional and on datos.gob.mx (CKAN: https://datos.gob.mx/api/3/action).
     Pemex's published sector split is: Distribuidoras, Industrial, Electrico (CFE), Pemex (own use) - not the
     end-use split of residential / commercial, which sits inside 'Distribuidoras'.
  2. SENER Sistema de Informacion Energetica (SIE) cuadros - Balance nacional de gas natural / Demanda -
     https://sie.energia.gob.mx - and the annual Prospectiva de Gas Natural (demand by sector incl. residential,
     services, transport, industry, electric, oil sector).
  3. CENAGAS / SISTRANGAS (pipeline deliveries by user class; no end-use split).
CRE publishes permits and prices, not volumes by sector.

Sectors written (million m3/day; MMpcd x 0.0283168):
  Power, Industrial, Residential, Commercial, Transport_CNG, Distribution_companies (Pemex 'Distribuidoras'
  - residential + commercial + small industry + CNG, not split by Pemex), Pemex_own_use, Other, Total.
Only the sectors a resource actually publishes are filled; the rest stay empty (nothing is estimated).

Incremental: the committed workbook is the history store ('Raw (as published)' sheet). Each resource is
downloaded only when its CKAN last_modified (or HTTP Last-Modified) differs from the stamp recorded on the Units
sheet; archived months are kept and the last REFRESH_MONTHS are overwritten by the new file.
"""
import argparse
import io
import os
import re
import sys
import time
import unicodedata

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root
import xlsx_charts  # noqa: E402
import xlsx_notes  # noqa: E402

CKAN = "https://datos.gob.mx/api/3/action"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept-Language": "es-MX,es;q=0.9"}
T = (20, 180)
DATA_START = pd.Timestamp(2015, 1, 1)
REFRESH_MONTHS = 6
MCF_TO_M3 = 0.0283168466
SEARCHES = ["ventas internas gas natural sector", "gas natural demanda por sector", "balance nacional gas natural",
            "consumo gas natural sector", "ventas gas natural distribuidoras industrial electrico"]
# Resources confirmed by the discovery run go here as (name, direct file URL). Empty until verified.
PINNED = []
SECTOR_PATTERNS = [  # first match wins; patterns are on accent-stripped lower-case labels
    ("Total", r"^total|demanda total|ventas totales"),
    ("Power", r"electric|cfe|generacion|termoelectric|ciclo combinado"),
    ("Residential", r"residencial|domestic|hogar"),
    ("Commercial", r"comercial|servicios"),
    ("Transport_CNG", r"transporte|vehicular|gnv|gnc|autotransporte"),
    ("Distribution_companies", r"distribuidora"),
    ("Industrial", r"industri"),
    ("Pemex_own_use", r"pemex|petrolero|autoconsumo"),
    ("Other", r"^otros?\b"),
]
SECTORS = ["Power", "Industrial", "Residential", "Commercial", "Transport_CNG", "Distribution_companies",
           "Pemex_own_use", "Other"]
RAW_SHEET, DEMAND_SHEET = "Raw (as published)", "Demand by sector"
MES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
       "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
       "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6, "jul": 7, "ago": 8, "sep": 9, "set": 9,
       "oct": 10, "nov": 11, "dic": 12}


def out(*a):
    print(*a, flush=True)


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return " ".join(s.replace("\n", " ").split())


def get(url, **kw):
    for attempt in range(3):
        try:
            r = requests.get(url, headers=H, timeout=T, **kw)
            if r.status_code == 200:
                return r
            out(f"  {url}: HTTP {r.status_code}")
        except requests.RequestException as e:
            out(f"  {url}: {type(e).__name__}: {str(e)[:150]}")
        time.sleep(5 * (attempt + 1))
    return None


def sector_of(label):
    n = norm(label)
    for key, pat in SECTOR_PATTERNS:
        if re.search(pat, n):
            return key
    return None


def unit_factor(text):
    """Multiplier to million cubic feet per day from the unit words found in `text`, else None (resource skipped)."""
    n = norm(text)
    if re.search(r"mmmpcd|miles de millones de pies cubicos diarios|bcfd", n):
        return 1000.0
    if re.search(r"mmpcd|millones de pies cubicos diarios|mmcfd|mmscfd", n):
        return 1.0
    if re.search(r"millones de metros cubicos diarios|mmm3d|mmmcd|mcm/d", n):
        return 1.0 / MCF_TO_M3
    return None


# ------------------------------------------------------------------ discovery of resources

def catalogue_resources():
    """[(name, url, last_modified)] of csv/xls(x) resources on datos.gob.mx that look like gas-by-sector series."""
    seen, found = set(), []
    for q in SEARCHES:
        r = get(f"{CKAN}/package_search", params={"q": q, "rows": 30})
        if r is None:
            continue
        try:
            results = r.json()["result"]["results"]
        except (ValueError, KeyError):
            continue
        for p in results:
            if "gas" not in norm(p.get("title", "")):
                continue
            for res in p.get("resources", []):
                url = res.get("url") or ""
                if res.get("id") in seen or not re.search(r"\.(csv|xlsx?)($|\?)", url, re.I):
                    continue
                seen.add(res.get("id"))
                found.append((f"{p.get('title')} / {res.get('name')}", url, res.get("last_modified") or ""))
    out(f"catalogue: {len(found)} candidate resources")
    return found


# ------------------------------------------------------------------ parser

def read_frames(content, url):
    """[(sheet label, DataFrame header=None)] for a csv/xls/xlsx download."""
    if re.search(r"\.csv($|\?)", url, re.I):
        for enc in ("utf-8-sig", "latin-1"):
            try:
                return [("csv", pd.read_csv(io.BytesIO(content), header=None, encoding=enc, dtype=object,
                                            on_bad_lines="skip"))]
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
        return []
    return list(pd.read_excel(io.BytesIO(content), sheet_name=None, header=None).items())


def to_month(v):
    """Cell -> month Timestamp or None (datetime, 'YYYY-MM', 'ene-2024', 'enero 2024')."""
    if hasattr(v, "year") and hasattr(v, "month"):
        return pd.Timestamp(v).to_period("M").to_timestamp()
    s = norm(v)
    m = re.match(r"^(\d{4})[-/](\d{1,2})", s)
    if m:
        return pd.Timestamp(int(m.group(1)), int(m.group(2)), 1)
    m = re.match(r"^([a-z]+)[ .\-/]+(?:de )?(\d{4})$", s)
    if m and m.group(1) in MES:
        return pd.Timestamp(int(m.group(2)), MES[m.group(1)], 1)
    return None


def parse_table(df, label):
    """Wide layout only: one row per sector label, one column per month (the layout of SIE/Pemex cuadros).
    Returns Month x sector frame in MMpcd, or empty if no unit / no sector rows / no month header is recognised."""
    head = " ".join(norm(v) for v in df.head(12).to_numpy().ravel() if isinstance(v, str)) + " " + norm(label)
    factor = unit_factor(head)
    if factor is None:
        return pd.DataFrame()
    for hdr in range(min(len(df), 25)):
        months = {c: to_month(v) for c, v in df.iloc[hdr].items() if pd.notna(v)}
        months = {c: m for c, m in months.items() if m is not None}
        if len(months) < 6:
            continue
        rows = {}
        for i in range(hdr + 1, min(hdr + 40, len(df))):
            lab = next((v for v in df.iloc[i, :3] if isinstance(v, str) and v.strip()), None)
            key = sector_of(lab) if lab else None
            if key and key not in rows:
                rows[key] = pd.Series({m: pd.to_numeric(str(df.iloc[i][c]).replace(",", ""), errors="coerce")
                                       for c, m in months.items()}) * factor
        if rows:
            w = pd.DataFrame(rows).sort_index()
            w.index.name = "Month"
            return w[w.index >= DATA_START]
    return pd.DataFrame()


# ------------------------------------------------------------------ archive

def load_raw(path):
    try:
        df = pd.read_excel(path, sheet_name=RAW_SHEET, index_col=0)
        df.index = pd.to_datetime(df.index.astype(str))
        out(f"existing archive: {len(df)} months {df.index.min():%Y-%m}..{df.index.max():%Y-%m}")
        stamps = {}
        for line in pd.read_excel(path, sheet_name="Units")["Notes"].dropna().astype(str):
            m = re.match(r"Resource stamp \[(.+?)\]: (.*)$", line)
            if m:
                stamps[m.group(1)] = m.group(2)
        return df, stamps
    except (FileNotFoundError, ValueError) as e:
        out(f"no existing archive ({e})")
        return pd.DataFrame(), {}


def merge(old, new):
    if old.empty:
        return new.sort_index()
    recent = sorted(old.index)[-REFRESH_MONTHS:]
    base = old.drop(index=[m for m in recent if m in new.index])
    return new.combine_first(base).sort_index()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/mexico_gas_demand_by_sector.xlsx")
    args = ap.parse_args()
    raw, stamps = load_raw(args.out)
    new_stamps, used = dict(stamps), []
    candidates = [(n, u, "") for n, u in PINNED] + catalogue_resources()
    for name, url, lm in candidates:
        if lm and stamps.get(url) == lm and not raw.empty:
            out(f"unchanged since last run: {name}")
            continue
        r = get(url)
        if r is None:
            continue
        stamp = lm or r.headers.get("Last-Modified", "")
        try:
            frames = read_frames(r.content, url)
        except Exception as e:  # noqa: BLE001
            out(f"  unreadable {url}: {type(e).__name__}")
            continue
        got = []
        for lab, df in frames:
            w = parse_table(df, f"{name} {lab}")
            if not w.empty:
                got.append(w)
        if not got:
            out(f"  no recognised sector/unit table in {name}")
            continue
        w = pd.concat(got).groupby(level=0).first()
        out(f"  parsed {name}: {list(w.columns)} {w.index.min():%Y-%m}..{w.index.max():%Y-%m}")
        raw = merge(raw, w)
        new_stamps[url] = stamp
        used.append((name, url))
    if raw.empty:
        raise SystemExit("no Mexican gas-by-sector table found or archived - nothing written (see discovery workflow)")
    cols = [c for c in SECTORS + ["Total"] if c in raw]
    mcm = (raw[cols] * MCF_TO_M3).round(4).add_suffix("_mcm_per_day")
    table = pd.concat([mcm, raw[cols].round(3).add_suffix("_MMPCD")], axis=1)
    sheets = {DEMAND_SHEET: table, RAW_SHEET: raw[cols].round(4)}
    for df in sheets.values():
        df.index = pd.DatetimeIndex(df.index).strftime("%Y-%m")
        df.index.name = "Month"
    notes = [
        "UNITS",
        "*_mcm_per_day = million cubic metres per day (monthly average); *_MMPCD = million cubic feet per day. "
        "mcm/d = MMPCD x 0.0283168. Source units are converted from the unit stated in each file's header; a file "
        "without a recognised unit is skipped.",
        f"'{RAW_SHEET}': every sector series as published, MMPCD.",
        "",
        "SECTORS",
        "Power: sector electrico (CFE and permit holders). Industrial: industrial. Residential / Commercial: "
        "residencial / servicios (only where the source splits them). Transport_CNG: transporte / vehicular. "
        "Distribution_companies: Pemex 'Distribuidoras' (residential + commercial + small industry + CNG combined). "
        "Pemex_own_use: sector petrolero / Pemex. Only sectors the source publishes are filled; nothing is estimated.",
        "",
        "COVERAGE",
        f"Monthly {raw.index.min():%Y-%m}..{raw.index.max():%Y-%m}; last {REFRESH_MONTHS} archived months re-read "
        "when a resource changes.",
        "",
        "SOURCE",
        "Pemex ventas internas de gas natural por sector / SENER SIE (Balance nacional de gas natural), via the open "
        "data catalogue: https://datos.gob.mx ; SENER SIE: https://sie.energia.gob.mx",
    ] + [f"Resource used: {n} - {u}" for n, u in used] \
      + [f"Resource stamp [{u}]: {s}" for u, s in new_stamps.items() if s]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "SECTORS", "COVERAGE", "SOURCE"})
    chart = (raw[[c for c in SECTORS if c in raw]] * MCF_TO_M3).rename(columns=lambda c: c.replace("_", " "))
    xlsx_charts.add_chart_sheet(args.out, chart.dropna(how="all"), "Mexico gas demand by sector", "million m3/day",
                                kind="stacked_bar")
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
