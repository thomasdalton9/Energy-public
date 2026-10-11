"""
Venezuela natural gas -> venezuela_gas.xlsx.

FINDING (research, Oct 2026): there is NO obtainable monthly gas demand-by-sector series for Venezuela.
  - PDVSA stopped publishing its operating statistics (Informe de Gestion Anual; Pdvsa en Cifras) after ~2016;
    those reports gave gas by use (electrico, siderurgico/aluminio, petroquimico, domestico/comercial, own use)
    only annually and only as PDF tables. MPPEP / MPPPE (ministry) publishes no sector volumes.
  - OPEC Annual Statistical Bulletin gives Venezuela gas production / marketed / (some years) consumption,
    annual, no sectors. Energy Institute Statistical Review: annual total consumption, no sectors.
  - JODI-Gas (jodidata.org) takes Venezuela's monthly submissions (production, imports, exports, stock
    change, own use / other, TOTAL DEMAND), no end-use sectors, with long gaps in Venezuela's reporting.
So this script keeps the best series that exists, clearly labelled as a FALLBACK: JODI-Gas monthly flows for
Venezuela exactly as reported (every flow kept, nothing derived apart from the per-day conversion), and it does
NOT invent a sector split. Sector columns the user wants (power, industrial, residential, commercial, CNG)
would need a PDVSA/MPPEE publication that is not currently released; see the Units sheet.

Not reachable from the editing sandbox (jodidata.org 403 via proxy); runs in GitHub Actions. The JODI download
link is read from the data-downloads page, and the CSV is parsed by header name (REF_AREA, TIME_PERIOD,
ENERGY_PRODUCT, FLOW_BREAKDOWN, UNIT_MEASURE, OBS_VALUE, ASSESSMENT_CODE). If the layout differs the run stops
without writing anything. Run discovery_archive/south_america/VENEZUELA_GAS_DISCOVERY.py first to confirm.

Incremental: the committed workbook ('JODI raw' sheet) is the history. JODI publishes whole files, so the zip is
downloaded only when its Last-Modified differs from the stamp on the Units sheet; every month in the file is
then taken from it (the file holds the full history, so revisions are picked up).
"""
import argparse
import io
import os
import re
import sys
import zipfile

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root
import xlsx_charts  # noqa: E402
import xlsx_notes  # noqa: E402

PAGES = ["https://www.jodidata.org/gas/database/data-downloads.aspx",
         "https://www.jodidata.org/gas/database/data-downloads"]
FALLBACK_ZIPS = ["https://www.jodidata.org/_resources/files/downloads/gas-data/jodi_gas_csv_beta.zip",
                 "https://www.jodidata.org/_resources/files/downloads/gas-data/jodi_gas_csv.zip"]
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
T = (20, 300)
AREA = "VE"
RAW_SHEET, DEMAND_SHEET = "JODI raw", "Monthly flows"


def out(*a):
    print(*a, flush=True)


def find_zip():
    """URL of the JODI-Gas CSV zip: link on the downloads page, else the known fallback names."""
    for page in PAGES:
        try:
            r = requests.get(page, headers=H, timeout=T)
        except requests.RequestException as e:
            out(f"  {page}: {type(e).__name__}")
            continue
        if r.status_code != 200:
            out(f"  {page}: HTTP {r.status_code}")
            continue
        links = re.findall(r'href="([^"]+\.zip)"', r.text, re.I)
        links = [l if l.startswith("http") else "https://www.jodidata.org" + l for l in links]
        hit = [l for l in links if "csv" in l.lower() and "gas" in l.lower()]
        if hit:
            return hit[0]
    return FALLBACK_ZIPS[0]


def load_existing(path):
    try:
        raw = pd.read_excel(path, sheet_name=RAW_SHEET, index_col=0)
        stamp = ""
        for line in pd.read_excel(path, sheet_name="Units")["Notes"].dropna().astype(str):
            m = re.match(r"JODI file Last-Modified: (.*)$", line)
            if m:
                stamp = m.group(1)
        return raw, stamp
    except (FileNotFoundError, ValueError):
        return pd.DataFrame(), ""


def parse_jodi(content):
    """JODI-Gas zip -> long frame for Venezuela: Month, flow, unit, value, assessment."""
    z = zipfile.ZipFile(io.BytesIO(content))
    name = next(n for n in z.namelist() if n.lower().endswith(".csv"))
    df = pd.read_csv(z.open(name), dtype=str)
    df.columns = [c.strip().lower() for c in df.columns]
    need = {"ref_area", "time_period", "flow_breakdown", "unit_measure", "obs_value"}
    if not need <= set(df.columns):
        raise SystemExit(f"unexpected JODI layout: {list(df.columns)}")
    v = df[df["ref_area"] == AREA].copy()
    if "energy_product" in v:
        v = v[v["energy_product"].str.upper().str.contains("NATGAS|NATURAL", na=False)]
    v["Month"] = pd.to_datetime(v["time_period"], format="%Y-%m", errors="coerce")
    v["value"] = pd.to_numeric(v["obs_value"], errors="coerce")  # JODI '..' / 'x' placeholders -> NaN
    v = v.dropna(subset=["Month", "value"])
    out(f"Venezuela rows: {len(v)}; flows {sorted(v['flow_breakdown'].unique())}; units {sorted(v['unit_measure'].unique())}")
    return v


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/venezuela_gas.xlsx")
    args = ap.parse_args()
    raw, stamp = load_existing(args.out)
    url = find_zip()
    out(f"JODI zip: {url}")
    try:
        head = requests.head(url, headers=H, timeout=T, allow_redirects=True)
        new_stamp = head.headers.get("Last-Modified", "")
    except requests.RequestException as e:
        out(f"  HEAD failed: {type(e).__name__}")
        new_stamp = ""
    if new_stamp and new_stamp == stamp and not raw.empty:
        out("JODI file unchanged since last run - nothing to do")
        return
    r = requests.get(url, headers=H, timeout=T)
    if r.status_code != 200:
        raise SystemExit(f"JODI download failed: HTTP {r.status_code}")
    new_stamp = new_stamp or r.headers.get("Last-Modified", "")
    v = parse_jodi(r.content)
    if v.empty:
        raise SystemExit("no Venezuela natural gas rows in the JODI file - nothing written")
    v["series"] = v["flow_breakdown"] + "|" + v["unit_measure"]
    wide = v.pivot_table(index="Month", columns="series", values="value", aggfunc="first").sort_index()
    wide.index = wide.index.strftime("%Y-%m")
    wide.index.name = "Month"
    sheets = {RAW_SHEET: wide}
    # per-day view only for flows reported in M3 (JODI-Gas: million cubic metres per month)
    m3 = [c for c in wide.columns if c.endswith("|M3")]
    dates = pd.to_datetime(wide.index)
    if m3:
        pd_ = wide[m3].div(dates.days_in_month, axis=0).round(4)
        pd_.columns = [c.split("|")[0] + "_mcm_per_day" for c in m3]
        sheets = {DEMAND_SHEET: pd_, **sheets}
    notes = [
        "UNITS",
        "'JODI raw': every Venezuela natural-gas flow as reported to JODI-Gas, one column per FLOW|UNIT, in the "
        "unit stated in the column name (M3 = million cubic metres per month; TJ = terajoules per month). "
        "'Monthly flows': the M3 flows divided by days in the month = million m3/day. JODI flow codes are kept "
        "as published (INDPROD production, TOTIMPSB imports, TOTEXPSB exports, STOCKCH stock change, "
        "TOTDEMO total demand / inland consumption, ...).",
        "",
        "SECTORS",
        "NOT AVAILABLE. No raw source publishes Venezuelan gas demand by sector (power, industrial, residential, "
        "commercial, CNG): PDVSA stopped issuing its annual reports (last gave gas by use ~2016), the ministry "
        "publishes none, OPEC/Energy Institute are annual totals. This workbook is the JODI total-demand "
        "FALLBACK, not a sector split.",
        "",
        "COVERAGE",
        f"Monthly {wide.index.min()}..{wide.index.max()} where Venezuela reported to JODI (gaps are Venezuela's "
        "non-reporting, not a pull failure).",
        "",
        "SOURCE",
        f"JODI-Gas World Database (FALLBACK, Venezuela's own submissions to JODI): {url}",
        f"JODI file Last-Modified: {new_stamp}",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "SECTORS", "COVERAGE", "SOURCE"})
    if m3:
        d = sheets[DEMAND_SHEET].copy()
        d.index = pd.to_datetime(d.index)
        d = d.rename(columns=lambda c: c.replace("_mcm_per_day", ""))
        keep = [c for c in d.columns if d[c].notna().sum() > 3]
        xlsx_charts.add_chart_sheet(args.out, d[keep].dropna(how="all"), "Venezuela gas flows (JODI)",
                                    "million m3/day", kind="line")
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
