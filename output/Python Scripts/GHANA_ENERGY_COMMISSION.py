"""
Ghana electricity generation (annual) and Akosombo / Bui lake levels (monthly), from the Energy Commission's
annual 'National Energy Statistics' PDF.

Why this source: GRIDCo (the grid operator, gridcogh.com) and VRA publish no data tables or dashboards - only annual-report
PDFs and news - and the Energy Commission's weekly wholesale-market statistics are PDF reports (see
discovery_archive/ghana/ and the workbook notes). The Energy Commission (energycom.gov.gh), Ghana's electricity
regulator and planner, publishes each year a statistics booklet built from GRIDCo, VRA, ECG and IPP returns:
  Table 'Annual Electricity Generation'  GWh by Hydro / Thermal / Other renewables, 2000 - last year
  Table 'Akosombo Dam Month-End Elevation (feet)'   monthly, 2000 - last year
  Table 'Bui Dam Month-End Elevation (feet)'         monthly, from Bui's first year
The list page https://www.energycom.gov.gh/index.php/planning/energy-statistics names each booklet
('2026 Energy Statistics' holds data to Dec 2025); download links are '?download=<id>:<slug>'.

Incremental: the PDF is downloaded only when a booklet with a higher id (or a different file size) than the one
recorded on the 'Release' sheet appears. Saved years are kept and the new booklet's rows replace the same years
(the Commission revises the last year).

Usage: python3 GHANA_ENERGY_COMMISSION.py [--out PATH] [--force]
"""
print("STARTING", flush=True)

import argparse
import io
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes  # noqa: E402

OUT = "output/Data and Chart Outputs/ghana_energy_commission_power.xlsx"
BASE = "https://www.energycom.gov.gh"
LIST_URL = BASE + "/index.php/planning/energy-statistics"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def list_booklets():
    """[(edition year, download id, slug)] newest edition first, from the list page."""
    r = requests.get(LIST_URL, headers=H, timeout=(10, 60))
    r.raise_for_status()
    found = set(re.findall(r"energy-statistics\?download=(\d+):(\d{4})-energy-statistics", r.text))
    out = sorted({(int(y), int(i)) for i, y in found}, reverse=True)
    return [(y, i, f"{i}:{y}-energy-statistics") for y, i in out]


def num(tok):
    tok = tok.replace(",", "").strip()
    if tok in ("-", "–", "—", ""):
        return None
    try:
        return float(tok)
    except ValueError:
        return "bad"


def parse_annual(text):
    """Rows 'YEAR hydro thermal [other|-] total share...' of the Annual Electricity Generation table."""
    rows = {}
    for line in text.split("\n"):
        m = re.match(r"^(20\d\d)\s+(.*)$", line.strip())
        if not m:
            continue
        toks = m.group(2).split()
        vals = [num(t) for t in toks]
        if "bad" in vals or len(vals) < 7:
            continue
        hydro, thermal, other, total = vals[0], vals[1], vals[2], vals[3]
        if hydro is None or thermal is None or total is None:
            continue
        s = hydro + thermal + (other or 0.0)
        if abs(s - total) > max(3.0, 0.002 * total):
            continue
        rows[int(m.group(1))] = {"Hydro_GWh": hydro, "Thermal_GWh": thermal, "Other_Renewables_GWh": other,
                                 "Total_GWh": total}
    return rows


def parse_levels(text):
    """Rows 'YEAR v1 .. v12' (fewer values for a part-year; months in order) -> {Timestamp month start: feet}."""
    out = {}
    for line in text.split("\n"):
        m = re.match(r"^(\d{4})\s+((?:\d{2,4}(?:\.\d+)?\s*)+)$", line.strip())
        if not m:
            continue
        y = int(m.group(1))
        vals = [float(v) for v in m.group(2).split()]
        if not (1990 <= y <= 2100) or not (1 <= len(vals) <= 12):
            continue
        for k, v in enumerate(vals):
            out[pd.Timestamp(y, k + 1, 1)] = v
    return out


def read_booklet(pdf_bytes):
    import pdfplumber
    annual, akos, bui = {}, {}, {}
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for n, page in enumerate(pdf.pages):
            tx = page.extract_text() or ""
            head = tx[:600]
            if re.search(r"Table 3\.\d+: Annual Electricity Generation", tx) and "Hydro" in tx and "Share" in tx:
                annual.update(parse_annual(tx))
            # level tables: heading names the dam; a table can run over two pages, so the page after a heading is
            # read for the dam of the last heading seen
            for name, store in (("Akosombo", akos), ("Bui", bui)):
                if re.search(rf"Table 3\.\d+: {name} Dam Month-End Elevation", tx):
                    store.update(parse_levels(tx.split(f"{name} Dam Month-End Elevation")[1]))
    return annual, akos, bui


def load(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index, errors="coerce")
    return df[df.index.notna()]


NOTES = [
    "UNITS",
    "Annual: electricity generation in GWh per CALENDAR year, one row per year dated 1 January: Hydro (Akosombo, Kpong, Bui, "
    "Tsatsadu), Thermal (gas, oil, LCO), Other_Renewables (utility solar and biogas, from 2013) and Total. 'Other_Renewables' is "
    "blank before 2013 (no such plants); grid-connected generation as reported by GRIDCo and ECG, so rooftop / "
    "behind-the-meter solar is not in it.",
    "Akosombo, Bui: reservoir elevation in FEET above sea level at the END of each month, one row per month dated "
    "the 1st of the month. The Commission prints whole feet. Akosombo minimum operating level 240 ft; Bui 550 ft.",
    "Release: which booklet the data came from (list-page edition, download id, file size, checked date).",
    "",
    "COVERAGE",
    "Annual from 2000; Akosombo monthly from Jan 2000; Bui monthly from its first printed year. The booklet appears about "
    "mid-year and covers the previous calendar year ('2026 Energy Statistics' = data to Dec 2025), so the lag is 6-18 months. "
    "There is NO sub-annual generation here; GRIDCo's daily dispatch figures are not published as data.",
    "",
    "SOURCE",
    "Energy Commission of Ghana, National Energy Statistics (annual booklet; tables 'Annual Electricity Generation', "
    f"'Akosombo Dam Month-End Elevation', 'Bui Dam Month-End Elevation'; source credited to GRIDCo, ECG, VRA): {LIST_URL}",
    "Pulled on the 1st and 15th by GitHub Actions (ghana_energy_commission.yml); the PDF is downloaded only when a new "
    "booklet appears.",
]
TITLES = ["UNITS", "COVERAGE", "SOURCE"]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--force", action="store_true", help="download even if the booklet is unchanged")
    args = ap.parse_args()

    books = list_booklets()
    print("Booklets on the list page:", books[:6], flush=True)
    if not books:
        print("No Energy Statistics download links found.", flush=True)
        sys.exit(1)
    edition, dl_id, slug = books[0]
    old_rel = load_release(args.out)
    if not args.force and old_rel and old_rel.get("download_id") == dl_id:
        print(f"Booklet {edition} (id {dl_id}) already saved - nothing to download.", flush=True)
        return
    url = f"{LIST_URL}?download={slug}"
    r = requests.get(url, headers=H, timeout=(10, 180))
    r.raise_for_status()
    if r.content[:4] != b"%PDF":
        print("Download is not a PDF:", r.headers.get("content-type"), flush=True)
        sys.exit(1)
    print(f"Downloaded {len(r.content):,} bytes ({r.headers.get('content-disposition')})", flush=True)
    annual, akos, bui = read_booklet(r.content)
    print(f"Parsed: annual {len(annual)} years, Akosombo {len(akos)} months, Bui {len(bui)} months", flush=True)
    if len(annual) < 20 or len(akos) < 200:
        print("Parse looks incomplete - not writing.", flush=True)
        sys.exit(1)

    a_new = pd.DataFrame.from_dict({pd.Timestamp(y, 1, 1): v for y, v in annual.items()}, orient="index").sort_index()
    sheets = {}
    for name, new in (("Annual", a_new),
                      ("Akosombo", pd.DataFrame({"Level_ft": pd.Series(akos)}).sort_index()),
                      ("Bui", pd.DataFrame({"Level_ft": pd.Series(bui)}).sort_index())):
        old = load(args.out, name)
        if not old.empty:
            keep = old[~old.index.isin(new.index)]
            new = pd.concat([keep, new]).sort_index()
        new.index.name = "date"
        sheets[name] = new
    sheets["Release"] = pd.DataFrame({"value": {"edition": edition, "download_id": dl_id, "bytes": len(r.content),
                                                 "checked": pd.Timestamp.now("UTC").strftime("%Y-%m-%d"),
                                                 "url": url}})
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, NOTES, TITLES)
    print(sheets["Annual"].tail(6).to_string(), flush=True)
    print(sheets["Akosombo"].tail(6).to_string(), flush=True)
    print(sheets["Bui"].tail(6).to_string(), flush=True)


def load_release(path):
    try:
        df = pd.read_excel(path, sheet_name="Release", index_col=0)
        return {k: (int(v) if str(v).isdigit() else v) for k, v in df["value"].items()}
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return {}


if __name__ == "__main__":
    main()
