"""
Tasmania hydro storage: Hydro Tasmania's weekly energy in storage (GWh) by
lake, from its public history workbook:

  https://www.hydro.com.au/docs/energyinstorage/download/EnergyInStorage-HistoricalData.xls

Writes au_hydro_storage.xlsx, sheet "Weekly": date, one GWh column per lake,
Total_GWh and Total_pct (of the full-supply energy in storage the workbook
lists). Charted as an Oct-Sep water-year chart.

The workbook is small and Hydro Tasmania revises it in place, so each run
reads it whole and merges it over the saved rows (weeks that drop out of
the published file are kept).

Usage: python3 AU_HYDRO_TAS_STORAGE.py [--out "output/Data and Chart Outputs/au_hydro_storage.xlsx"]
"""
import argparse
import io
import os
import re
import sys

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

URL = "https://www.hydro.com.au/docs/energyinstorage/download/EnergyInStorage-HistoricalData.xls"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0"}
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "au_hydro_storage.xlsx")


def parse(content):
    """Sheet1: title rows, catchment row, lake-name rows, a 'Full Supply' row, then one dated row per week."""
    raw = pd.read_excel(io.BytesIO(content), sheet_name=0, header=None)
    dates = pd.to_datetime(raw.iloc[:, 0], errors="coerce")
    first = dates.first_valid_index()
    full_row = next(i for i in range(first) if str(raw.iat[i, 0]).strip().lower().startswith("full supply"))
    names = {}
    for c in range(1, raw.shape[1]):
        label = [str(raw.iat[r, c]).strip() for r in range(full_row) if pd.notna(raw.iat[r, c])]
        # a column can hold several storages listed on successive rows (e.g. Lake St. Clair + Lake King William)
        lake = " + ".join(x for x in label if re.search(r"lake|lagoon|pond|total|reservoir|dam", x, re.I)) or (
            label[-1] if label else None)
        if lake and pd.notna(pd.to_numeric(raw.iat[full_row, c], errors="coerce")):
            names[c] = re.sub(r"\s+", " ", lake)
    body = raw.loc[dates.notna(), list(names)].apply(pd.to_numeric, errors="coerce")
    body.index = dates[dates.notna()]
    body.columns = [names[c] for c in names]
    full = pd.to_numeric(raw.loc[full_row, list(names)], errors="coerce")
    full.index = body.columns
    body = body.T.groupby(level=0, sort=False).sum(min_count=1).T   # a lake split over two columns
    full = full.groupby(level=0, sort=False).sum()
    totals = [c for c in body.columns if re.search(r"total", c, re.I)]
    lakes = [c for c in body.columns if c not in totals]
    return body, full, lakes, totals


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()

    r = requests.get(URL, headers=HEADERS, timeout=(10, 120))
    r.raise_for_status()
    body, full, lakes, totals = parse(r.content)
    print(f"{len(body)} weeks {body.index.min():%Y-%m-%d}..{body.index.max():%Y-%m-%d}; lakes {lakes}; "
          f"total columns {totals}", flush=True)
    w = body[lakes].add_suffix("_GWh")
    w["Total_GWh"] = body[totals[0]] if totals else body[lakes].sum(axis=1, min_count=1)
    cap = full[totals[0]] if totals else full[lakes].sum()
    w["Total_pct"] = (w["Total_GWh"] / cap * 100).round(1)
    try:
        old = pd.read_excel(args.out, sheet_name="Weekly", index_col=0)
        old.index = pd.to_datetime(old.index, errors="coerce")
        w = w.combine_first(old[old.index.notna()])[w.columns]
    except (FileNotFoundError, ValueError, KeyError, OSError):
        pass
    w.index.name = "date"
    print(w.tail(3).to_string(), flush=True)
    notes = [
        "UNITS",
        "GWh of energy in storage (water stored x the generation it can produce downstream), weekly, by lake; "
        f"Total_GWh = all storages; Total_pct = Total_GWh / full-supply energy in storage ({cap:,.0f} GWh).",
        "",
        "COVERAGE",
        f"Tasmania (Hydro Tasmania's storages - the NEM's largest hydro system), weekly, {w.index.min():%d %b %Y} to "
        f"{w.index.max():%d %b %Y}. Snowy Hydro (NSW/VIC) does not publish a storage series.",
        "",
        "SOURCE",
        f"Hydro Tasmania, Energy in Storage historical data: {URL}",
        "https://www.hydro.com.au/water/energy-in-storage",
    ]
    xlsx_notes.write_workbook(args.out, {"Weekly": w.round(2)}, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}", flush=True)


if __name__ == "__main__":
    main()
