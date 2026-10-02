"""
US <-> Mexico natural gas pipeline capacity by line, annual, from EIA's "State to State Capacity" workbook
(EIA-StatetoStateCapacity_<Mon><Year>.xlsx, sheet "Pipeline State2State CapacityH", MMcf/d). Every pipeline row
whose destination is Mexico is a US export line; rows whose origin is Mexico are imports into the US.

Writes us_mexico_pipeline_capacity.xlsx:
  Export capacity by line   year x pipeline, MMcf/d, US -> Mexico  (+ Total)
  Import capacity by line   year x pipeline, MMcf/d, Mexico -> US  (+ Total)
  Raw                       every Mexico row of EIA's sheet: pipeline, state from, state to, year, MMcf/d
  Units                     notes, plus the EIA release file the data came from (add_charts.py then adds the
                            charts, including monthly exports against capacity from mexico_gas.xlsx)

Incremental / monthly: EIA posts a new file a few times a year. The script reads EIA's natural-gas data page, picks
the newest StatetoStateCapacity file and downloads it only when it differs from the release saved on the Units
sheet (so most runs just re-save nothing).

Usage: python3 US_MEXICO_PIPELINE_CAPACITY.py [--out "output/Data and Chart Outputs/us_mexico_pipeline_capacity.xlsx"] [--force]
"""
import argparse
import io
import os
import re
import sys

import pandas as pd
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

PAGE = "https://www.eia.gov/naturalgas/data.php"
BASE = "https://www.eia.gov/naturalgas/pipelines/"
SHEET = "Pipeline State2State CapacityH"
UA = {"User-Agent": "Mozilla/5.0 (energy-data research)"}
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "us_mexico_pipeline_capacity.xlsx")
MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]


def latest_release():
    """Newest EIA-StatetoStateCapacity file name listed on EIA's natural gas data page."""
    html = requests.get(PAGE, headers=UA, timeout=60).text
    found = set(re.findall(r"EIA-StatetoStateCapacity_([A-Za-z]{3})[a-z]*(\d{4})\.xlsx", html))
    if not found:
        raise RuntimeError("no EIA-StatetoStateCapacity file listed on " + PAGE)
    mon, yr = max(found, key=lambda t: (int(t[1]), MONTHS.index(t[0].lower())))
    return f"EIA-StatetoStateCapacity_{mon}{yr}.xlsx"


def saved_release(out):
    try:
        units = pd.read_excel(out, sheet_name="Units")
        for line in units.iloc[:, 0].astype(str):
            m = re.match(r"EIA release file: (\S+)", line)
            if m:
                return m.group(1)
    except Exception:  # noqa: BLE001
        pass
    return None


def parse(content):
    raw = pd.read_excel(io.BytesIO(content), sheet_name=SHEET, header=None)
    hdr = raw.index[raw[0].astype(str).str.strip().eq("Pipeline")][0]
    years = pd.to_numeric(raw.iloc[hdr, 3:], errors="coerce")
    d = raw.iloc[hdr + 1:, :3 + len(years)].copy()
    d.columns = ["Pipeline", "State from", "State to"] + [int(y) if pd.notna(y) else f"c{i}" for i, y in enumerate(years)]
    d = d.drop(columns=[c for c in d.columns if str(c).startswith("c")])
    d["Pipeline"] = d["Pipeline"].ffill()               # pivot layout: names and origin states appear once per group
    d["State from"] = d.groupby("Pipeline")["State from"].ffill()
    d = d[d["Pipeline"].notna() & ~d["Pipeline"].astype(str).str.contains("Grand Total", case=False)]
    return d[(d["State to"] == "Mexico") | (d["State from"] == "Mexico")]


def by_line(d, side):
    sub = d[d["State to"] == "Mexico"] if side == "export" else d[d["State from"] == "Mexico"]
    yrs = [c for c in sub.columns if isinstance(c, int)]
    g = sub.groupby("Pipeline")[yrs].sum(min_count=1).T            # year x pipeline
    g = g.dropna(how="all").fillna(0.0)
    g = g.loc[:, (g > 0).any()]
    g.index = pd.to_datetime(g.index.astype(str) + "-01-01")
    g = g.sort_index()
    g.index.name = "Year"
    g["Total"] = g.sum(axis=1)
    return g


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--force", action="store_true", help="re-download even if the release is unchanged")
    args = ap.parse_args()

    release = latest_release()
    if not args.force and os.path.exists(args.out) and saved_release(args.out) == release:
        print(f"{release}: already saved, nothing to do", flush=True)
        return
    r = requests.get(BASE + release, headers=UA, timeout=120)
    r.raise_for_status()
    d = parse(r.content)
    exp, imp = by_line(d, "export"), by_line(d, "import")
    yrs = [c for c in d.columns if isinstance(c, int)]
    raw = d.melt(id_vars=["Pipeline", "State from", "State to"], value_vars=yrs, var_name="Year",
                 value_name="Capacity MMcf/d").dropna(subset=["Capacity MMcf/d"]).sort_values(["Pipeline", "Year"])
    notes = [
        "UNITS",
        f"EIA release file: {release}",
        f"{BASE}{release}  (listed on {PAGE})",
        "Capacity is in million cubic feet per day (MMcf/d), annual, one value per pipeline per year, as EIA's "
        "State to State Capacity workbook gives it. Divide by 1,000 for Bcf/d (the charts do).",
        "Export lines are rows whose destination is Mexico; import lines are rows whose origin is Mexico "
        "(Kinder Morgan Border, Tennessee Gas, Texas Eastern, North Baja south-to-north segment).",
        "This is design / contracted transport capacity by line, not daily flow or operationally available "
        "capacity. Monthly US pipeline exports to Mexico (mexico_gas.xlsx, EIA N9132MX2) are charted against it.",
        "A pipeline's origin state is filled down from the line above, as EIA's pivot layout only prints it once.",
        "",
        "SOURCE",
        "U.S. Energy Information Administration, Natural Gas Pipelines: State to State Capacity.",
        "Updated monthly (1st); the file is downloaded only when EIA lists a new release.",
    ]
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Export capacity by line": exp, "Import capacity by line": imp, "Raw": raw.set_index("Pipeline")},
                              notes, ["UNITS", "SOURCE"])
    print(f"{release}: {len(exp.columns) - 1} export lines, {len(imp.columns) - 1} import lines, "
          f"latest year {exp.index.max().year}, total {exp['Total'].iloc[-1]:.0f} MMcf/d", flush=True)


if __name__ == "__main__":
    main()
