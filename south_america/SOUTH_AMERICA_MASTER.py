"""
South America master workbook: one file with every South American
dataset this repo pulls, and a Dashboard front page carrying all their
charts.

  Dashboard          - title, index of charts (latest month, data tab),
                       then every chart (native Excel, no borders, mmm/yy)
  <CC> <chart> data  - the table each Dashboard chart plots
  <CC> <dataset> raw - the full data sheet from each source workbook
  Sources            - where each dataset comes from, units and notes

Reads (doesn't refetch) the workbooks the scheduled South America pulls
write to "output/Data and Chart Outputs/". Chart definitions come from
add_charts.py's registry, so the master and each country workbook always
show the same charts. A missing input is listed on the Dashboard and
skipped rather than stopping the rest.

Usage: python3 SOUTH_AMERICA_MASTER.py [--out "output/Data and Chart Outputs/south_america_master.xlsx"]
"""
import argparse
import os
import sys
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import add_charts  # noqa: E402
import xlsx_charts  # noqa: E402

DATA_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
DARK_BLUE = "17365D"

# (country code, country, workbook, raw sheet to copy, short dataset name)
DATASETS = [
    ("AR", "Argentina", "argentina_gas_monthly.xlsx", "National", "gas"),
    ("BR", "Brazil", "brazil_gas_monthly.xlsx", "Demand by segment", "gas"),
    ("BO", "Bolivia", "bolivia_gas_demand_by_sector.xlsx", "Demand by sector", "gas"),
    ("CL", "Chile", "chile_gas_imports.xlsx", "Gas imports", "gas imports"),
    ("CL", "Chile", "chile_power_by_type.xlsx", "Generation by type", "power"),
    ("CO", "Colombia", "colombia_gas_demand_by_sector.xlsx", "Demand by sector", "gas"),
    ("EC", "Ecuador", "ecuador_gas.xlsx", "Gas by use", "gas"),
]

CHART_W, CHART_H = 17.0, 8.5       # cm
ROWS_PER_CHART = 19
COLS = ("B", "G")                  # two charts per row (B-E are the index columns, ~17 cm wide)


def sheet_name(text, used):
    name = text.replace("/", "-")[:31]
    base, k = name, 2
    while name in used:
        name = f"{base[:28]} {k}"
        k += 1
    used.add(name)
    return name


def write_frame(ws, df):
    ws.append([str(c) for c in df.columns])
    for row in df.itertuples(index=False):
        ws.append([None if (not isinstance(v, str) and pd.isna(v)) else
                   (v.to_pydatetime() if isinstance(v, pd.Timestamp) else v) for v in row])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"


def notes_text(path):
    try:
        u = pd.read_excel(path, sheet_name=0, header=None)
        return " ".join(str(x) for x in u.iloc[:, 0].dropna().tolist())
    except Exception:
        return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(DATA_DIR, "south_america_master.xlsx"))
    ap.add_argument("--data-dir", default=DATA_DIR)
    args = ap.parse_args()

    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    used = {"Dashboard", "Sources"}
    charts, index_rows, missing, sources = [], [], [], []

    for code, country, fname, raw_sheet, short in DATASETS:
        path = os.path.join(args.data_dir, fname)
        if not os.path.exists(path):
            missing.append(f"{country} {short} ({fname})")
            continue
        try:
            specs = add_charts.REGISTRY.get(fname, add_charts.generic)(path)
        except Exception as e:
            missing.append(f"{country} {short} ({fname}: {type(e).__name__}: {e})")
            continue
        for s in specs:
            if "water_year" in s:
                continue
            df, n_bars = xlsx_charts.prepare(s["df"].dropna(how="all"), s.get("line_cols", ()))
            if df.empty:
                continue
            ws = wb.create_sheet(sheet_name(f"{code} {s['name']} data", used))
            xlsx_charts.write_table(ws, df, s["date_format"])
            ch = xlsx_charts.build_chart(ws, df, n_bars, s["title"], s["units"], s["kind"],
                                         s["date_format"], width=CHART_W, height=CHART_H)
            charts.append(ch)
            index_rows.append((country, s["title"], df.index.max().strftime("%b/%y"), ws.title))
        try:
            raw = add_charts.read(path, raw_sheet)
            ws_raw = wb.create_sheet(sheet_name(f"{code} {short} raw", used))
            write_frame(ws_raw, raw)
        except Exception as e:
            missing.append(f"{country} {short} raw sheet ({e})")
        sources.append((country, short, fname, notes_text(path)))

    # ---- Dashboard
    dash["B1"] = "South America energy - master dashboard"
    dash["B1"].font = Font(bold=True, size=18, color=DARK_BLUE)
    dash["B2"] = (f"Built {date.today():%d %b %Y} from the scheduled pulls in output/Data and Chart Outputs. "
                  "Every chart's data is on its own tab; full source data on the 'raw' tabs; units on 'Sources'.")
    dash["B2"].font = Font(italic=True, color="6B6B6B")
    hdr = ["Country", "Chart", "Latest", "Data tab"]
    for j, h in enumerate(hdr):
        c = dash.cell(row=4, column=2 + j, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=DARK_BLUE)
    for i, r in enumerate(index_rows, start=5):
        for j, v in enumerate(r):
            dash.cell(row=i, column=2 + j, value=v)
        dash.cell(row=i, column=5).hyperlink = f"#'{r[3]}'!A1"
        dash.cell(row=i, column=5).font = Font(color="2A78D6", underline="single")
    row = 5 + len(index_rows) + 1
    if missing:
        dash.cell(row=row, column=2, value="Not available this run: " + "; ".join(missing)).font = Font(color="E34948")
        row += 1
    for col, w in (("B", 12), ("C", 48), ("D", 9), ("E", 24)):
        dash.column_dimensions[col].width = w
    start = row + 1
    for k, ch in enumerate(charts):
        r = start + (k // 2) * ROWS_PER_CHART
        dash.add_chart(ch, f"{COLS[k % 2]}{r}")

    # ---- Sources
    src = wb.create_sheet("Sources")
    src.append(["Country", "Dataset", "Workbook", "Units and notes (from the source workbook)"])
    for cell in src[1]:
        cell.font = Font(bold=True)
    for s in sources:
        src.append(list(s))
    src.column_dimensions["A"].width = 12
    src.column_dimensions["B"].width = 14
    src.column_dimensions["C"].width = 38
    src.column_dimensions["D"].width = 160

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_charts.save_atomic(wb, args.out)
    print(f"Saved {args.out}: {len(charts)} charts, tabs {wb.sheetnames}")
    if missing:
        print("missing:", missing)


if __name__ == "__main__":
    main()
