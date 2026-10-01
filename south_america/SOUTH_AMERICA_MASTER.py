"""
South America master workbook: one file with every South American
dataset this repo pulls, and a Dashboard front page carrying all their
charts.

  Dashboard          - gas: title, index of charts (latest month, data
                       tab), then every chart (native Excel, no borders, mmm/yy)
  Dashboard - Power & Hydro - power generation by type for every country
                       Ember covers, plus reservoir water-year charts
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
import water_year_chart  # noqa: E402
import xlsx_charts  # noqa: E402

DATA_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
DARK_BLUE = "17365D"

# (country code, country, workbook, raw sheet to copy, short dataset name)
DATASETS = [
    ("AR", "Argentina", "argentina_gas_monthly.xlsx", "National", "gas"),
    ("BR", "Brazil", "brazil_gas_monthly.xlsx", "Demand by segment", "gas"),
    ("BO", "Bolivia", "bolivia_gas_demand_by_sector.xlsx", "Demand by sector", "gas"),
    ("CL", "Chile", "chile_gas_imports.xlsx", "Gas imports", "gas imports"),
    ("CO", "Colombia", "colombia_gas_demand_by_sector.xlsx", "Demand by sector", "gas"),
    ("EC", "Ecuador", "ecuador_gas.xlsx", "Gas by use", "gas"),
]

# Second dashboard: (code, country, workbook, raw sheet or "*" for every data sheet, short name)
POWER_HYDRO_DATASETS = [
    ("SA", "South America", "south_america_power_by_type.xlsx", "*", "power"),
    ("BR", "Brazil", "brazil_hydro_reservoirs.xlsx", "Daily", "hydro"),
    ("CO", "Colombia", "colombia_hydro_reservoirs.xlsx", "Daily", "hydro"),
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


def collect(wb, datasets, data_dir, used, sources):
    """Data + raw tabs for one dashboard's datasets; returns (charts, index rows, missing)."""
    charts, index_rows, missing = [], [], []
    for code, country, fname, raw_sheet, short in datasets:
        path = os.path.join(data_dir, fname)
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
                table, meta = water_year_chart.water_year_table(s["water_year"])
                ws = wb.create_sheet(sheet_name(f"{code} {s['name']} data", used))
                water_year_chart.write_table(ws, table)
                charts.append(water_year_chart.build_chart(ws, table, meta, s["title"], s["units"],
                                                           width=CHART_W, height=CHART_H))
                index_rows.append((country, f"{s['title']} (water year)", pd.Timestamp(meta["last"]).strftime("%d/%m/%y"),
                                   ws.title))
                continue
            df, n_bars = xlsx_charts.prepare(s["df"].dropna(how="all"), s.get("line_cols", ()))
            if df.empty:
                continue
            label = f"{s['name']} {short} data" if raw_sheet == "*" else f"{code} {s['name']} data"
            ws = wb.create_sheet(sheet_name(label, used))
            xlsx_charts.write_table(ws, df, s["date_format"])
            charts.append(xlsx_charts.build_chart(ws, df, n_bars, s["title"], s["units"], s["kind"],
                                                  s["date_format"], width=CHART_W, height=CHART_H))
            index_rows.append((s["name"] if raw_sheet == "*" else country, s["title"], df.index.max().strftime("%b/%y"),
                               ws.title))
        raw_sheets = ([n for n in pd.ExcelFile(path).sheet_names
                       if n.lower() not in ("units", "notes") and not n.startswith("Chart")
                       and n != water_year_chart.SHEET] if raw_sheet == "*" else [raw_sheet])
        for rs in raw_sheets:
            try:
                raw = add_charts.read(path, rs)
                label = f"{rs} {short} raw" if raw_sheet == "*" else f"{code} {short} raw"
                write_frame(wb.create_sheet(sheet_name(label, used)), raw)
            except Exception as e:
                missing.append(f"{country} {short} raw sheet {rs} ({e})")
        sources.append((country, short, fname, notes_text(path)))
    return charts, index_rows, missing


def draw_dashboard(dash, heading, charts, index_rows, missing):
    dash["B1"] = heading
    dash["B1"].font = Font(bold=True, size=18, color=DARK_BLUE)
    dash["B2"] = (f"Built {date.today():%d %b %Y} from the scheduled pulls in output/Data and Chart Outputs. "
                  "Every chart's data is on its own tab; full source data on the 'raw' tabs; units on 'Sources'.")
    dash["B2"].font = Font(italic=True, color="6B6B6B")
    for j, h in enumerate(["Country", "Chart", "Latest", "Data tab"]):
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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(DATA_DIR, "south_america_master.xlsx"))
    ap.add_argument("--data-dir", default=DATA_DIR)
    args = ap.parse_args()

    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    dash2 = wb.create_sheet("Dashboard - Power & Hydro")
    used = {"Dashboard", "Dashboard - Power & Hydro", "Sources"}
    sources = []

    gas = collect(wb, DATASETS, args.data_dir, used, sources)
    power = collect(wb, POWER_HYDRO_DATASETS, args.data_dir, used, sources)
    draw_dashboard(dash, "South America energy - gas dashboard", *gas)
    draw_dashboard(dash2, "South America energy - power generation & hydro", *power)

    src = wb.create_sheet("Sources")
    src.append(["Country", "Dataset", "Workbook", "Units and notes (from the source workbook)"])
    for cell in src[1]:
        cell.font = Font(bold=True)
    for s in sources:
        src.append(list(s))
    for col, w in (("A", 12), ("B", 14), ("C", 38), ("D", 160)):
        src.column_dimensions[col].width = w

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_charts.save_atomic(wb, args.out)
    print(f"Saved {args.out}: {len(gas[0])} gas charts, {len(power[0])} power/hydro charts; tabs {wb.sheetnames}")
    for label, m in (("gas", gas[2]), ("power/hydro", power[2])):
        if m:
            print(f"missing ({label}):", m)


if __name__ == "__main__":
    main()
