"""
Crude and products master workbook: US crude and refined-product stocks and crude oil prices in one file, with Dashboard
front pages of every chart - the same layout and chart code as the other masters (south_america/SOUTH_AMERICA_MASTER.py).

  Dashboard - Stocks   US weekly stocks (EIA): Big Four (crude incl. SPR + gasoline + distillate + jet fuel), crude incl.
                       SPR, gasoline, distillate, jet fuel, commercial crude and SPR, each as a calendar-year-weeks chart:
                       shaded 5-year min-max band, 5-year average, last year and current year
  Dashboard - Prices   Brent and WTI spot (daily), Brent monthly average, Brent-WTI spread (EIA)
  <chart> data         the table each Dashboard chart plots
  <dataset> raw        the full data sheets from each source workbook
  Sources              where each dataset comes from

Reads (doesn't refetch) the workbooks the scheduled pulls write to "output/Data and Chart Outputs/". A missing input is
listed on the Dashboard and skipped. To add a dataset: append it to STOCK_DATASETS / PRICE_DATASETS and add a SOURCES entry.

Usage: python3 americas/CRUDE_AND_PRODUCTS_MASTER.py [--out "output/Data and Chart Outputs/Master Outputs/crude_and_products_master.xlsx"]
"""
import argparse
import os
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
import add_charts  # noqa: E402
import xlsx_charts  # noqa: E402
import SOUTH_AMERICA_MASTER as sam  # noqa: E402

DATA_DIR = sam.DATA_DIR
# Owner asked for the plot inside each chart to be 1.5 cm longer at the bottom (0.5 cm, then 1 cm more), the chart box itself staying the same size
# (CHART_H cm): the plot-area height fraction grows by 0.5 / CHART_H. Set on the shared module for this process only.
_x, _y, _w, _h = xlsx_charts.DASHBOARD_INNER
xlsx_charts.DASHBOARD_INNER = (_x, _y, _w, _h + 1.5 / sam.CHART_H)

# (code, name, workbook, raw sheet / tuple of raw sheets / "*", short dataset name)
STOCK_DATASETS = [
    ("US", "United States", "us_petroleum_stocks_weekly.xlsx", "Data", "petroleum stocks"),
]
PRICE_DATASETS = [
    ("GL", "Global", "brent_wti_daily.xlsx", ("Data", "Monthly"), "Brent and WTI"),
]
HYDRO_DATASETS = []
HYDRO_EXTRA = {}
DASHBOARD_ONLY = {}
EMBER = set()
OPERATORS = {}

SOURCES = {
    "us_petroleum_stocks_weekly.xlsx": ("EIA, Weekly Petroleum Status Report (stocks of crude oil, SPR, motor gasoline, distillate, "
                                        "jet fuel)", "https://www.eia.gov/petroleum/supply/weekly/"),
    "brent_wti_daily.xlsx": ("EIA, spot prices (Brent: Europe Brent Spot Price FOB; WTI: Cushing)",
                             "https://www.eia.gov/dnav/pet/pet_pri_spt_s1_d.htm"),
}

STOCK_SERIES = [  # (sheet column, chart name, title)
    ("Big Four", "Big Four", "US Big Four stocks (crude incl. SPR + gasoline + distillate + jet fuel)"),
    ("Crude incl SPR", "Crude incl SPR", "US crude oil stocks incl. SPR"),
    ("Gasoline", "Gasoline", "US motor gasoline stocks"),
    ("Distillate", "Distillate", "US distillate fuel oil stocks"),
    ("Jet fuel", "Jet fuel", "US jet fuel stocks"),
    ("Crude excl SPR", "Crude excl SPR", "US commercial crude oil stocks (excl. SPR)"),
    ("SPR crude", "SPR crude", "US Strategic Petroleum Reserve crude"),
]


def stock_specs(path):
    """Calendar-year-weeks band charts (5-year range, 5-year average, last year, current year) from the Data sheet."""
    d = add_charts.by_date(add_charts.read(path, "Data"), "date")
    return [{"name": name, "week_year": pd.to_numeric(d[col], errors="coerce"), "title": title,
             "units": "thousand barrels", "y_decimals": 0} for col, name, title in STOCK_SERIES if col in d.columns]


MASTER_SPECS = {"us_petroleum_stocks_weekly.xlsx": stock_specs}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(DATA_DIR, "Master Outputs", "crude_and_products_master.xlsx"))
    ap.add_argument("--data-dir", default=DATA_DIR)
    args = ap.parse_args()
    cfg = sys.modules[__name__]

    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard - Stocks"
    dash2 = wb.create_sheet("Dashboard - Prices")
    used = {"Dashboard - Stocks", "Dashboard - Prices", "Sources"}
    sources = []
    stocks = sam.collect(wb, STOCK_DATASETS, args.data_dir, used, sources, cfg=cfg)
    prices = sam.collect(wb, PRICE_DATASETS, args.data_dir, used, sources, cfg=cfg)
    sam.draw_dashboard(dash, "Crude and products - US weekly stocks (calendar-year weeks, 5-year range)", *stocks)
    sam.draw_dashboard(dash2, "Crude and products - crude oil prices", *prices)

    src = wb.create_sheet("Sources")
    src.append(["Country", "Dataset", "Workbook", "Publisher", "Link", "Units and notes (from the source workbook)"])
    for cell in src[1]:
        cell.font = Font(bold=True)
    for s in sources:
        src.append(list(s))
        if s[4]:
            src.cell(row=src.max_row, column=5).hyperlink = s[4]
    for col, w in (("A", 14), ("B", 18), ("C", 38), ("D", 60), ("E", 50), ("F", 160)):
        src.column_dimensions[col].width = w

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_charts.save_atomic(wb, args.out)
    print(f"Saved {args.out}: {len(stocks[0])} stock charts, {len(prices[0])} price charts; tabs {wb.sheetnames}")
    for label, m in (("stocks", stocks[2]), ("prices", prices[2])):
        if m:
            print(f"missing ({label}):", m)


if __name__ == "__main__":
    main()
