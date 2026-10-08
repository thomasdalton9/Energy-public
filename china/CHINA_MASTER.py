"""
China master workbook: every China dataset this repo pulls in one file, with Dashboard front pages carrying all their
charts - the same layout and chart code as the other masters (south_america/SOUTH_AMERICA_MASTER.py).

  Dashboard - Power      power generation by source (average GW), total generation, solar cell / generator-set output
  Dashboard - Fuels      raw coal, coke, crude oil and refinery runs, natural gas output (NBS monthly)
  Dashboard - Industry   energy-intensive output (steel, cement, glass, non-ferrous, chemicals, vehicles) and capacity
                         utilisation by industry
  Dashboard - Prices     10-day producer-goods prices (coal, coke, LNG, fuels, steel, metals, chemicals, solar and
                         battery materials, building materials) and PPI y/y by industry
  Dashboard - Long-term  annual summary and history from 2000 (fundamentals.py)
  <chart> data           the table each Dashboard chart plots
  <dataset> raw          the full data sheet(s) from each source workbook
  Sources                where each dataset comes from, units and notes, and the sources checked but not reachable

All series are China's National Bureau of Statistics (NBS) releases (asia/CHINA_NBS_*.py), the national statistics
office. NBS output covers industrial enterprises above designated size, so distributed solar and small plants are not
in the generation split, and 'Thermal' is NBS's own category (coal, gas, oil, biomass and waste). January and February
are published only combined, so the monthly series start in March 2021 and every January/February is a gap (the
combined figures are on the 'Jan-Feb' raw tabs). Ember is used only for the annual Long-term page.

Gas: only natural gas OUTPUT (NBS) is built. Gas consumption, imports (pipeline and LNG), storage and sector demand
have no official source reachable from GitHub Actions (see the Sources tab), so there is no gas balance and nothing
is estimated.

Reads (doesn't refetch) the workbooks the scheduled pulls write to "output/Data and Chart Outputs/". A missing input
is listed on its Dashboard and skipped. To add a dataset: append it to the lists below plus a SOURCES entry.

Usage: python3 china/CHINA_MASTER.py [--out "output/Data and Chart Outputs/Master Outputs/china_master.xlsx"]
"""
import argparse
import os
import sys
from types import SimpleNamespace

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
import add_charts  # noqa: E402
import fundamentals  # noqa: E402
import xlsx_charts  # noqa: E402
import SOUTH_AMERICA_MASTER as sam  # noqa: E402

DATA_DIR = sam.DATA_DIR

PROD = "china_nbs_energy_production_monthly.xlsx"
IND = "china_nbs_industrial_output_monthly.xlsx"
GAS = "china_ndrc_gas_monthly.xlsx"
PRICES = "china_nbs_market_prices_10day.xlsx"
PPI = "china_nbs_ppi_monthly.xlsx"
CAPU = "china_nbs_capacity_utilization_quarterly.xlsx"
# china_nbs_clean_energy_products_monthly.xlsx is a subset of the industrial-output workbook (same Data columns), so
# it is not read again.

# (dashboard, [(code, country, workbook, raw sheets, short name, charts to show (None = all))])
POWER = [("CN", "China", PROD, ("Data", "Jan-Feb"), "energy production",
          {"power generation by sou", "electricity generation "}),
         ("CN", "China", IND, (), "industrial output", {"solar cell and power eq"})]
FUELS = [("CN", "China", PROD, (), "energy production",
          {"raw coal output", "coke output", "crude oil output and re"})]
GAS_SETS = [("CN", "China", GAS, ("Data", "Jan-Feb"), "gas consumption", {"gas consumption y/y"}),
            ("CN", "China", PROD, (), "energy production", {"natural gas output"})]
INDUSTRY = [("CN", "China", IND, ("Data", "Jan-Feb"), "industrial output",
             {"iron and steel output", "cement output", "plate glass output", "non-ferrous metals outp",
              "chemicals output", "vehicle output"}),
            ("CN", "China", CAPU, ("Data",), "capacity utilisation", None)]
PRICE = [("CN", "China", PRICES, ("Data",), "market prices",
          {"coal and coke prices", "oil and gas product pri", "steel prices", "non-ferrous metal price",
           "basic chemical prices", "polysilicon prices", "lithium iron phosphate ", "building material price",
           "fertiliser and agrochem"}),
         ("CN", "China", PPI, ("Data",), "PPI", None)]

NBS = "National Bureau of Statistics of China (NBS)"
SOURCES = {
    GAS: ("National Development and Reform Commission (NDRC), Operation Bureau, national natural gas operation "
          "bulletin (全国天然气运行快报): apparent consumption", "https://www.ndrc.gov.cn/fggz/jjyxtj/"),
    PROD: (f"{NBS}, monthly 'industrial added value' release, table of output of major industrial products "
           "(industrial enterprises above designated size)", "https://www.stats.gov.cn/sj/zxfb/"),
    IND: (f"{NBS}, monthly 'industrial added value' release, table of output of major industrial products",
          "https://www.stats.gov.cn/sj/zxfb/"),
    PRICES: (f"{NBS}, 10-day market prices of important means of production in the circulation sector",
             "https://www.stats.gov.cn/sj/zxfb/"),
    PPI: (f"{NBS}, monthly producer price (PPI) release, by industry", "https://www.stats.gov.cn/sj/zxfb/"),
    CAPU: (f"{NBS}, quarterly industrial capacity utilisation release", "https://www.stats.gov.cn/sj/zxfb/"),
}
MASTER_SPECS = {}
HYDRO_DATASETS = []
HYDRO_EXTRA = {}
EMBER = set()
OPERATORS = {}
DASHBOARD_ONLY = {}

# Sources checked and NOT built (filled in from the probe runs, discovery_archive/china/ and results/china/).
NOT_AVAILABLE = []


def collect(wb, datasets, data_dir, used, sources):
    """sam.collect per dataset so each can show a chosen subset of its charts (names = spec names; None = all)."""
    charts, rows, missing = [], [], []
    for code, country, fname, raw, short, only in datasets:
        cfg = SimpleNamespace(**{k: getattr(sys.modules[__name__], k) for k in (
            "SOURCES", "MASTER_SPECS", "HYDRO_DATASETS", "HYDRO_EXTRA", "EMBER", "OPERATORS")},
            DASHBOARD_ONLY={fname: only} if only else {})
        before = set(wb.sheetnames)
        c, r, m = sam.collect(wb, [(code, country, fname, raw, short)], data_dir, used, sources, cfg=cfg)
        for ws in wb.worksheets:   # 'CN Data raw 2' -> 'CN industrial output Data raw': say which dataset it is
            if ws.title not in before and ws.title.endswith(" raw") or (ws.title not in before and " raw " in ws.title):
                sheet = ws.title.split(" ", 1)[1].rsplit(" raw", 1)[0]
                used.discard(ws.title)
                ws.title = sam.sheet_name(f"CN {short} {sheet} raw", used)
        charts += c
        rows += r
        missing += m
    return charts, rows, missing


def gas_combined(wb, used, data_dir):
    """Apparent consumption (NDRC) next to NBS output on one chart, over the months NDRC covers. The gap between the
    two is imports plus the output NBS does not count; it is not computed, because neither bulletin gives imports."""
    try:
        cons = add_charts.by_date(add_charts.read(os.path.join(data_dir, GAS), "Data"), "month")["Apparent_Consumption_Bcm"]
        out = add_charts.by_date(add_charts.read(os.path.join(data_dir, PROD), "Data"), "month")["Natural_Gas_Bcm"]
    except Exception as e:  # noqa: BLE001
        return None, f"gas consumption vs output ({type(e).__name__}: {e})"
    df = pd.DataFrame({"Apparent consumption (NDRC)": cons, "Output of enterprises above designated size (NBS)": out})
    df = df[df.iloc[:, 0].notna()]
    title = "China natural gas: apparent consumption (NDRC) and output (NBS)"
    df, n_bars = xlsx_charts.prepare(df, tuple(df.columns))
    ws = wb.create_sheet(sam.sheet_name("CN gas consumption vs output data", used))
    xlsx_charts.write_table(ws, df)
    src = ("NDRC national natural gas operation bulletin (consumption); NBS monthly industrial output release (output)",
           "https://www.ndrc.gov.cn/fggz/jjyxtj/")
    chart = xlsx_charts.build_chart(ws, df, n_bars, title, "bcm per month", "line", width=sam.CHART_W,
                                    height=sam.CHART_H, gridlines=False, inner=xlsx_charts.DASHBOARD_INNER)
    return (chart, src, ("China", title, df.index.max().strftime("%b/%y"), ws.title, *src)), None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(DATA_DIR, "Master Outputs", "china_master.xlsx"))
    ap.add_argument("--data-dir", default=DATA_DIR)
    args = ap.parse_args()

    wb = Workbook()
    dash_power = wb.active
    dash_power.title = "Dashboard - Power"
    names = ["Dashboard - Gas", "Dashboard - Fuels", "Dashboard - Industry", "Dashboard - Prices"]
    dashes = [wb.create_sheet(n) for n in names]
    used = {"Dashboard - Power", *names, "Dashboard - Long-term", "Sources"}
    sources = []

    sections = [(dash_power, "China - power generation (NBS, average GW)", POWER),
                (dashes[0], "China - natural gas: apparent consumption (NDRC) and output (NBS)", GAS_SETS),
                (dashes[1], "China - fuel output: coal, coke and crude oil (NBS)", FUELS),
                (dashes[2], "China - energy-intensive industrial output and capacity utilisation (NBS)", INDUSTRY),
                (dashes[3], "China - producer-goods prices and PPI (NBS)", PRICE)]
    counts = []
    for dash, heading, datasets in sections:
        charts, rows, missing = collect(wb, datasets, args.data_dir, used, sources)
        if datasets is GAS_SETS:
            both, why = gas_combined(wb, used, args.data_dir)
            if both:
                charts.insert(0, both[:2])
                rows.insert(0, both[2])
            else:
                missing.append(why)
        sam.draw_dashboard(dash, heading, charts, rows, missing)
        counts.append(len(charts))
        if missing:
            print(f"missing ({heading}):", missing)

    fundamentals.add_long_term_dashboard(wb, used, sam.sheet_name, "China", args.data_dir, sam.CHART_W, sam.CHART_H,
                                         rows_per_chart=sam.ROWS_PER_CHART, degree_days=("CDD_18", "HDD_18"),
                                         index=sum(s.startswith("Dashboard") for s in wb.sheetnames))

    src = wb.create_sheet("Sources")
    src.append(["Country", "Dataset", "Workbook", "Publisher", "Link", "Units and notes (from the source workbook)"])
    for cell in src[1]:
        cell.font = Font(bold=True)
    seen = set()
    for s in sources:
        if s[2] in seen:   # a workbook read for two dashboards is listed once
            continue
        seen.add(s[2])
        src.append(list(s))
        if s[4]:
            src.cell(row=src.max_row, column=5).hyperlink = s[4]
    if NOT_AVAILABLE:
        src.append([])
        src.append(["Checked, NOT built", "Dataset", "Publisher / page", "Result from GitHub Actions", "Link", "Notes"])
        for cell in src[src.max_row]:
            cell.font = Font(bold=True)
        for row in NOT_AVAILABLE:
            src.append(list(row))
            if row[4]:
                src.cell(row=src.max_row, column=5).hyperlink = row[4]
    for col, w in (("A", 18), ("B", 22), ("C", 44), ("D", 60), ("E", 50), ("F", 160)):
        src.column_dimensions[col].width = w

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_charts.save_atomic(wb, args.out)
    print(f"Saved {args.out}: charts per dashboard {dict(zip(['power', *names], counts))}; tabs {wb.sheetnames}")


if __name__ == "__main__":
    main()
