"""
Europe master workbook: one file with every European dataset this repo pulls straight from national TSOs, and
Dashboard front pages carrying all their charts - the same layout as the North America and Australia/NZ masters
(south_america/SOUTH_AMERICA_MASTER.py, whose table/chart/dashboard code this reuses).

  Dashboard          - gas: Denmark gas flows (Energinet)
  Dashboard - Power  - Europe generation by source (sum of the countries below), then each country's generation
                       by source (GB NESO, Belgium Elia, Denmark Energinet, Spain REE)
  <CC> <chart> data  - the table each Dashboard chart plots
  <CC> <dataset> raw - the full data sheet(s) from each source workbook
  Sources            - where each dataset comes from, units and notes

Sources are the national transmission system operators' own open data (no ENTSO-E / ENTSOG). A country is added by
writing its pull (europe/<country>_generation_<tso>.py -> <country>_power_generation_daily.xlsx), then adding it to
RAW_POWER_DATASETS, SOURCES and add_charts.REGISTRY.

Reads (doesn't refetch) the workbooks the scheduled pulls write to "output/Data and Chart Outputs/". A missing
input is listed on the Dashboard and skipped rather than stopping the rest.

Usage: python3 EUROPE_MASTER.py [--out "output/Data and Chart Outputs/europe_master.xlsx"]
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
FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Other"]   # same order/colours as the other dashboards

# (country code, country, workbook, raw sheet / tuple of raw sheets / "*", short dataset name)
DATASETS = [
    ("DK", "Denmark", "denmark_gas_flows.xlsx", "Daily", "gas"),
]
RAW_POWER_DATASETS = [
    ("GB", "Great Britain", "great_britain_power_generation_daily.xlsx", "Daily", "power"),
    ("BE", "Belgium", "belgium_power_generation_daily.xlsx", "Daily", "power"),
    ("DK", "Denmark", "denmark_power_generation_daily.xlsx", "Daily", "power"),
    ("ES", "Spain", "spain_power_generation_daily.xlsx", "Daily", "power"),
]
CAPACITY_DATASETS = []
HYDRO_DATASETS = []
HYDRO_EXTRA = {}
DASHBOARD_ONLY = {}
MASTER_SPECS = {}
EMBER = set()
OPERATORS = {}

SOURCES = {
    "denmark_gas_flows.xlsx": ("Energinet Gas TSO, Energi Data Service 'Gasflow' (commercial gas amounts)",
                               "https://www.energidataservice.dk/tso-gas/Gasflow"),
    "great_britain_power_generation_daily.xlsx": ("NESO (National Energy System Operator), Historic Generation Mix",
                                                  "https://www.neso.energy/data-portal/historic-generation-mix"),
    "belgium_power_generation_daily.xlsx": ("Elia (Belgian TSO) open data, ods201 total generation by fuel type",
                                            "https://opendata.elia.be/explore/dataset/ods201/"),
    "denmark_power_generation_daily.xlsx": ("Energinet (Danish TSO), Energi Data Service GenerationProdTypeExchange",
                                            "https://www.energidataservice.dk/tso-electricity/GenerationProdTypeExchange"),
    "spain_power_generation_daily.xlsx": ("Red Eléctrica de España (REE), REData generation structure",
                                          "https://www.ree.es/en/datos/generation/generation-structure"),
}


def europe_generation(data_dir, have):
    """Europe generation by source, TWh per month: the sum of each country's monthly mix from the series the
    dashboard shows. Only months every country has in full (a daily feed's current month is left out until it is
    complete). Returns (frame, per-country notes)."""
    frames, notes = {}, []
    for code, country, fname, sheet, _ in RAW_POWER_DATASETS:
        if country not in have:
            continue
        try:
            path = os.path.join(data_dir, fname)
            m = add_charts.REGISTRY[fname](path)[0]["df"].apply(pd.to_numeric, errors="coerce")
            last = add_charts.by_date(add_charts.read(path, sheet), "date").index.max()
            if last < last + pd.offsets.MonthEnd(0):   # current month not complete yet
                m = m[m.index < last.to_period("M").to_timestamp()]
            frames[country] = m
            notes.append(f"{country}: {SOURCES[fname][0]}, {m.index.min():%b/%y}-{m.index.max():%b/%y}")
        except Exception as e:  # noqa: BLE001
            notes.append(f"{country}: not available ({type(e).__name__}: {e})")
    if not frames:
        return pd.DataFrame(), notes
    # a country's early gap (its feed starts later) is reported, not estimated: run from the latest first month
    months = sorted(set.intersection(*[set(f.dropna(how="all").index) for f in frames.values()]))
    total = sum(f.reindex(index=months, columns=FUELS).fillna(0) for f in frames.values()) / 1000.0   # GWh -> TWh
    total.index.name = "date"
    return total, notes


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(DATA_DIR, "europe_master.xlsx"))
    ap.add_argument("--data-dir", default=DATA_DIR)
    args = ap.parse_args()
    cfg = sys.modules[__name__]

    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    dash2 = wb.create_sheet("Dashboard - Power")
    used = {"Dashboard", "Dashboard - Power", "Sources"}
    sources = []

    gas = sam.collect(wb, DATASETS, args.data_dir, used, sources, cfg=cfg)
    raw_power = [d for d in RAW_POWER_DATASETS if os.path.exists(os.path.join(args.data_dir, d[2]))]
    have = {d[1] for d in raw_power}
    power = sam.collect(wb, raw_power, args.data_dir, used, sources, cfg=cfg)
    power[2].extend(f"{d[1]} power ({d[2]}) not built yet" for d in RAW_POWER_DATASETS if d[1] not in have)

    total, notes = europe_generation(args.data_dir, have)
    if not total.empty:
        ws = wb.create_sheet(sam.sheet_name("Europe generation total data", used))
        df, n_bars = xlsx_charts.prepare(total)
        xlsx_charts.write_table(ws, df)
        ws.cell(row=1, column=df.shape[1] + 4, value="Countries summed (only months all of them have):")
        for i, note in enumerate(notes, start=2):
            ws.cell(row=i, column=df.shape[1] + 4, value=note)
        src = ("Sum of the country series on this dashboard (national TSOs)", None)
        power[0].insert(0, (xlsx_charts.build_chart(ws, df, n_bars, "Europe power generation by source "
                                                    f"({', '.join(sorted(have))})", "TWh per month", "stacked_bar",
                                                    width=sam.CHART_W, height=sam.CHART_H, gridlines=False,
                                                    inner=xlsx_charts.DASHBOARD_INNER), src))
        power[1].insert(0, ("Europe", f"Europe power generation by source ({len(have)} countries covered so far)",
                            df.index.max().strftime("%b/%y"), ws.title, *src))
        print("Europe generation total:", "; ".join(notes))

    sam.draw_dashboard(dash, "Europe energy (national TSOs) - gas dashboard", *gas)
    sam.draw_dashboard(dash2, "Europe energy (national TSOs) - power generation", *power)

    src = wb.create_sheet("Sources")
    src.append(["Country", "Dataset", "Workbook", "Publisher", "Link", "Units and notes (from the source workbook)"])
    for cell in src[1]:
        cell.font = Font(bold=True)
    for s in sources:
        src.append(list(s))
        if s[4]:
            src.cell(row=src.max_row, column=5).hyperlink = s[4]
    for col, w in (("A", 14), ("B", 14), ("C", 38), ("D", 40), ("E", 50), ("F", 160)):
        src.column_dimensions[col].width = w

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_charts.save_atomic(wb, args.out)
    print(f"Saved {args.out}: {len(gas[0])} gas charts, {len(power[0])} power charts; tabs {wb.sheetnames}")
    for label, m in (("gas", gas[2]), ("power", power[2])):
        if m:
            print(f"missing ({label}):", m)


if __name__ == "__main__":
    main()
