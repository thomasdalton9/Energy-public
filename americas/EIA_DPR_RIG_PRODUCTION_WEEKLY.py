"""
Pull EIA's Drilling Productivity Report (DPR) and maintain a growing
archive of rig counts and oil/gas production by basin - one tab per
basin, plus a DPR-covered-total tab, each with a native chart.

Data source: https://www.eia.gov/petroleum/drilling/xls/dpr-data.xlsx -
a direct public .xlsx download, no API key needed (unlike EIA-930/STEO)
and, unlike Baker Hughes' own rig count site, not gated behind Akamai
(Baker Hughes blocks cloud/data-centre IPs - see the "RUN IT ON YOUR
PC" warning in Baker_Hughes_Rig_Count.py; this DPR file is a plain
download and works fine from GitHub Actions). Found via
EIA_DPR_DISCOVERY.py, including the exact sheet/column layout below
(confirmed against the real file, not guessed).

SCOPE: DPR covers 7 major shale regions (not the whole US): Anadarko,
Appalachia, Bakken, Eagle Ford, Haynesville, Niobrara, Permian. Each
region's own sheet reports, monthly, back to 2007:
  Rig count                     rigs operating in the region
  Oil: Production per rig       new-well oil output per rig, bbl/d
  Oil: Legacy production change bbl/d change from existing (non-new) wells
  Oil: Total production         bbl/d, the region's actual oil output
  Gas: Production per rig       new-well gas output per rig, Mcf/d
  Gas: Legacy production change Mcf/d change from existing wells
  Gas: Total production         Mcf/d, the region's actual gas output
"Legacy production change" is always negative or near-zero - it is the
natural decline of wells drilled in prior months, which new-well
production (via rig count x production-per-rig) has to outrun for
total production to grow.

This complements, not replaces, the other rig/production sources in
this repo: Baker_Hughes_Rig_Count.py has a finer weekly cadence but
must run locally; US_Gas_Production.py pulls STEO's own Lower-48 and
three-basin (Appalachia/Haynesville/Permian) dry gas forecast into the
Windows dashboard workbook. This script is EIA-native, basin-complete
(7 regions, not 3), includes rig count AND oil production (STEO's
script doesn't), and runs unattended on GitHub Actions.

CADENCE: DPR itself is published monthly (roughly mid-month) - the
workflow runs weekly regardless (a run with no new month published is
a harmless no-op, same pattern as MISO's script safely re-running
mid-day).

No history is discarded on later runs missing a month: this always
re-downloads and re-parses the whole file (it's small, ~150KB) and
upserts by month, so a source correction to an older month is picked
up too, not just new months appended.
"""

import argparse
import os
import sys
import time
from datetime import date

import openpyxl
import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

URL = "https://www.eia.gov/petroleum/drilling/xls/dpr-data.xlsx"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 60)
FETCH_ATTEMPTS = 5
RETRY_BACKOFF_SECONDS = [5, 15, 30, 60]

# source sheet name -> our tab name
REGIONS = {
    "Anadarko Region": "Anadarko",
    "Appalachia Region": "Appalachia",
    "Bakken Region": "Bakken",
    "Eagle Ford Region": "EagleFord",
    "Haynesville Region": "Haynesville",
    "Niobrara Region": "Niobrara",
    "Permian Region": "Permian",
}

COLUMNS = [
    "Rig_Count",
    "Oil_Production_per_Rig_bbl_d",
    "Oil_Legacy_Change_bbl_d",
    "Oil_Total_Production_bbl_d",
    "Gas_Production_per_Rig_Mcf_d",
    "Gas_Legacy_Change_Mcf_d",
    "Gas_Total_Production_Mcf_d",
]

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "eia_dpr_rig_and_production_weekly.xlsx")

# DPR publishes monthly, with roughly a 2-week lag after month end - a
# comfortable multiple of that before this is treated as genuinely
# stale (the workflow itself only runs weekly, so this also has to
# tolerate the gap between runs).
STALE_AFTER_DAYS = 45


def fetch_workbook_bytes():
    last_error = None
    for attempt in range(1, FETCH_ATTEMPTS + 1):
        try:
            r = requests.get(URL, headers=HEADERS, timeout=TIMEOUT)
            r.raise_for_status()
            return r.content
        except requests.RequestException as e:
            last_error = e
            print(f"  attempt {attempt}/{FETCH_ATTEMPTS} failed: {type(e).__name__}: {e}", file=sys.stderr)
            if attempt < FETCH_ATTEMPTS:
                wait = RETRY_BACKOFF_SECONDS[attempt - 1]
                print(f"  waiting {wait}s before retrying...", file=sys.stderr)
                time.sleep(wait)
    raise last_error


def parse_region_sheet(wb, source_sheet_name):
    """Row 1 is a merged oil/gas group header, row 2 the real column
    header, data from row 3 - see the module docstring for the exact
    layout as confirmed by EIA_DPR_DISCOVERY.py."""
    ws = wb[source_sheet_name]
    records = []
    for row in ws.iter_rows(min_row=3, max_col=8, values_only=True):
        month = row[0]
        if month is None:
            continue
        records.append({
            "date": pd.Timestamp(month).date(),
            "Rig_Count": row[1],
            "Oil_Production_per_Rig_bbl_d": row[2],
            "Oil_Legacy_Change_bbl_d": row[3],
            "Oil_Total_Production_bbl_d": row[4],
            "Gas_Production_per_Rig_Mcf_d": row[5],
            "Gas_Legacy_Change_Mcf_d": row[6],
            "Gas_Total_Production_Mcf_d": row[7],
        })
    df = pd.DataFrame.from_records(records).set_index("date")
    df.index.name = "date"
    return df.sort_index()


def build_total_sheet(region_frames):
    """DPR-covered total: sums what's meaningfully additive (rig
    count, legacy change, total production) across all 7 regions.
    Production-per-rig is a per-rig average, not summable across
    regions with very different rig productivity, so it's left out of
    the total rather than being a misleading sum."""
    summable = [
        "Rig_Count", "Oil_Legacy_Change_bbl_d", "Oil_Total_Production_bbl_d",
        "Gas_Legacy_Change_Mcf_d", "Gas_Total_Production_Mcf_d",
    ]
    total = None
    for df in region_frames.values():
        piece = df[summable]
        total = piece if total is None else total.add(piece, fill_value=0)
    return total.sort_index()


def format_date_columns(path, sheet_names):
    wb = openpyxl.load_workbook(path)
    for sheet in sheet_names:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for (cell,) in ws.iter_rows(min_row=2, max_col=1):
            cell.number_format = "mmm-yyyy"
        ws.column_dimensions["A"].width = 12
    wb.save(path)


CHART_FONT_SIZE = 1200  # hundredths of a point - 1200 = 12pt


def _sized_title(text):
    from openpyxl.chart.text import RichText, Text
    from openpyxl.chart.title import Title
    from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RegularTextRun

    cp = CharacterProperties(sz=CHART_FONT_SIZE, b=True)
    run = RegularTextRun(t=text, rPr=cp)
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[run])
    return Title(tx=Text(rich=RichText(p=[para])))


def _sized_text_props():
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties

    cp = CharacterProperties(sz=CHART_FONT_SIZE)
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp))])


def add_region_chart(ws, sheet_name, n_data_rows):
    """Rig count (line, primary axis) vs total gas production (line,
    secondary axis) - two genuinely different units, so a secondary
    axis is the honest way to show both without one dwarfing the
    other. Native/editable in Excel, not a static image - openpyxl
    chart objects, no win32com/Excel install needed (this runs on
    Linux GitHub Actions runners)."""
    from openpyxl.chart import LineChart, Reference
    from openpyxl.utils import column_index_from_string

    if n_data_rows == 0:
        return

    last_row = n_data_rows + 1  # +1 for the header row
    rig_col = column_index_from_string("B")  # Rig_Count
    gas_total_col = column_index_from_string("H")  # Gas_Total_Production_Mcf_d

    primary = LineChart()
    primary.title = _sized_title(f"{sheet_name}: rig count vs total gas production")
    primary.y_axis.title = "Rig count"
    primary.x_axis.title = "Month"
    primary.x_axis.txPr = _sized_text_props()
    primary.y_axis.txPr = _sized_text_props()
    primary.height = 11
    primary.width = 26
    rig_data = Reference(ws, min_col=rig_col, max_col=rig_col, min_row=1, max_row=last_row)
    cats = Reference(ws, min_col=1, min_row=2, max_row=last_row)
    primary.add_data(rig_data, titles_from_data=True)
    primary.set_categories(cats)

    secondary = LineChart()
    gas_data = Reference(ws, min_col=gas_total_col, max_col=gas_total_col, min_row=1, max_row=last_row)
    secondary.add_data(gas_data, titles_from_data=True)
    secondary.y_axis.axId = 200
    secondary.y_axis.title = "Gas total production (Mcf/d)"
    secondary.y_axis.txPr = _sized_text_props()
    secondary.y_axis.crosses = "max"

    primary += secondary
    if primary.legend is not None:
        primary.legend.txPr = _sized_text_props()

    ws.add_chart(primary, "K2")


def add_charts(path, sheet_row_counts):
    wb = openpyxl.load_workbook(path)
    for sheet_name, n_rows in sheet_row_counts.items():
        if sheet_name not in wb.sheetnames:
            continue
        add_region_chart(wb[sheet_name], sheet_name, n_rows)
    wb.save(path)


NOTES_LINES = [
    "UNITS",
    "Rig_Count: number of rigs operating in the region that month. Oil_*_bbl_d / Gas_*_Mcf_d: "
    "barrels per day / thousand cubic feet per day. *_Production_per_Rig: new-well output per rig - a "
    "drilling/completion productivity measure, NOT total output. *_Legacy_Change: the monthly change in "
    "production from wells drilled in prior months (natural decline - almost always negative). "
    "*_Total_Production: the region's actual total output (new-well + legacy combined).",
    "",
    "TABS / SCOPE",
    "One tab per DPR basin: Anadarko, Appalachia, Bakken, EagleFord, Haynesville, Niobrara, Permian - "
    "these 7 shale regions are DPR's full coverage, NOT the whole US (see US_Gas_Production.py / "
    "MISO_FUEL_MIX_DAILY.py / eia930_fuel_mix_daily.xlsx for broader US supply/generation). DPR_Total "
    "sums what's meaningfully additive (Rig_Count, Legacy_Change, Total_Production) across all 7 - "
    "*_Production_per_Rig is a per-rig average and isn't summed into a region total.",
    "",
    "WHY THIS SOURCE",
    "Baker Hughes' own weekly rig count (baker_hughes_rig_count.xlsx) is finer-grained but its site "
    "blocks cloud/data-centre IPs and must run locally (see that script's docstring). This DPR file is "
    "a plain public download with no such gate, so it runs unattended here - at DPR's own monthly "
    "cadence, not weekly, even though the workflow itself checks weekly (a run with nothing new "
    "published is a harmless no-op).",
    "",
    "COVERAGE",
    "Re-downloads and re-parses the entire DPR file on every run (~150KB, back to 2007) and upserts by "
    "month, so a source revision to an older month is picked up too, not just newly-appended months.",
    "",
    "SOURCE",
    f"EIA Drilling Productivity Report: {URL} - a direct public .xlsx download, no API key needed.",
]
NOTES_SECTION_TITLES = {"UNITS", "TABS / SCOPE", "WHY THIS SOURCE", "COVERAGE", "SOURCE"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    print(f"Fetching {URL} ...", file=sys.stderr)
    raw = fetch_workbook_bytes()
    src_wb = openpyxl.load_workbook(__import__("io").BytesIO(raw), data_only=True)

    sheets = {}
    row_counts = {}
    for source_name, tab_name in REGIONS.items():
        df = parse_region_sheet(src_wb, source_name)
        sheets[tab_name] = df
        row_counts[tab_name] = len(df)
        print(f"  {tab_name}: {len(df)} months ({df.index.min()} to {df.index.max()})", file=sys.stderr)

    total_df = build_total_sheet(sheets)
    sheets["DPR_Total"] = total_df
    row_counts["DPR_Total"] = len(total_df)

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, NOTES_LINES, NOTES_SECTION_TITLES)
    format_date_columns(args.out, list(sheets.keys()))
    add_charts(args.out, row_counts)

    latest = total_df.index.max() if not total_df.empty else None
    print(f"Saved to {args.out}. DPR_Total covers {len(total_df)} months, latest {latest}.")

    age = (date.today() - latest).days if latest else None
    if latest is None or age > STALE_AFTER_DAYS:
        print(f"STALE SOURCE: newest saved month is {latest} ({age} days old) - "
              f"{URL} may have changed or stopped updating.", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
