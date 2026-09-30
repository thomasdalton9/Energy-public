"""
Pull EIA-930 (Hourly Electric Grid Monitor) daily net generation by fuel
type for the whole US grid AND for a set of major individual ISOs/RTOs,
and maintain a growing archive - one tab per region in a single
workbook.

Data source: EIA API v2, electricity/rto/daily-fuel-type-data -
https://api.eia.gov/v2/electricity/rto/daily-fuel-type-data/data/ -
"Daily net generation by balancing authority and energy source" per
Form EIA-930. Needs a free EIA API key (unlike MISO/FRED) - read from
the EIA_API_KEY environment variable (a GitHub Actions secret in this
repo, never printed or logged by this script). Found via
EIA930_DISCOVERY.py, including the exact respondent codes below
(confirmed against the API's own respondent facet list, not guessed).

REGIONS: US48 (EIA's own pre-aggregated whole-Lower-48 total) plus PJM,
MISO, CAISO, ERCOT, SPP, NYISO, ISO-NE, Southern Co., and TVA - each is
its own respondent code in EIA-930, reported the same way as US48 (not
summed here from smaller pieces). MISO here is a second, independent,
coarser (daily, not 5-minute) cross-check on MISO_FUEL_MIX_DAILY.py's
own real-time pull - both are kept, not merged.

Unlike MISO's own API (only "today"/"yesterday", no range query), this
dataset supports a real start/end range query with pagination (5000
rows/page) - so, like HENRY_HUB_DAILY.py, one script both backfills the
full history and keeps it current per region: the first run (no archive
yet, or a region's tab missing from it) pages through that region's
entire history back to the dataset's own start (2019-01-01); every run
after that just re-pulls a short rolling window per region (see
ROLLING_WINDOW_DAYS) and upserts, which also naturally catches any
revisions EIA makes to recently-published days.

TIMEZONE FACET: EIA-930 reports each region's total five times over,
once per US timezone (Arizona/Central/Eastern/Mountain/Pacific) - the
"timezone" facet is which timezone's midnight is used to bucket hourly
values into a calendar day, not a regional subset of the data. This
script fixes timezone=Eastern throughout for one consistent daily
series per region - an arbitrary but documented choice, not "the"
official cutoff (and not necessarily each ISO's own local timezone).

Values are net generation in megawatthours (MWh) - a daily total, not
an average MW like MISO's own script (MISO's 5-minute API instead means
a mean MW; there's no equivalent 5-minute read here, only EIA's own
daily aggregate).

Fuel-type categories are read from whatever the API actually reports
each pull (via each row's own "type-name"), not hardcoded, so a new
category EIA adds shows up automatically as a new column - and not
every region reports every category (e.g. ISO-NE has no Coal some
years), which shows up as that column simply being absent or NaN for
that region's tab.
"""

import argparse
import os
import sys
import time
from datetime import date, timedelta

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

API_KEY = os.environ.get("EIA_API_KEY")
URL = "https://api.eia.gov/v2/electricity/rto/daily-fuel-type-data/data/"
TIMEOUT = (10, 30)

# respondent code -> sheet name. Order here is the order tabs appear in
# the workbook. Codes confirmed via EIA930_DISCOVERY.py's respondent
# facet listing.
REGIONS = {
    "US48": "US_Total",
    "PJM": "PJM",
    "MISO": "MISO",
    "CISO": "CAISO",
    "ERCO": "ERCOT",
    "SWPP": "SPP",
    "NYIS": "NYISO",
    "ISNE": "ISONE",
    "SOCO": "Southern",
    "TVA": "TVA",
}

TIMEZONE = "Eastern"

# The dataset's own startPeriod as of EIA930_DISCOVERY.py's probe.
HISTORY_START = date(2019, 1, 1)

# Every run re-pulls this many trailing days per region (a single
# request each, well under the 5000-row page limit) and upserts them -
# catches EIA revising recently-published days, at negligible cost.
ROLLING_WINDOW_DAYS = 14

PAGE_LENGTH = 5000
FETCH_ATTEMPTS = 5
RETRY_BACKOFF_SECONDS = [5, 15, 30, 60]

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "eia930_fuel_mix_daily.xlsx")

# EIA-930 publishes daily, with a lag of a few days for the "final"
# vintage - a few extra days of slack absorb that before this is
# treated as a genuinely stale source. Checked against US48 only (the
# headline series); a single smaller region lagging behind doesn't by
# itself fail the whole run.
STALE_AFTER_DAYS = 10


def fetch_page(respondent, start, end, offset):
    params = {
        "frequency": "daily",
        "data[0]": "value",
        "facets[respondent][]": respondent,
        "facets[timezone][]": TIMEZONE,
        "start": start.isoformat(),
        "end": end.isoformat(),
        "sort[0][column]": "period",
        "sort[0][direction]": "asc",
        "offset": offset,
        "length": PAGE_LENGTH,
        "api_key": API_KEY,
    }
    last_error = None
    for attempt in range(1, FETCH_ATTEMPTS + 1):
        try:
            r = requests.get(URL, params=params, timeout=TIMEOUT)
            r.raise_for_status()
            return r.json()
        except requests.RequestException as e:
            last_error = e
            print(f"    attempt {attempt}/{FETCH_ATTEMPTS} failed: {type(e).__name__}: {e}", file=sys.stderr)
            if attempt < FETCH_ATTEMPTS:
                wait = RETRY_BACKOFF_SECONDS[attempt - 1]
                print(f"    waiting {wait}s before retrying...", file=sys.stderr)
                time.sleep(wait)
    raise last_error


def fetch_range(respondent, start, end):
    """Page through offset/length until every row in [start, end] is
    collected for one respondent."""
    rows = []
    offset = 0
    while True:
        payload = fetch_page(respondent, start, end, offset)
        page_rows = payload["response"]["data"]
        rows.extend(page_rows)
        total = int(payload["response"]["total"])
        offset += len(page_rows)
        print(f"    fetched {offset}/{total} rows...", file=sys.stderr)
        if not page_rows or offset >= total:
            break
    return rows


def sanitize_column(type_name):
    return type_name.strip().replace(" ", "_").replace("-", "_")


def rows_to_wide(rows):
    """Long rows (period, fueltype, type-name, value) -> wide
    DataFrame, index = date, columns = "{type-name}_MWh"."""
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    df["period"] = pd.to_datetime(df["period"]).dt.date
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    df["column"] = df["type-name"].map(sanitize_column) + "_MWh"
    wide = df.pivot_table(index="period", columns="column", values="value", aggfunc="last")
    wide.index.name = "date"
    return wide.sort_index()


def add_summary_columns(df):
    fuel_cols = [c for c in df.columns if c.endswith("_MWh")]
    df["Total_MWh"] = df[fuel_cols].sum(axis=1, skipna=True)
    # Renewables share computed from known type-names actually present, matched case-
    # insensitively so it works regardless of EIA's exact capitalization, and regardless
    # of which categories a given region happens to report.
    renewable_keywords = ["wind", "solar", "hydro", "geothermal", "biomass"]
    renewable_cols = [c for c in fuel_cols if any(kw in c.lower() for kw in renewable_keywords)]
    renewable_sum = df[renewable_cols].sum(axis=1, skipna=True) if renewable_cols else 0.0
    df["Renewables_Share"] = renewable_sum / df["Total_MWh"].where(df["Total_MWh"] != 0)
    return df


def load_archive(path, sheet_name):
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, index_col=0)
    except (FileNotFoundError, ValueError, KeyError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index).date
    df.index.name = "date"
    return df


def upsert(existing, new_df):
    if new_df.empty:
        return existing
    if existing.empty:
        combined = new_df
    else:
        combined = pd.concat([existing, new_df])
        combined = combined[~combined.index.duplicated(keep="last")]
    combined.index.name = "date"
    return combined.sort_index()


def format_date_columns(path, sheet_names):
    import openpyxl

    wb = openpyxl.load_workbook(path)
    for sheet in sheet_names:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for (cell,) in ws.iter_rows(min_row=2, max_col=1):
            cell.number_format = "dd-mmm-yyyy"
        ws.column_dimensions["A"].width = 14
    wb.save(path)


CHART_FONT_SIZE = 1200  # openpyxl font sizes are in hundredths of a point - 1200 = 12pt


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


def add_region_chart(ws, sheet_name, n_data_rows, fuel_cols):
    """Native (embedded, editable-in-Excel) stacked area chart of a
    region's daily generation by fuel type - not a static image, and
    not win32com-automated (this pipeline runs on Linux GitHub Actions
    runners, no Excel installed), just openpyxl's chart objects, which
    Excel opens and renders like any chart built by hand."""
    from openpyxl.chart import AreaChart, Reference

    if n_data_rows == 0 or not fuel_cols:
        return

    chart = AreaChart()
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = _sized_title(f"{sheet_name} daily generation by fuel type")
    chart.y_axis.title = "MWh"
    chart.x_axis.title = "Date"
    chart.x_axis.txPr = _sized_text_props()
    chart.y_axis.txPr = _sized_text_props()
    chart.height = 11
    chart.width = 26

    last_row = n_data_rows + 1  # +1 for the header row
    first_fuel_col = 2  # column A is the date index; data starts at column B
    last_fuel_col = first_fuel_col + len(fuel_cols) - 1

    data = Reference(ws, min_col=first_fuel_col, max_col=last_fuel_col, min_row=1, max_row=last_row)
    cats = Reference(ws, min_col=1, min_row=2, max_row=last_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    if chart.legend is not None:
        chart.legend.txPr = _sized_text_props()

    anchor_col_index = last_fuel_col + 3  # a couple of columns clear of Total_MWh/Renewables_Share
    from openpyxl.utils import get_column_letter

    ws.add_chart(chart, f"{get_column_letter(anchor_col_index)}2")


def add_charts(path, region_fuel_cols):
    """region_fuel_cols: {sheet_name: [fuel column names]}, in the same
    column order they were written to that sheet."""
    import openpyxl

    wb = openpyxl.load_workbook(path)
    for sheet_name, fuel_cols in region_fuel_cols.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        n_data_rows = ws.max_row - 1  # minus the header row
        add_region_chart(ws, sheet_name, n_data_rows, fuel_cols)
    wb.save(path)


NOTES_LINES = [
    "UNITS",
    "All *_MWh columns are megawatthours - the total net generation for that fuel category over the "
    "whole day (a daily energy total, not an average MW). Total_MWh is the sum of the category "
    "columns. Renewables_Share is (Wind + Solar + Hydro + Geothermal + Biomass) as a share of "
    "Total_MWh.",
    "",
    "TABS / SCOPE",
    "One tab per region, each EIA-930's own pre-aggregated total for that respondent code - not "
    "summed here from smaller pieces: US_Total (US48, the whole Lower 48), PJM, MISO, CAISO (CISO), "
    "ERCOT (ERCO), SPP (SWPP), NYISO (NYIS), ISONE (ISNE), Southern (SOCO), TVA. The MISO tab here is "
    "a second, independent, coarser (daily, not 5-minute) cross-check on miso_fuel_mix_daily.xlsx's "
    "real-time pull - both are kept, not merged.",
    "",
    "CATEGORIES",
    "Whatever EIA-930 itself reports per region (e.g. Coal, Natural gas, Nuclear, Wind, Solar, Hydro, "
    "Petroleum, Battery storage, Other) - not a fixed list maintained here, so a category EIA adds "
    "shows up as a new column automatically, and a region that doesn't report a given category simply "
    "has no column (or NaN rows) for it.",
    "",
    "TIMESTAMPS",
    f"The 'date' index uses EIA-930's own \"{TIMEZONE}\" timezone facet to bucket hourly readings into "
    "a calendar day - EIA-930 reports each region's total once per US timezone (Arizona/Central/"
    f"Eastern/Mountain/Pacific), each with a different day boundary; {TIMEZONE} was picked as one "
    "consistent choice applied to every tab, not an official EIA cutoff or each ISO's own local zone.",
    "",
    "COVERAGE",
    f"Backed by a real start/end range query (unlike MISO's own API) - each tab is seeded back to "
    f"{HISTORY_START.isoformat()} (the dataset's own start) the first time that tab is built, then "
    f"kept current by re-pulling and upserting the trailing {ROLLING_WINDOW_DAYS} days on every run, "
    "which also catches EIA revising recently-published days.",
    "",
    "SOURCE",
    f"EIA (US Energy Information Administration) API v2, Form EIA-930 (Hourly Electric Grid Monitor): "
    f"{URL} - needs a free EIA API key (api.eia.gov), unlike MISO's public API or FRED's CSV endpoint.",
]
NOTES_SECTION_TITLES = {"UNITS", "TABS / SCOPE", "CATEGORIES", "TIMESTAMPS", "COVERAGE", "SOURCE"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    if not API_KEY:
        print("EIA_API_KEY is not set in the environment - aborting.", file=sys.stderr)
        sys.exit(1)

    today = date.today()
    sheets = {}
    region_fuel_cols = {}
    us_total_latest = None

    for respondent, sheet_name in REGIONS.items():
        print(f"[{sheet_name}] ({respondent})", file=sys.stderr)
        existing = load_archive(args.out, sheet_name)

        if existing.empty:
            print(f"  no existing tab - backfilling full history {HISTORY_START} to {today} ...",
                  file=sys.stderr)
            start = HISTORY_START
        else:
            start = max(HISTORY_START, today - timedelta(days=ROLLING_WINDOW_DAYS))
            print(f"  existing tab has {len(existing)} days - refreshing rolling window "
                  f"{start} to {today} ...", file=sys.stderr)

        rows = fetch_range(respondent, start, today)
        new_df = rows_to_wide(rows)
        if not new_df.empty:
            new_df = add_summary_columns(new_df)

        before_days = set(existing.index) if not existing.empty else set()
        combined = upsert(existing, new_df)
        new_days = sorted(set(combined.index) - before_days)

        print(f"  added {len(new_days)} new day(s){' - ' + str(new_days[-5:]) if new_days else ''}")
        if not combined.empty:
            print(f"  tab now has {len(combined)} days ({combined.index.min()} to {combined.index.max()})")

        sheets[sheet_name] = combined
        region_fuel_cols[sheet_name] = [
            c for c in combined.columns if c.endswith("_MWh") and c != "Total_MWh"
        ]
        if respondent == "US48" and not combined.empty:
            us_total_latest = max(combined.index)

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, NOTES_LINES, NOTES_SECTION_TITLES)
    format_date_columns(args.out, list(sheets.keys()))
    add_charts(args.out, region_fuel_cols)

    print(f"Saved to {args.out}")

    age = (date.today() - us_total_latest).days if us_total_latest else None
    if us_total_latest is None or age > STALE_AFTER_DAYS:
        print(f"STALE SOURCE: US_Total's newest saved day is {us_total_latest} ({age} days old) - "
              f"{URL} may have changed or stopped updating.", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
