"""
Egypt's daily grid load from the Ministry of Electricity and Renewable Energy (MOEE) home page 'Grid Loads' panel
(http://www.moee.gov.eg/english_new/home.aspx; https is refused): yesterday's peak and minimum load (MW), a temperature
(degrees C) and today's expected evening peak (MW), under a page date (span ids ctl00_content_h_load_day / h_max_load /
h_min_load / h_temp / h_exp_load).

NO HISTORY: the panel shows only the current day, so this runs every day (twice, so one failed run loses nothing) and
the committed workbook is the history store from the first run (11 Oct 2026). A missed day stays a gap.

DATES: a page dated D gives the peak and minimum of D-1 ('Yesterday') and the expected peak of D. The 'Daily' sheet is
keyed by OPERATING day: Peak_Load_MW / Min_Load_MW / Temperature_C from the page dated the day after, Expected_Peak_MW
from the page dated that day. The panel does not say which day the temperature is for or where it is measured; it is
stored with the peak it is shown next to. A page whose date is more than 3 days old is not stored (the panel has
stopped updating) and the run fails so the gap is seen.

Probe: discovery_archive/middle_east/EGYPT_MOEE_GRID_LOAD_PROBE.py (and EGYPT_ISRAEL_GAS_POWER_PROBE.py).
Workflow: .github/workflows/egypt_grid_load.yml.
"""

import argparse
import os
import re
import sys
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

URL = "http://www.moee.gov.eg/english_new/home.aspx"
SPANS = {"h_load_day": "page_date", "h_max_load": "peak", "h_min_load": "min", "h_temp": "temp", "h_exp_load": "expected"}
COLUMNS = ["Peak_Load_MW", "Min_Load_MW", "Temperature_C", "Expected_Peak_MW", "Page_date_peak", "Page_date_expected"]
MAX_PAGE_AGE_DAYS = 3
PLAUSIBLE_MW = (5000, 80000)   # Egypt's peak is ~33-39 GW; anything outside this is a parsing error

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "Data and Chart Outputs", "egypt_grid_load_daily.xlsx")


def fetch_panel():
    r = requests.get(URL, timeout=60, headers={"User-Agent": "Mozilla/5.0"})
    r.raise_for_status()
    vals = {}
    for sid, txt in re.findall(r'<span[^>]*id="ctl00_content_(h_[a-z_]+)"[^>]*>(.*?)</span>', r.text, re.S):
        if sid in SPANS:
            vals[SPANS[sid]] = " ".join(re.sub(r"<[^>]+>", " ", txt).split())
    missing = [k for k in SPANS.values() if not vals.get(k)]
    if missing:
        raise RuntimeError(f"Grid Loads panel fields missing: {missing} (got {vals}) - page layout changed?")
    page_date = datetime.strptime(vals["page_date"], "%Y/%m/%d").date()

    def num(key, lo=None, hi=None):
        v = float(vals[key].replace(",", ""))
        if lo is not None and not lo <= v <= hi:
            raise RuntimeError(f"{key} = {v} outside {lo}-{hi} - parsing error?")
        return v

    return page_date, {"peak": num("peak", *PLAUSIBLE_MW), "min": num("min", *PLAUSIBLE_MW),
                       "temp": num("temp", -10, 60), "expected": num("expected", *PLAUSIBLE_MW)}


def load(path):
    try:
        df = pd.read_excel(path, sheet_name="Daily", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame(columns=COLUMNS, dtype=object)
    df.index = pd.to_datetime(df.index).date
    return df.reindex(columns=COLUMNS).astype(object)


NOTES_LINES = [
    "UNITS",
    "Peak_Load_MW / Min_Load_MW: the national grid's highest and lowest load of the operating day, MW, as shown under "
    "'Yesterday' on the Ministry of Electricity's home page. Expected_Peak_MW: the ministry's expected evening peak for "
    "that day, shown the same morning. Temperature_C: the temperature the panel shows next to yesterday's loads (the "
    "page does not say which day or place it refers to).",
    "",
    "DATES",
    "Rows are operating days. A page dated D gives day D-1's peak and minimum and day D's expected peak, so the "
    "latest row usually has an expected peak only. Page_date_peak / Page_date_expected record the page each value came "
    "from.",
    "",
    "HISTORY",
    "The page shows only the current day and keeps no archive, so this workbook IS the history: it starts on 10 Oct "
    "2026 (first pull 11 Oct 2026) and a day the pull missed stays blank - nothing is filled or estimated.",
    "",
    "SOURCE",
    f"Ministry of Electricity and Renewable Energy (MOEE), Egypt, home page 'Grid Loads' panel: {URL}",
]
NOTES_SECTION_TITLES = {"UNITS", "DATES", "HISTORY", "SOURCE"}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()

    page_date, v = fetch_panel()
    today = datetime.now(timezone.utc).date()
    print(f"MOEE Grid Loads page dated {page_date}: yesterday peak {v['peak']:.0f} MW, min {v['min']:.0f} MW, "
          f"{v['temp']:.0f} C; expected peak today {v['expected']:.0f} MW")
    if (today - page_date).days > MAX_PAGE_AGE_DAYS:
        sys.exit(f"Page date {page_date} is more than {MAX_PAGE_AGE_DAYS} days old - the panel has stopped updating; "
                 "nothing stored")

    df = load(args.out)
    before = df.copy()
    op_day = page_date - timedelta(days=1)
    for day in (op_day, page_date):
        if day not in df.index:
            df.loc[day] = [None] * len(COLUMNS)
    df.loc[op_day, ["Peak_Load_MW", "Min_Load_MW", "Temperature_C", "Page_date_peak"]] = \
        [v["peak"], v["min"], v["temp"], page_date.isoformat()]
    df.loc[page_date, ["Expected_Peak_MW", "Page_date_expected"]] = [v["expected"], page_date.isoformat()]
    df = df.sort_index()
    if df.equals(before.sort_index()) and os.path.exists(args.out):
        print("No change (same page as the last run).")
        return
    df.index = pd.Index(df.index, name="date")
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, {"Daily": df[COLUMNS]}, NOTES_LINES, NOTES_SECTION_TITLES)
    print(f"Saved {len(df)} day(s), {min(df.index)} to {max(df.index)} -> {args.out}")


if __name__ == "__main__":
    main()
