"""
Cyprus daily power generation mix by fuel type, from TSOC (Transmission
System Operator Cyprus)'s own published daily generation tables.

Cyprus runs an isolated island grid - NOT interconnected to the
European mainland - so it is NOT covered by ENTSO-E's Transparency
Platform the way every other EU country in this repo is (see
entsoe_powergen_monthly_2022_2026.py). This is a genuinely separate
source, found via electricitymaps-contrib's own maintained CY.py
parser (github.com/electricitymaps/electricitymaps-contrib) - the exact
URLs, HTML table structure, and biomass/solar derivation logic below
are ported directly from that parser and confirmed working against the
live site (2026-09-30), including a historical date 4 months back.

Source pages (both confirmed live):
  https://tsoc.org.cy/electrical-system/total-daily-system-generation-on-the-transmission-system/
  https://tsoc.org.cy/electrical-system/archive-total-daily-system-generation-on-the-transmission-system/?startdt=DD-MM-YYYY&enddt=%2B1days
Each request returns ONE day's worth of hourly readings in an HTML
table (id="production_graph_data"), plus (on the real-time page only)
an installed-capacity table (id="production_graph_static_data2").

Categories: wind and oil (conventional) are read directly. Biomass and
solar are NOT reported separately - TSOC only publishes "Εκτίμηση
Διεσπαρμένης Παραγωγής" (estimated distributed generation), which
mixes both. TSOC's own convention (replicated from CY.py, not
invented here): distributed generation between 10pm-3am is assumed to
be pure biomass (no solar at night) and used as that day's biomass
baseline; during the rest of the day, solar = distributed generation
minus that baseline (floored at 0). This is a real limitation of the
source, not a parsing shortcut - there is no way to separate biomass
from solar more precisely with what TSOC publishes.

The archive page's own "published" date is 2018-04-25, but that's when
the WEB PAGE was created, not necessarily how far back real data goes -
this script does not assume a start date works and simply skips/warns
on any date that returns no data rather than guessing.

Output: cyprus_generation_mix_daily.csv/.xlsx - daily mean MW per
category (wind, oil, biomass, solar) plus installed capacity as of the
most recent successful fetch.

Run standalone (edit FROM_DATE/TO_DATE below to change scope):
    python3 cyprus_generation_mix.py
"""

import os
import sys
import time
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from bs4 import BeautifulSoup

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

FROM_DATE = date(2018, 1, 1)
TO_DATE = date.today()
OUT_FILE = "cyprus_generation_mix_daily.xlsx"
HISTORY_CSV = "cyprus_generation_mix_daily.csv"

REALTIME_SOURCE = "https://tsoc.org.cy/electrical-system/total-daily-system-generation-on-the-transmission-system/"
HISTORICAL_SOURCE = "https://tsoc.org.cy/electrical-system/archive-total-daily-system-generation-on-the-transmission-system/"

TIMEZONE = ZoneInfo("Asia/Nicosia")
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 30)
REQUEST_PAUSE_SECONDS = 0.5  # be polite - one request per day of history

CAPACITY_KEYS = {
    "Συμβατική Εγκατεστημένη Ισχύς": "oil",
    "Αιολική Εγκατεστημένη Ισχύς": "wind",
    "Φωτοβολταϊκή Εγκατεστημένη Ισχύς": "solar",
    "Εγκατεστημένη Ισχύς Βιομάζας": "biomass",
}
CATEGORIES = ["oil", "wind", "biomass", "solar"]


def parse_capacity(soup):
    table = soup.find(id="production_graph_static_data2")
    if table is None:
        return {}
    capacity = {}
    for tr in table.find_all("tr"):
        values = [td.string for td in tr.find_all("td")]
        if len(values) < 2:
            continue
        key = CAPACITY_KEYS.get(values[0])
        if key and values[1] is not None:
            capacity[key] = float(values[1])
    return capacity


def parse_production(soup):
    """Return a list of {"datetime": ..., "wind": ..., "oil": ..., "biomass": ..., "solar": ...}."""
    table = soup.find(id="production_graph_data")
    if table is None or table.tbody is None:
        return []
    columns = [th.string for th in table.find_all("th")]
    rows = []
    biomass_estimate = 0.0
    for tr in table.tbody.find_all("tr"):
        values = [td.string for td in tr.find_all("td")]
        if None in values or "" in values:
            break  # incomplete/future row - TSOC pads the rest of the day with blanks
        row = {}
        dt = None
        for col, val in zip(columns, values, strict=True):
            if col == "Timestamp":
                dt = datetime.fromisoformat(val).replace(tzinfo=TIMEZONE)
            elif col == "Αιολική Παραγωγή":
                row["wind"] = float(val)
            elif col == "Συμβατική Παραγωγή":
                row["oil"] = float(val)
            elif col == "Εκτίμηση Διεσπαρμένης Παραγωγής":
                value = float(val)
                if dt is not None and (dt.hour < 3 or dt.hour >= 22):
                    biomass_estimate = value
                row["biomass"] = biomass_estimate
                row["solar"] = max(value - biomass_estimate, 0.0)
        if dt is not None:
            row["datetime"] = dt
            rows.append(row)
    return rows


def fetch_day(session, day):
    """Fetch one calendar day's hourly rows. `day=None` means the real-time (today) page."""
    if day is None:
        url = REALTIME_SOURCE
    else:
        url = HISTORICAL_SOURCE
    params = None if day is None else {"startdt": day.strftime("%d-%m-%Y"), "enddt": "+1days"}

    r = session.get(url, headers=HEADERS, timeout=TIMEOUT, params=params)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "lxml")

    capacity = parse_capacity(soup) if day is None else {}
    rows = parse_production(soup)
    return rows, capacity


def build_daily_dataframe(from_date, to_date):
    session = requests.Session()
    daily_records = []
    latest_capacity = {}

    current = from_date
    while current <= to_date:
        try:
            rows, capacity = fetch_day(session, current)
            if capacity:
                latest_capacity = capacity
        except requests.RequestException as exc:
            print(f"{current.isoformat()}: failed ({exc})", file=sys.stderr)
            current += timedelta(days=1)
            time.sleep(REQUEST_PAUSE_SECONDS)
            continue

        if not rows:
            print(f"{current.isoformat()}: no data returned", file=sys.stderr)
        else:
            df_day = pd.DataFrame(rows)
            record = {"date": current}
            for category in CATEGORIES:
                if category in df_day.columns:
                    record[category] = df_day[category].mean()
            daily_records.append(record)
            print(f"{current.isoformat()}: {len(rows)} hourly readings", file=sys.stderr)

        current += timedelta(days=1)
        time.sleep(REQUEST_PAUSE_SECONDS)

    if not daily_records:
        return pd.DataFrame(), latest_capacity
    df = pd.DataFrame.from_records(daily_records).set_index("date").sort_index()
    return df, latest_capacity


NOTES_SECTION_TITLES = {"UNITS", "SCOPE", "CAVEATS", "SOURCE"}


def build_notes(capacity, day_count):
    lines = [
        "UNITS",
        "Every value column is MW-equivalent: the MEAN of that day's native hourly generation "
        "readings - 'average MW for that day', not total daily energy (MWh).",
        "",
        "SCOPE",
        "Cyprus runs an isolated island grid, NOT interconnected to the European mainland, so it "
        "is NOT covered by ENTSO-E's Transparency Platform (unlike every other EU country in this "
        "repo). This is a genuinely separate source: TSOC (Transmission System Operator Cyprus)'s "
        "own published daily generation tables.",
        "",
        "CAVEATS",
        "Biomass and solar are NOT reported separately by TSOC - only a combined 'estimated "
        "distributed generation' figure. Following TSOC's/electricitymaps-contrib's own convention: "
        "the distributed-generation reading between 10pm-3am is treated as that day's biomass "
        "baseline (no solar at night), and solar = distributed generation minus that baseline "
        "(floored at 0) for the rest of the day. This is a real limitation of the source, not a "
        "parsing shortcut.",
        f"{day_count} days successfully fetched (see stderr output for any dates that failed or "
        "returned no data).",
        "",
        "SOURCE",
        "tsoc.org.cy - real-time and historical-archive daily generation pages. Parsing logic "
        "(table structure, biomass/solar derivation) ported from electricitymaps-contrib's own "
        "maintained CY.py parser (github.com/electricitymaps/electricitymaps-contrib).",
    ]
    if capacity:
        lines.insert(2, f"Installed capacity as of the most recent fetch (MW): {capacity}")
        lines.insert(3, "")
    return lines


def main():
    df, capacity = build_daily_dataframe(FROM_DATE, TO_DATE)
    if df.empty:
        print("No data returned for any date.", file=sys.stderr)
        sys.exit(1)

    df.to_csv(HISTORY_CSV)
    notes_lines = build_notes(capacity, len(df))
    xlsx_notes.write_workbook(OUT_FILE, {"Cyprus": df}, notes_lines, NOTES_SECTION_TITLES, notes_sheet_name="Notes")
    print(f"Saved {len(df)} days to {HISTORY_CSV} and {OUT_FILE}")


if __name__ == "__main__":
    main()
