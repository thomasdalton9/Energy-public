"""
Pull South Africa's power generation mix by fuel type (coal, gas, oil,
nuclear, hydro, pumped storage, wind, solar, biomass) from Eskom's
public "Station Build Up" CSV, and maintain a growing local daily
archive.

Data source: https://www.eskom.co.za/dataportal/wp-content/uploads/{year}/{month}/Station_Build_Up.csv
This is NOT a historical archive to backfill from. Despite the URL
being dated by year/month, the file itself is always just a rolling
window of the last ~7 days of hourly readings - the source page is
literally titled "Station Build Up for the last 7 days", and Eskom's
own reference parser (in electricitymaps-contrib) refuses to look
back further than a week. There's no --from-date/--to-date here
because there's no range to request; Eskom simply doesn't expose
deeper history publicly.

Each hourly row already reports that hour's own generation value (not
a cumulative running total), so no subtraction is needed to isolate
"new" data - this script just re-fetches the current window on every
run and upserts by date into a local archive CSV (--out), keeping
whatever days it already has and adding/refreshing whatever the
current window shows. Run it at least once a week (the window is 7
days) so no day is ever skipped; running it more often (e.g. daily) is
harmless, just redundant. History only starts accumulating from
whenever you start running this script.
"""

import argparse
import csv
import io
import sys
from datetime import date, datetime, timezone

import openpyxl
import pandas as pd
import requests

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes/api_keys

import xlsx_notes

# No-cache headers plus a timestamp query param (see fetch_current_window):
# two scheduled runs three days apart got byte-identical windows ending
# 23 Sep, while Electricity Maps' own Eskom parser still uses this same
# URL - so a cached copy in front of Eskom's site is a plausible cause.
HEADERS = {
    "User-Agent": "gas-demand-scripts/1.0",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
}

# Column index (within the row, after the leading datetime column) to
# category. Indices not listed here are storage/pumping-consumption/
# interruption columns, not generation, and are skipped.
COLUMN_CATEGORY = {
    0: "coal",  # Thermal_Gen_Excl_Pumping_and_SCO
    6: "nuclear",  # Nuclear_Generation
    8: "oil",  # Eskom_OCGT_Generation
    9: "gas",  # Eskom_Gas_Generation
    10: "oil",  # Dispatchable_IPP_OCGT
    11: "hydro",  # Hydro_Water_Generation
    12: "pumped_storage",  # Pumped_Water_Generation
    16: "wind",  # Wind
    17: "solar",  # PV
    18: "solar",  # CSP
    19: "biomass",  # Other_RE
}
CATEGORIES = ["coal", "gas", "oil", "nuclear", "hydro", "pumped_storage", "wind", "solar", "biomass"]


def make_session():
    session = requests.Session()
    session.headers.update(HEADERS)
    return session


def get_url():
    today = datetime.now(timezone.utc)
    return f"https://www.eskom.co.za/dataportal/wp-content/uploads/{today:%Y}/{today:%m}/Station_Build_Up.csv"


def fetch_current_window(session):
    """Fetch and parse Eskom's current rolling ~7-day hourly CSV."""
    response = session.get(get_url(), params={"t": int(datetime.now(timezone.utc).timestamp())}, timeout=30)
    response.raise_for_status()
    reader = csv.reader(io.StringIO(response.text))
    header = next(reader, None)
    if header is None:
        return pd.DataFrame()

    rows = []
    for row in reader:
        if not row or not row[0]:
            continue
        try:
            dt = pd.to_datetime(row[0])
        except (ValueError, TypeError):
            continue
        values = row[1:]
        entry = {"datetime": dt}
        for category in CATEGORIES:
            entry[category] = 0.0
        for index, category in COLUMN_CATEGORY.items():
            if index >= len(values):
                continue
            raw = (values[index] or "").strip()
            if not raw:
                continue
            try:
                entry[category] += max(float(raw), 0.0)
            except ValueError:
                continue
        rows.append(entry)

    if not rows:
        return pd.DataFrame()
    return pd.DataFrame.from_records(rows).set_index("datetime").sort_index()


# Eskom's CSV includes rows for hours that haven't been reported yet
# (blank values), which the parser turns into all-zero hours. A day made
# mostly of those averaged down to near-zero and got saved as if it were
# real - so only hours with any generation count, and only days with a
# near-complete set of hours are kept.
MIN_HOURS_PER_DAY = 22


def to_daily_mean(hourly_df):
    if hourly_df.empty:
        return pd.DataFrame()
    reported = hourly_df[hourly_df[CATEGORIES].sum(axis=1) > 0]
    grouped = reported.groupby(reported.index.date)[CATEGORIES]
    counts = grouped.size()
    daily = grouped.mean()[counts >= MIN_HOURS_PER_DAY]
    daily.index.name = "date"
    return daily


# Second and third tabs: system-level context from other files on the
# same Eskom data portal (found by ESKOM_DATAPORTAL_DISCOVERY.py). Eskom
# publishes nothing per power station, so these are whole-system.
#   System     - daily mean of hourly demand and unplanned outages (MW)
#   Weekly EAF - Eskom fleet availability breakdown (% of capacity)
# Same rolling-window upsert as the Data tab. A failure here is only a
# warning: it must never stop the generation mix archive updating.
SYSTEM_SOURCES = {
    # file name: (datetime column, {source column: our column})
    "System_hourly_actual_and_forecasted_demand": (
        "DateTimeKey",
        {"Residual Demand": "residual_demand_mw", "RSA Contracted Demand": "rsa_contracted_demand_mw"},
    ),
    "Hourly_UCLF_and_OCLF_Trend": ("DateTimeKey", {"Hourly UCLF+OCLF": "unplanned_outages_mw"}),
}
SYSTEM_COLUMNS = ["residual_demand_mw", "rsa_contracted_demand_mw", "unplanned_outages_mw"]

WEEKLY_EAF_SOURCE = "Weekly_Eskom_generation_capacity_breakdown"
WEEKLY_EAF_COLUMNS = {
    "Weekly EAF": "eaf_pct",
    "Weekly PCLF": "planned_outages_pct",
    "Weekly UCLF": "unplanned_outages_pct",
    "Weekly OCLF": "other_outages_pct",
}


# Pumped storage and Load shedding tabs both come from this one file,
# which is hourly and ~8 days long, lagging about a week. Its "Gen Unit
# Hours" columns are NOT hours run: they move smoothly hour to hour,
# rising overnight (pumping) and falling through the day (generating),
# between fixed "Min Hours" and "Max Hours" lines - i.e. the upper
# reservoir's stored energy, as unit-hours of generation left.
PUMPED_FILE = "Pumped_storage_gen_hours_gas_generation_and_manual_load_reduction"
PUMPED_STATIONS = ["Drakensberg", "Ingula", "Palmiet"]
MLR_COLUMN = "Manual Load Reduction(MLR)"


def portal_url(name):
    today = datetime.now(timezone.utc)
    return f"https://www.eskom.co.za/dataportal/wp-content/uploads/{today:%Y}/{today:%m}/{name}.csv"


def fetch_portal_csv(session, name):
    response = session.get(portal_url(name), params={"t": int(datetime.now(timezone.utc).timestamp())}, timeout=30)
    response.raise_for_status()
    return pd.read_csv(io.StringIO(response.text))


def fetch_system_daily(session):
    """Daily mean of each hourly system series. Each column is kept only
    for days with MIN_HOURS_PER_DAY actual (non-blank) hours - the demand
    file runs into the future with blank actuals and forecast values."""
    columns = []
    for name, (time_col, mapping) in SYSTEM_SOURCES.items():
        raw = fetch_portal_csv(session, name)
        hourly = raw[list(mapping)].apply(pd.to_numeric, errors="coerce").rename(columns=mapping)
        hourly.index = pd.to_datetime(raw[time_col])
        for col in hourly.columns:
            series = hourly[col].dropna()
            grouped = series.groupby(series.index.date)
            daily = grouped.mean()[grouped.size() >= MIN_HOURS_PER_DAY]
            columns.append(daily.rename(col))
    if not columns:
        return pd.DataFrame()
    df = pd.concat(columns, axis=1).reindex(columns=SYSTEM_COLUMNS)
    df.index.name = "date"
    return df.dropna(how="all").sort_index()


def fetch_weekly_eaf(session):
    raw = fetch_portal_csv(session, WEEKLY_EAF_SOURCE)
    df = raw[list(WEEKLY_EAF_COLUMNS)].apply(pd.to_numeric, errors="coerce").rename(columns=WEEKLY_EAF_COLUMNS)
    df.index = pd.to_datetime(raw["Week Start Date"]).dt.date
    df.index.name = "week_start"
    return df.dropna(how="all").sort_index()


_pumped_cache = {}


def fetch_pumped_hourly(session):
    """The pumped storage/MLR file, fetched once per run for both tabs,
    as an hourly frame with complete days only."""
    if "df" not in _pumped_cache:
        raw = fetch_portal_csv(session, PUMPED_FILE)
        df = raw.drop(columns=["Date"]).apply(pd.to_numeric, errors="coerce")
        df.index = pd.to_datetime(raw["Date"])
        df = df.dropna(subset=[f"{s} Gen Unit Hours" for s in PUMPED_STATIONS], how="all")  # unreported hours
        counts = df.groupby(df.index.date).size()
        complete = set(counts[counts >= MIN_HOURS_PER_DAY].index)
        _pumped_cache["df"] = df[[d in complete for d in df.index.date]]
    return _pumped_cache["df"]


def fetch_pumped_storage(session):
    hourly = fetch_pumped_hourly(session)
    days = hourly.groupby(hourly.index.date)
    out = {}
    for station in PUMPED_STATIONS:
        level = hourly[f"{station} Gen Unit Hours"]
        full = hourly[f"{station} Max Hours"]
        key = station.lower()
        out[f"{key}_close_unit_hours"] = days[f"{station} Gen Unit Hours"].last()
        out[f"{key}_close_pct_full"] = (level / full * 100).groupby(hourly.index.date).last()
        out[f"{key}_low_pct_full"] = (level / full * 100).groupby(hourly.index.date).min()
    df = pd.DataFrame(out)
    df.index.name = "date"
    return df.sort_index()


def fetch_load_shedding(session):
    hourly = fetch_pumped_hourly(session)
    mlr = hourly[MLR_COLUMN].fillna(0)
    days = mlr.groupby(mlr.index.date)
    df = pd.DataFrame({
        "load_shed_avg_mw": days.mean(),
        "load_shed_peak_mw": days.max(),
        "hours_with_load_shedding": days.apply(lambda s: int((s > 0).sum())),
    })
    df.index.name = "date"
    return df.sort_index()


def load_sheet(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError):  # ValueError: sheet not in workbook yet
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index).date
    return df


def upsert_columns(existing, new, index_name):
    """Like upsert, but cell by cell: a fresh value replaces an old one,
    while a column the fresh window has no complete day for keeps the
    archived value instead of being blanked."""
    if existing.empty:
        combined = new
    elif new.empty:
        combined = existing
    else:
        combined = new.combine_first(existing)[list(new.columns)]
    combined.index.name = index_name
    return combined.sort_index()


def update_context_sheet(path, sheet, fetch, session, index_name):
    existing = load_sheet(path, sheet)
    try:
        fresh = fetch(session)
    except Exception as exc:  # network, moved file, renamed column...
        print(f"WARNING: {sheet} tab not refreshed ({type(exc).__name__}: {exc}) - keeping archived rows.", file=sys.stderr)
        fresh = pd.DataFrame()
    return upsert_columns(existing, fresh, index_name)


def load_archive(path):
    # Reads the date from the first column whatever its header says - an
    # earlier version saved it without a header ("Unnamed: 0"), and
    # looking it up by name then failed, so the whole archive was
    # silently treated as empty and overwritten.
    try:
        df = pd.read_excel(path, sheet_name="Data", index_col=0)
    except FileNotFoundError:
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index).date
    df.index.name = "date"
    return df[df[CATEGORIES].sum(axis=1) > 0]


def upsert(existing, new):
    if existing.empty:
        return new
    if new.empty:
        return existing
    combined = pd.concat([existing, new])
    combined = combined[~combined.index.duplicated(keep="last")]
    combined.index.name = "date"
    return combined.sort_index()


# Eskom's CSV normally refreshes daily. If the newest complete day it
# returns is older than this (Eskom itself runs ~4 days behind, so 6 leaves margin while still catching a real outage before its 7-day window rolls past), the source has gone stale (moved URL, stopped
# publishing) - fail the run so GitHub flags it, rather than quietly
# re-saving the same week forever.
STALE_AFTER_DAYS = 6


NOTES_LINES = [
    "UNITS",
    "Every category column is in MW, the daily mean of Eskom's hourly generation readings - "
    "'average MW for that day', not total daily energy.",
    "",
    "CATEGORIES",
    "coal, gas, oil, nuclear, hydro, wind, solar, biomass: generation by fuel/technology type.",
    "pumped_storage: pumped-hydro storage discharge.",
    "",
    "COVERAGE",
    "This is a rolling archive, not a one-shot historical pull - Eskom's source CSV is always "
    "just the last ~7 days, so this script upserts by date into the Data tab on every run, "
    "building up history from whenever you first started running it. Run it at least weekly "
    "so no day is skipped.",
    "",
    "SYSTEM TAB (daily mean of hourly values, MW)",
    "residual_demand_mw: demand met by Eskom-dispatched plant, i.e. excluding renewable IPPs.",
    "rsa_contracted_demand_mw: residual demand plus renewable IPP output - total contracted demand.",
    "unplanned_outages_mw: Eskom capacity out on unplanned (UCLF) plus other (OCLF) outages - "
    "a fleet health indicator. Eskom publishes nothing per power station, so this is whole-fleet.",
    "",
    "WEEKLY EAF TAB (% of Eskom installed capacity, by week starting)",
    "eaf_pct: Energy Availability Factor - share of capacity available to generate.",
    "planned_outages_pct (PCLF), unplanned_outages_pct (UCLF), other_outages_pct (OCLF): "
    "the losses; the four add up to ~100%.",
    "",
    "PUMPED STORAGE TAB (upper reservoir level, per station)",
    "Eskom reports each station's stored water as 'generating unit-hours' left - how many hours "
    "of generation the upper reservoir holds, summed over units. It is a storage level, not hours run: "
    "it rises overnight while pumping and falls during the day while generating.",
    "*_close_unit_hours: level at the last hour of the day. *_close_pct_full: that level as % of the "
    "station's full reservoir (Drakensberg 102, Ingula 67, Palmiet 58.7 unit-hours). "
    "*_low_pct_full: the lowest level in the day - how deep the peak drew it down.",
    "",
    "LOAD SHEDDING TAB (Eskom manual load reduction, MLR)",
    "load_shed_avg_mw: daily mean MW shed. load_shed_peak_mw: highest hour. "
    "hours_with_load_shedding: hours in the day with any MLR. All zero while there is no load shedding.",
    "Coverage lags about a week behind the other tabs - Eskom's source file does.",
    "",
    "SOURCE",
    "Eskom data portal CSVs: Station_Build_Up (Data), System_hourly_actual_and_forecasted_demand "
    "and Hourly_UCLF_and_OCLF_Trend (System), Weekly_Eskom_generation_capacity_breakdown (Weekly EAF), "
    "Pumped_storage_gen_hours_gas_generation_and_manual_load_reduction (Pumped storage, Load shedding).",
]
NOTES_SECTION_TITLES = {
    "UNITS",
    "CATEGORIES",
    "COVERAGE",
    "SYSTEM TAB (daily mean of hourly values, MW)",
    "WEEKLY EAF TAB (% of Eskom installed capacity, by week starting)",
    "PUMPED STORAGE TAB (upper reservoir level, per station)",
    "LOAD SHEDDING TAB (Eskom manual load reduction, MLR)",
    "SOURCE",
}


def format_date_column(path, sheets=("Data",)):
    # pandas' default "YYYY-MM-DD" format in a narrow column was shown by
    # some viewers (e.g. phone previews) as "26/9" - year and month only -
    # making every September row look identical. A spelled-out month is
    # unambiguous in any viewer or locale.
    wb = openpyxl.load_workbook(path)
    for sheet in sheets:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for (cell,) in ws.iter_rows(min_row=2, max_col=1):
            cell.number_format = "dd-mmm-yyyy"
        ws.column_dimensions["A"].width = 14
    wb.save(path)


REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(REPO_ROOT, "output", "south_africa_generation_mix_daily.xlsx")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=DEFAULT_OUT)
    args = parser.parse_args()

    session = make_session()
    print(f"Fetching current Station Build Up window from {get_url()}...", file=sys.stderr)
    hourly = fetch_current_window(session)
    if hourly.empty:
        print("No data returned.", file=sys.stderr)
        sys.exit(1)
    new_daily = to_daily_mean(hourly)

    existing = load_archive(args.out)
    before_days = set(existing.index) if not existing.empty else set()
    combined = upsert(existing, new_daily)
    new_or_updated = sorted(set(combined.index) - before_days)

    system = update_context_sheet(args.out, "System", fetch_system_daily, session, "date")
    weekly_eaf = update_context_sheet(args.out, "Weekly EAF", fetch_weekly_eaf, session, "week_start")
    pumped = update_context_sheet(args.out, "Pumped storage", fetch_pumped_storage, session, "date")
    load_shedding = update_context_sheet(args.out, "Load shedding", fetch_load_shedding, session, "date")

    sheets = {"Data": combined}
    for name, df in [("System", system), ("Weekly EAF", weekly_eaf), ("Pumped storage", pumped), ("Load shedding", load_shedding)]:
        if not df.empty:
            sheets[name] = df
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, NOTES_LINES, NOTES_SECTION_TITLES)
    format_date_column(args.out, sheets)
    print(f"Archive now has {len(combined)} days ({len(new_or_updated)} new since last run). Saved to {args.out}")
    print(combined.tail())
    print(f"\nSystem tab: {len(system)} days")
    print(system.tail())
    print(f"\nWeekly EAF tab: {len(weekly_eaf)} weeks")
    print(weekly_eaf.tail())
    print(f"\nPumped storage tab: {len(pumped)} days")
    print(pumped.tail())
    print(f"\nLoad shedding tab: {len(load_shedding)} days")
    print(load_shedding.tail())

    latest = max(new_daily.index) if not new_daily.empty else None
    age = (date.today() - latest).days if latest else None
    if latest is None or age > STALE_AFTER_DAYS:
        print(
            f"STALE SOURCE: newest complete day from Eskom is {latest} "
            f"({age} days old) - {get_url()} may have stopped updating or moved.",
            file=sys.stderr,
        )
        sys.exit(1)


if __name__ == "__main__":
    main()
