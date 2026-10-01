"""
Shared helpers for the Central America raw grid-operator daily generation pulls
(PANAMA_CND_GENERATION_DAILY.py, COSTA_RICA_CENCE_GENERATION_DAILY.py,
NICARAGUA_CNDC_GENERATION_DAILY.py).

Every workbook gets the standard layout the South & Central America master and
add_charts.power_daily() read:
  sheet "Daily": date, Hydro_MWh, Gas_MWh, Wind_MWh, Solar_MWh, Coal_MWh,
                 Nuclear_MWh, Oil_MWh, Bioenergy_MWh, Other_MWh, Total_MWh
  (MWh per day; every column is always present, 0 where a country has none)
plus a "Detail" sheet with the operator's own categories/units per day, and a
"Units" notes tab written with xlsx_notes.write_workbook() (atomic write).
"""

import datetime as dt
import os
import sys
import unicodedata

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes  # noqa: E402

FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]
COLUMNS = [f"{f}_MWh" for f in FUELS] + ["Total_MWh"]
HISTORY_START = dt.date(2021, 1, 1)
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")


def norm(s):
    """Lower-case, accents stripped, single spaces - for matching source labels."""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return " ".join(s.lower().split())


def load_sheet(path, sheet):
    """A saved sheet indexed by date (Timestamp), or an empty frame."""
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index, errors="coerce")
    df = df[df.index.notna()]
    df.index.name = "date"
    return df


def days_to_fetch(existing, start, end, refresh_days=14):
    """Days in start..end with no saved Total_MWh, plus the last `refresh_days`
    days up to `end` (operators revise recent days)."""
    have = set()
    if existing is not None and not existing.empty and "Total_MWh" in existing:
        have = {d.date() for d in existing.index[existing["Total_MWh"].notna()]}
    days = [start + dt.timedelta(days=i) for i in range((end - start).days + 1)]
    out = {d for d in days if d not in have}
    out |= {d for d in days if d > end - dt.timedelta(days=refresh_days)}
    return sorted(out)


def standardise(fuels):
    """Frame of fuel columns (bare names) -> every standard column, Total = sum of fuels."""
    out = pd.DataFrame(index=pd.to_datetime(fuels.index))
    for f in FUELS:
        out[f"{f}_MWh"] = fuels[f].astype(float).values if f in fuels else 0.0
    out = out.fillna(0.0)
    out["Total_MWh"] = out[[f"{f}_MWh" for f in FUELS]].sum(axis=1)
    out.index.name = "date"
    return out.sort_index().round(1)


def merge(new, existing):
    """New rows win over saved ones."""
    if existing is None or existing.empty:
        return new.sort_index()
    if new is None or new.empty:
        return existing.sort_index()
    out = pd.concat([existing[~existing.index.isin(new.index)], new])
    return out[~out.index.duplicated(keep="last")].sort_index()


def coverage_lines(daily):
    if daily.empty:
        return ["No data yet."]
    ok = daily["Total_MWh"].notna()
    first, last = daily.index[ok].min(), daily.index[ok].max()
    expected = pd.date_range(first, last, freq="D")
    missing = expected.difference(daily.index[ok])
    lines = [f"{first:%d-%b-%Y} to {last:%d-%b-%Y}: {ok.sum():,} days with data, {len(missing):,} missing."]
    if len(missing):
        lines.append("Missing days (source had no usable file/answer): "
                     + ", ".join(f"{d:%Y-%m-%d}" for d in missing[:60]) + (" ..." if len(missing) > 60 else ""))
    return lines


def write(path, daily, notes, detail=None):
    """'Daily' (standard columns), then 'Detail', with the notes tab first.
    `notes` is a list of strings; ALL-CAPS lines are highlighted as section titles."""
    daily = daily.reindex(columns=COLUMNS)
    sheets = {"Daily": daily}
    if detail is not None and not detail.empty:
        sheets["Detail"] = detail.sort_index()
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    titles = {line for line in notes if line and line.isupper()}
    xlsx_notes.write_workbook(path, sheets, notes, titles)
    print(f"Saved {path}: {daily['Total_MWh'].notna().sum():,} days "
          f"({daily.index.min():%d-%b-%Y} to {daily.index.max():%d-%b-%Y})", flush=True)


def print_mapping(mapping, title):
    print(f"Category mapping ({title}):", flush=True)
    for k, v in mapping.items():
        print(f"  {k!s:<40} -> {v}", flush=True)


def print_monthly(daily, n=4):
    m = daily[COLUMNS].resample("MS").sum(min_count=1) / 1000
    print("Last months, GWh:", flush=True)
    print(m.tail(n).round(1).to_string(), flush=True)
