"""
Shared helpers for the scheduled raw grid-operator daily generation pulls
(ARGENTINA_POWER_DAILY.py, URUGUAY_POWER_DAILY.py, BOLIVIA_POWER_DAILY.py).

Every country writes the same standard layout so the dashboard and
add_charts.power_daily() read them all the same way:
  sheet "Daily": date, Hydro_MWh, Gas_MWh, Wind_MWh, Solar_MWh, Coal_MWh,
                 Nuclear_MWh, Oil_MWh, Bioenergy_MWh, Other_MWh, Total_MWh
  (energy per day, MWh; fuels a country never has are left out)
"""

import datetime as dt
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]
HISTORY_START = dt.date(2021, 1, 1)


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


def missing_days(existing, start, end, refresh_days=3):
    """Days in start..end with no saved Total_MWh, plus the latest
    `refresh_days` saved days (sources revise the last few)."""
    have = set()
    if existing is not None and not existing.empty and "Total_MWh" in existing:
        have = {d.date() for d in existing.index[existing["Total_MWh"].notna()]}
    days = [start + dt.timedelta(days=i) for i in range((end - start).days + 1)]
    out = [d for d in days if d not in have]
    if have:
        last = max(have)
        out += [d for d in days if last - dt.timedelta(days=refresh_days) < d <= last and d in have]
    return sorted(set(out))


def ranges(days, max_len=31):
    """Sorted dates -> list of (first, last) runs of consecutive days, each at most max_len long."""
    out = []
    for d in sorted(days):
        if out and (d - out[-1][1]).days == 1 and (d - out[-1][0]).days < max_len:
            out[-1][1] = d
        else:
            out.append([d, d])
    return [tuple(r) for r in out]


def standardise(daily):
    """Fuel columns (bare names or *_MWh) -> the standard column order plus Total_MWh, rounded."""
    daily = daily.rename(columns={f: f"{f}_MWh" for f in FUELS})
    cols = [f"{f}_MWh" for f in FUELS if f"{f}_MWh" in daily.columns]
    out = daily[cols].astype(float)
    out["Total_MWh"] = out.sum(axis=1, min_count=1)
    out.index = pd.to_datetime(out.index)
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


def write(path, daily, notes, extra_sheets=None):
    """Daily first (the standard sheet), then any extra sheets, plus the Units tab."""
    sheets = {"Daily": daily}
    sheets.update(extra_sheets or {})
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    titles = {line for line in notes if line and line.isupper()}
    xlsx_notes.write_workbook(path, sheets, notes, titles)
    print(f"Saved {path}: {daily['Total_MWh'].notna().sum():,} days "
          f"({daily.index.min():%d-%b-%Y} to {daily.index.max():%d-%b-%Y})", flush=True)
    print(daily.tail(5).to_string(), flush=True)
