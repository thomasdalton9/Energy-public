"""
Shared pieces for the installed-capacity pulls (Peru, Bolivia, Uruguay,
Ecuador): the standard 'Monthly' layout, the workbook writer, and the
comparison against Ember's yearly capacity data.

Standard layout (what add_charts.power_capacity and the South America
master read): sheet 'Monthly', columns
    date, Hydro_MW, Gas_MW, Wind_MW, Solar_MW, Coal_MW, Nuclear_MW,
    Oil_MW, Bioenergy_MW, Other_MW, Total_MW
date = 1st of the month (annual-only sources: one row per year dated 1 January).
"""

import io
import os
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes
import xlsx_notes  # noqa: E402

FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Other"]
COLUMNS = [f"{f}_MW" for f in FUELS] + ["Total_MW"]
START = pd.Timestamp("2021-01-01")

EMBER_URL = ("https://storage.googleapis.com/emb-prod-bkt-publicdata/public-downloads/"
             "yearly_full_release_long_format.csv")
# Ember capacity variables -> our columns (Ember's 'Other Fossil' is oil products in these four countries;
# 'Other Renewables' is geothermal etc.).
EMBER_MAP = {"Hydro": "Hydro", "Gas": "Gas", "Wind": "Wind", "Solar": "Solar", "Coal": "Coal",
             "Nuclear": "Nuclear", "Other Fossil": "Oil", "Bioenergy": "Bioenergy", "Other Renewables": "Other"}


def standardise(df):
    """Index = month (Timestamp, 1st of month) named 'date'; all fuel columns present (0 if none); Total_MW
    recomputed as the sum of the fuel columns."""
    out = pd.DataFrame(index=pd.DatetimeIndex(df.index).normalize())
    for c in COLUMNS[:-1]:
        out[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0).values if c in df else 0.0
    out["Total_MW"] = out[COLUMNS[:-1]].sum(axis=1)
    out.index.name = "date"
    return out.sort_index().round(2)


def load_sheet(path, sheet):
    """A sheet written by write() (index in column A), or an empty frame."""
    try:
        d = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError) as e:
        print(f"No '{sheet}' in {path} yet ({type(e).__name__})", flush=True)
        return pd.DataFrame()
    return d


def load_monthly(path):
    d = load_sheet(path, "Monthly")
    if d.empty:
        return d
    d.index = pd.to_datetime(d.index)
    d.index.name = "date"
    return d


def ember_capacity(country):
    """Ember yearly capacity (GW) for one country: DataFrame index Year, columns FUELS + Total."""
    r = requests.get(EMBER_URL, timeout=(15, 300), headers={"User-Agent": "Mozilla/5.0"})
    r.raise_for_status()
    d = pd.read_csv(io.BytesIO(r.content), low_memory=False,
                    usecols=["Area", "Year", "Category", "Subcategory", "Variable", "Unit", "Value"])
    c = d[(d["Area"] == country) & (d["Category"] == "Capacity") & (d["Subcategory"] == "Fuel")]
    w = c.pivot_table(index="Year", columns="Variable", values="Value", aggfunc="sum")
    out = pd.DataFrame(index=w.index)
    for v, f in EMBER_MAP.items():
        out[f] = out.get(f, 0) + (w[v].fillna(0) if v in w else 0)
    out = out.reindex(columns=FUELS).fillna(0.0)
    out["Total"] = out.sum(axis=1)
    return out


def validation_lines(monthly, country):
    """Latest capacity by fuel vs Ember's latest year (and same year, if present). Lines for notes + log."""
    try:
        e = ember_capacity(country)
    except Exception as ex:  # noqa: BLE001
        return [f"Ember comparison unavailable ({type(ex).__name__}: {str(ex)[:120]})"]
    if e.empty or monthly.empty:
        return ["Ember comparison unavailable (no rows)"]
    last = monthly.index.max()
    ours = monthly.loc[last]
    year = int(last.year) if last.year in e.index else int(e.index.max())
    lines = [f"Latest row here ({last:%Y-%m}) vs Ember yearly capacity {year} (GW; diff = (ours - Ember) / Ember):"]
    for f in FUELS + ["Total"]:
        o = float(ours.get(f"{f}_MW", 0)) / 1000
        m = float(e.loc[year, f])
        if o == 0 and m == 0:
            continue
        pct = f"{(o - m) / m * 100:+.1f}%" if m else "n/a (Ember 0)"
        lines.append(f"  {f}: {o:.3f} GW here vs {m:.3f} GW Ember -> {pct}")
    return lines


def write(path, monthly, notes, extra_sheets=None, sections=("UNITS", "COVERAGE", "SOURCE", "MAPPING",
                                                                    "VALIDATION")):
    """Write the workbook: Units/notes tab, 'Monthly' (standard columns), then any extra raw sheets."""
    sheets = {"Monthly": standardise(monthly)}
    sheets.update(extra_sheets or {})
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    xlsx_notes.write_workbook(path, sheets, notes, set(sections))
    return sheets["Monthly"]


def unit_notes(granularity):
    return ["UNITS",
            "Capacity in MW (megawatts). 'Monthly' has one row per " + granularity + "; Total_MW = sum of the fuel "
            "columns. The chart shows the same columns in GW."]
