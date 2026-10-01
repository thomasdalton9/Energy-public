"""
Shared helpers for the installed-capacity pulls (BRAZIL_ANEEL_CAPACITY_MONTHLY.py,
ARGENTINA_CAMMESA_CAPACITY.py, CHILE_CNE_CAPACITY.py, COLOMBIA_XM_CAPACITY.py).

Every country writes the same standard layout, read by add_charts.power_capacity()
and the South America master:
  sheet "Monthly": date (1st of month), Hydro_MW, Gas_MW, Wind_MW, Solar_MW, Coal_MW,
                   Nuclear_MW, Oil_MW, Bioenergy_MW, Other_MW, Total_MW
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
# Ember yearly capacity 'Variable' -> standard fuel (Ember has no storage in its fuel list)
EMBER_TO_FUEL = {"Hydro": "Hydro", "Gas": "Gas", "Wind": "Wind", "Solar": "Solar", "Coal": "Coal",
                 "Nuclear": "Nuclear", "Other Fossil": "Oil", "Bioenergy": "Bioenergy", "Other Renewables": "Other"}


def standard(by_fuel):
    """DataFrame indexed by month (any fuel columns named like FUELS) -> the standard Monthly layout, MW."""
    out = pd.DataFrame(index=pd.DatetimeIndex(by_fuel.index, name="date"))
    for f in FUELS:
        out[f"{f}_MW"] = pd.to_numeric(by_fuel[f], errors="coerce").fillna(0.0) if f in by_fuel else 0.0
    out["Total_MW"] = out[[f"{f}_MW" for f in FUELS]].sum(axis=1)
    return out.round(1).sort_index()


def load_monthly(path):
    """The saved Monthly sheet (date index), or an empty frame."""
    try:
        df = pd.read_excel(path, sheet_name="Monthly", index_col=0)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index, errors="coerce")
    df = df[df.index.notna()]
    df.index.name = "date"
    return df


def load_sheet(path, sheet, index_col=None):
    try:
        return pd.read_excel(path, sheet_name=sheet, index_col=index_col)
    except (FileNotFoundError, ValueError, KeyError, OSError):
        return pd.DataFrame()


def ember_check(country, monthly):
    """Latest Ember yearly capacity (GW) by fuel vs this pull's December (or latest) value of the same year.
    Returns a DataFrame for a 'Ember check' sheet; empty if Ember can't be reached."""
    try:
        r = requests.get(EMBER_URL, timeout=(15, 300), headers={"User-Agent": "gas-demand-scripts/1.0"})
        r.raise_for_status()
        e = pd.read_csv(io.BytesIO(r.content), low_memory=False)
    except Exception as exc:  # noqa: BLE001
        print(f"  Ember check skipped: {type(exc).__name__}: {exc}", flush=True)
        return pd.DataFrame()
    e = e[(e["Area"] == country) & (e["Category"] == "Capacity") & (e["Unit"] == "GW")]
    if e.empty or monthly.empty:
        return pd.DataFrame()
    rows = []
    for year in sorted(e["Year"].unique())[-3:]:
        sub = monthly[monthly.index.year == year]
        if sub.empty:
            continue
        ours = sub.iloc[-1]
        ey = e[e["Year"] == year]
        emb = {}
        for _, x in ey.iterrows():
            f = EMBER_TO_FUEL.get(x["Variable"])
            if f and x["Subcategory"] == "Fuel":
                emb[f] = emb.get(f, 0.0) + float(x["Value"])
        tot = ey[ey["Variable"].isin(["Total Generation", "Total"]) | (ey["Subcategory"] == "Total")]["Value"]
        emb["Total"] = float(tot.iloc[0]) if len(tot) else sum(emb.values())
        for f in FUELS + ["Total"]:
            g = float(ours[f"{f}_MW"]) / 1000.0
            eg = emb.get(f)
            rows.append({"year": int(year), "compared_month": sub.index[-1].strftime("%Y-%m"), "fuel": f,
                         "this_pull_GW": round(g, 3), "ember_GW": None if eg is None else round(eg, 3),
                         "diff_pct": None if not eg else round(100 * (g - eg) / eg, 1)})
    out = pd.DataFrame(rows)
    if not out.empty:
        print("Ember check (latest year):", flush=True)
        print(out[out["year"] == out["year"].max()].to_string(index=False), flush=True)
    return out


def write(out_path, monthly, extra_sheets, notes, titles):
    """Monthly first, then extra sheets; notes tab via xlsx_notes (atomic write)."""
    sheets = {"Monthly": monthly}
    sheets.update({k: v for k, v in extra_sheets.items() if v is not None and not v.empty})
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    xlsx_notes.write_workbook(out_path, sheets, notes, titles)
    last = monthly.iloc[-1]
    print(f"Saved {out_path}: {len(monthly)} rows {monthly.index.min():%Y-%m}..{monthly.index.max():%Y-%m}; latest "
          + ", ".join(f"{f} {last[f + '_MW'] / 1000:.2f}" for f in FUELS) + f", Total {last['Total_MW'] / 1000:.2f} GW",
          flush=True)
