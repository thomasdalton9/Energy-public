"""Chart specs for the Japan workbooks (used by add_charts.REGISTRY and japan/JAPAN_MASTER.py).

Each function takes the workbook path and returns a list of add_charts.spec(...) dicts."""
import os
import sys

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import add_charts  # noqa: E402

FUEL_ORDER = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Oil", "Bioenergy", "Geothermal", "Other"]
AREAS = ["Hokkaido", "Tohoku", "Tokyo", "Chubu", "Hokuriku", "Kansai", "Chugoku", "Shikoku", "Kyushu"]


def _monthly_gwh(path, cols, start="2021-04-01"):
    d = add_charts.by_date(add_charts.read(path, "Daily"), "date")
    d = d[[c for c in cols if c in d.columns]].apply(pd.to_numeric, errors="coerce")
    m = add_charts.complete_months(d, d.resample("MS").sum(min_count=1) / 1000.0)
    last = d.dropna(how="all").index.max()
    if last < last + pd.offsets.MonthEnd(0):   # month in progress
        m = m[m.index < last.to_period("M").to_timestamp()]
    return m[m.index >= start]


def power(path):
    g = _monthly_gwh(path, [f"{f}_MWh" for f in FUEL_ORDER]).rename(columns=lambda c: c.replace("_MWh", ""))
    g = g[[f for f in FUEL_ORDER if f in g.columns]]
    out = [add_charts.spec("Generation", g, "Japan power generation by source (TSO area data)", "GWh per month",
                           "stacked_bar")]
    dem = _monthly_gwh(path, ["Demand_MWh"]).rename(columns={"Demand_MWh": "Area demand"})
    if dem["Area demand"].notna().any():
        out.append(add_charts.spec("Demand", dem, "Japan power demand (TSO area data)", "GWh per month", "line"))
    cur = _monthly_gwh(path, ["Solar_Curtailed_MWh", "Wind_Curtailed_MWh"]).rename(
        columns={"Solar_Curtailed_MWh": "Solar output control", "Wind_Curtailed_MWh": "Wind output control"})
    if len(cur.dropna(how="all")):
        out.append(add_charts.spec("Output control", cur, "Japan solar and wind output control (TSO area data)",
                                   "GWh per month", "stacked_bar"))
    return out


def jepx(path):
    d = add_charts.by_date(add_charts.read(path, "Daily"), "date")
    cols = ["System_JPY_per_kWh"] + [f"{a}_JPY_per_kWh" for a in AREAS]
    m = d[[c for c in cols if c in d.columns]].resample("MS").mean()
    last = d.index.max()
    if last < last + pd.offsets.MonthEnd(0):
        m = m[m.index < last.to_period("M").to_timestamp()]
    m = m.rename(columns=lambda c: c.replace("_JPY_per_kWh", ""))
    return [add_charts.spec("Spot prices", m.dropna(how="all"), "Japan JEPX day-ahead spot price by area (monthly average)",
                            "JPY per kWh", "line")]


def fx(path):
    d = add_charts.by_date(add_charts.read(path, "Data"), "date")
    return [add_charts.spec("Daily", add_charts.daily(d[["JPY_per_USD"]].rename(
        columns={"JPY_per_USD": "Yen per US$ (Fed H.10)"}).dropna(), "2021-01-01"),
        "Japan yen per US dollar (Federal Reserve H.10, daily)", "JPY per US$", "line", "%b/%y")]


def mof_imports(path):
    """MOF monthly press release: LNG, coal and LPG in Mt (thousand tonnes / 1000), crude oil in million kl."""
    d = add_charts.by_date(add_charts.read(path, "Monthly"), "month")
    out = []
    if "LNG_kt" in d:
        out.append(add_charts.spec("LNG imports", (d[["LNG_kt"]] / 1000.0).rename(columns={"LNG_kt": "LNG"}).dropna(),
                                   "Japan LNG imports (Ministry of Finance customs statistics)", "Mt per month", "line"))
    fuels = {"Coal_kt": "Coal", "Steam_coal_kt": "Steam coal", "LPG_kt": "LPG"}
    cols = {k: v for k, v in fuels.items() if k in d and d[k].notna().any()}
    if cols:
        out.append(add_charts.spec("Coal and LPG imports", (d[list(cols)] / 1000.0).rename(columns=cols).dropna(how="all"),
                                   "Japan coal and LPG imports (Ministry of Finance customs statistics)", "Mt per month", "line"))
    if "Crude_oil_thousand_kl" in d:
        out.append(add_charts.spec("Crude oil imports", (d[["Crude_oil_thousand_kl"]] / 1000.0).rename(
            columns={"Crude_oil_thousand_kl": "Crude oil"}).dropna(),
            "Japan crude oil imports (Ministry of Finance customs statistics)", "million kl per month", "line"))
    return out


def lng_stock(path):
    """Weekly power-utility LNG stock (ANRE) as an Oct-Sep water-year chart, kt."""
    d = add_charts.by_date(add_charts.read(path, "Weekly"), "date")
    s = (d["Stock_kt"].dropna()).rename("LNG stock")
    return [{"name": "LNG stock", "title": "Japan power-utility LNG stock (ANRE, weekly)", "units": "kt",
             "water_year": s, "df": s.to_frame(), "kind": "line", "date_format": "%Y-%m", "line_cols": ()}] if len(s) > 60 else \
        [add_charts.spec("LNG stock", s.to_frame(), "Japan power-utility LNG stock (ANRE, weekly)", "kt", "line", "%b/%y")]
