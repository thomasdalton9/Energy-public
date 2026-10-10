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
    # the shared chart palette holds 8 series: oil and geothermal (about 2 GW together) are folded into 'Other' here; the
    # 'Daily' sheet keeps them separate
    g["Other"] = g[[c for c in ("Oil", "Geothermal", "Other") if c in g.columns]].sum(axis=1, min_count=1)
    g = g[[f for f in FUEL_ORDER if f in g.columns and f not in ("Oil", "Geothermal")]]
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


def lng_stock_weekly(path):
    """Weekly power-utility LNG stock (ANRE), thousand tonnes."""
    d = add_charts.by_date(add_charts.read(path, "Weekly"), "date")
    return [add_charts.spec("LNG stock", d[["Stock_kt"]].rename(columns={"Stock_kt": "LNG stock"}).dropna(),
                            "Japan power-company LNG stock (ANRE, weekly)", "kt", "line", "%b/%y")]


CAP_GROUPS = {"Hydro": ["Hydro_kW"], "Pumped hydro": ["Pumped hydro_kW"], "Gas": ["LNG_kW", "LPG_kW", "Other gas_kW"],
              "Wind": ["Wind_kW"], "Solar": ["Solar_kW"], "Coal": ["Coal_kW"], "Nuclear": ["Nuclear_kW"],
              "Oil": ["Oil_kW"], "Geothermal": ["Geothermal_kW"], "Battery": ["Battery_kW"],
              "Other": ["Other thermal_kW", "Bituminous mixture_kW", "Other_kW"]}


def meti_stats(path):
    """METI/ANRE electric power statistics: installed capacity (GW) and LNG at power stations (kt)."""
    out = []
    cap = add_charts.read(path, "Capacity")
    g = pd.DataFrame()
    if "month" in cap.columns and len(cap):
        c = add_charts.by_date(cap, "month")
        g = pd.DataFrame({k: c[[x for x in v if x in c.columns]].sum(axis=1, min_count=1) for k, v in CAP_GROUPS.items()}) / 1e6
        g = g.dropna(how="all", axis=1)
    if len(g):
        out.append(add_charts.spec("Capacity", g, "Japan installed capacity by type (METI electric power statistics)",
                                   "GW installed", "stacked_bar"))
    f = add_charts.by_date(add_charts.read(path, "Fuel"), "month")
    if "LNG_stock_t" in f and f["LNG_stock_t"].notna().sum() > 12:
        stock = (f["LNG_stock_t"].dropna() / 1000.0)
        stock.index = stock.index + pd.offsets.MonthEnd(0)   # month-end stock dated on its last day
        out.append({"name": "LNG stock", "water_year": stock.resample("D").interpolate(limit=31, limit_area="inside"),
                    "title": "Japan LNG stock at power stations, month-end (METI)", "units": "kt"})
    cons = pd.DataFrame({"Receipts": f.get("LNG_receipts_t"), "Consumption": f.get("LNG_consumption_t")}) / 1000.0
    if cons.notna().any().any():
        out.append(add_charts.spec("LNG use", cons.dropna(how="all"),
                                   "Japan LNG receipts and consumption at power stations (METI)", "kt per month", "line"))
    return out
