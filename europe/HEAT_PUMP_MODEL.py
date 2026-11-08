"""
Heat pump demand model for Europe, country by country, as a LIVE EXCEL WORKBOOK (owner's choice, Nov 2026: the model
is run and changed in Excel; this script only fills the inputs and writes the formulas).

Output: output/Data and Chart Outputs/heat_pump_demand_model.xlsx

What the workbook does (all forecast numbers are Excel formulas on the yellow input cells):
  heat delivered  = installed heat pump capacity (MW) x full-load hours x monthly share of the year's heat
  electricity     = heat / COP   (air source: monthly COP from outdoor temperature, scaled so the year averages the
                    country's seasonal performance factor; ground and water source: the SPF, flat)
  gas avoided     = heat x gas share of household space heating / gas boiler efficiency
  monthly share   = hot-water share / 12 + (1 - hot-water share) x HDD of the month ('Degree days') / normal HDD of the year
  forecast        = monthly, Jan 2025 - Dec 2030, at the degree days on 'Degree days' (actual months, then 1991-2020
                    normals; paste your own), heat pump stock on an S-curve of household penetration (or linear MW/yr).

Inputs (each cell carries a status on 'Input status': SOURCED / DERIVED / ASSUMPTION):
  Eurostat nrg_inf_hptc  installed thermal capacity (MW), seasonal performance factor, time usage (h) by country and
                         technology (aerothermal = air, geothermal = ground, hydrothermal = water), 1990-2024 (UK to 2019)
  Eurostat nrg_d_hhq     household final energy: ambient heat for space vs water heating (hot-water share), natural gas
                         share of household space heating
  When2Heat (Ruhnau, Hirth, Praktiknjo 2019; github.com/oruhnau/when2heat input/cop/cop_parameters.csv, scripts/cop.py)
                         COP = (c0 + c1 dT + c2 dT^2) x 0.85, dT = sink - source; radiator sink 40 - T, floor 30 - 0.5 T
  Repo weather           europe_degree_days.xlsx 'Normals month' (HDD 18C, 1991-2020) and the NASA POWER city store
                         (monthly normal temperature, population-weighted like the degree days)

Run: python3 europe/HEAT_PUMP_MODEL.py [--out PATH]   (manual workflow heat_pump_model.yml; Europe is not scheduled)
"""

import argparse
import datetime as dt
import io
import os
import sys

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_charts
import xlsx_charts

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTDIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
OUT = os.path.join(OUTDIR, "heat_pump_demand_model.xlsx")
DEGREE_DAYS = os.path.join(OUTDIR, "europe_degree_days.xlsx")

EUROSTAT = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/"
COP_PARAMS_URL = "https://raw.githubusercontent.com/oruhnau/when2heat/master/input/cop/cop_parameters.csv"
H = {"User-Agent": "Mozilla/5.0 (energy data repo)"}

# Eurostat geo code -> (name, weather code in europe_degree_days.xlsx)
COUNTRIES = {
    "AT": ("Austria", "AT"), "BE": ("Belgium", "BE"), "BG": ("Bulgaria", "BG"), "HR": ("Croatia", "HR"),
    "CY": ("Cyprus", "CY"), "CZ": ("Czechia", "CZ"), "DK": ("Denmark", "DK"), "EE": ("Estonia", "EE"),
    "FI": ("Finland", "FI"), "FR": ("France", "FR"), "DE": ("Germany", "DE"), "EL": ("Greece", "GR"),
    "HU": ("Hungary", "HU"), "IE": ("Ireland", "IE"), "IT": ("Italy", "IT"), "LV": ("Latvia", "LV"),
    "LT": ("Lithuania", "LT"), "LU": ("Luxembourg", "LU"), "MT": ("Malta", "MT"), "NL": ("Netherlands", "NL"),
    "PL": ("Poland", "PL"), "PT": ("Portugal", "PT"), "RO": ("Romania", "RO"), "SK": ("Slovakia", "SK"),
    "SI": ("Slovenia", "SI"), "ES": ("Spain", "ES"), "SE": ("Sweden", "SE"), "UK": ("United Kingdom", "GB"),
    "NO": ("Norway", "NO"),
}
ISO3 = {"AT": "AUT", "BE": "BEL", "BG": "BGR", "HR": "HRV", "CY": "CYP", "CZ": "CZE", "DK": "DNK", "EE": "EST",
        "FI": "FIN", "FR": "FRA", "DE": "DEU", "EL": "GRC", "HU": "HUN", "IE": "IRL", "IT": "ITA", "LV": "LVA",
        "LT": "LTU", "LU": "LUX", "MT": "MLT", "NL": "NLD", "PL": "POL", "PT": "PRT", "RO": "ROU", "SK": "SVK",
        "SI": "SVN", "ES": "ESP", "SE": "SWE", "UK": "GBR", "NO": "NOR"}
MACRO = os.path.join(OUTDIR, "macro_drivers.xlsx")
POP_YEARS = list(range(2015, 2032))
SMAX_DEFAULT = 0.70        # ASSUMPTION: saturation of the heat-equivalent household penetration
LINEAR_ABOVE = 0.60        # penetration this high already means the metric is not meaningful -> Linear method
K_BOUNDS = (0.05, 0.35)   # ASSUMPTION: ceiling on the fitted speed (short histories start with fast early growth)
TECHS = [("ATH", "Air"), ("GTH", "Ground"), ("HTH", "Water")]
FIRST_HIST = 2015
FC_START, FC_END = dt.date(2025, 1, 1), dt.date(2030, 12, 1)
GROWTH_YEARS = 3

YELLOW = PatternFill("solid", fgColor="FFF2CC")
GREY = PatternFill("solid", fgColor="EDEDED")
BOLD = Font(bold=True)


def out(*a):
    print(*a, flush=True)


# ---------------------------------------------------------------- data
def eurostat(dataset, **params):
    q = [("format", "JSON"), ("lang", "EN")]
    for k, v in params.items():
        for x in (v if isinstance(v, (list, tuple)) else [v]):
            q.append((k, x))
    r = requests.get(EUROSTAT + dataset, params=q, headers=H, timeout=180)
    r.raise_for_status()
    d = r.json()
    dims, size = d["id"], d["size"]
    inv = {k: {i: c for c, i in d["dimension"][k]["category"]["index"].items()} for k in dims}
    rows = []
    for n, v in d["value"].items():
        n, rec = int(n), {}
        for k, s in reversed(list(zip(dims, size))):
            rec[k] = inv[k][n % s]
            n //= s
        rec["value"] = v
        rows.append(rec)
    out(f"Eurostat {dataset}: {len(rows)} values (updated {d.get('updated')})")
    return pd.DataFrame(rows), d.get("updated")


GW_SCALED = set()   # countries whose capacity Eurostat publishes 1,000 times too small (filled in heat_pump_stock)


def heat_pump_stock():
    df, upd = eurostat("nrg_inf_hptc", hp_tech=[t for t, _ in TECHS] + ["ATH_ABVTH"], geo=list(COUNTRIES),
                       sinceTimePeriod=str(FIRST_HIST))
    df["time"] = df["time"].astype(int)
    piv = df.pivot_table(index=["geo", "hp_tech", "time"], columns="plant_tec", values="value", aggfunc="first")
    # Finland: the total rows are about 1/1000 of its own 'SPF above the threshold' sub-rows (4.5 vs 4,485 in 2022),
    # i.e. GW figures in the MW column; such a country's capacity is multiplied by 1,000 (status says so)
    for geo in COUNTRIES:
        try:
            tot = piv.loc[(geo, "ATH")]["CAP_HEAT"]
            abv = piv.loc[(geo, "ATH_ABVTH")]["CAP_HEAT"]
        except KeyError:
            continue
        ratio = (abv / tot.replace(0, float("nan"))).dropna()
        if len(ratio) and ratio.between(900, 1100).any():
            GW_SCALED.add(geo)
            idx = piv.index.get_level_values(0) == geo
            piv.loc[idx, "CAP_HEAT"] = piv.loc[idx, "CAP_HEAT"] * 1000
            out(f"{geo}: capacity published in GW-sized numbers; multiplied by 1,000")
    piv = piv[piv.index.get_level_values(1) != "ATH_ABVTH"]
    return piv, upd


def household_split():
    df, upd = eurostat("nrg_d_hhq", unit="TJ", siec=["RA600", "G3000", "TOTAL"],
                       nrg_bal=["FC_OTH_HH_E_SH", "FC_OTH_HH_E_WH"], geo=list(COUNTRIES), sinceTimePeriod="2015")
    df["time"] = df["time"].astype(int)
    return df.pivot_table(index=["geo", "time"], columns=["siec", "nrg_bal"], values="value", aggfunc="first"), upd


def cop_parameters():
    r = requests.get(COP_PARAMS_URL, headers=H, timeout=60)
    r.raise_for_status()
    p = pd.read_csv(io.StringIO(r.content.decode("utf-8-sig")), sep=";", decimal=",").set_index("Coefficient")
    return p


def weather_normals():
    """Normal (1991-2020) monthly HDD and CDD 18C and the actual monthly series, from europe_degree_days.xlsx."""
    norm = pd.read_excel(DEGREE_DAYS, sheet_name="Normals month").set_index("month")
    act = pd.read_excel(DEGREE_DAYS, sheet_name="Monthly")
    act["month"] = pd.to_datetime(act["month"]).dt.date
    act = act.set_index("month")
    out(f"Weather: normals and actual monthly degree days to {max(act.index):%b %Y}")
    return norm, act


def population():
    pop = pd.read_excel(MACRO, sheet_name="Population_IMF_m").set_index("year")
    return pd.DataFrame({geo: pop[iso] for geo, iso in ISO3.items()}).loc[POP_YEARS]


def household_size():
    df, _ = eurostat("ilc_lvph01", geo=list(COUNTRIES), sinceTimePeriod="2015")
    df["time"] = df["time"].astype(int)
    res = {}
    for geo, g in df.dropna(subset=["value"]).groupby("geo"):
        r = g.sort_values("time").iloc[-1]
        res[geo] = (float(r["value"]), int(r["time"]))
    return res


def household_heat():
    """Household space + water heating, all fuels incl. ambient heat and electricity, TWh by country and year."""
    df, _ = eurostat("nrg_d_hhq", unit="TJ", siec="TOTAL", nrg_bal=["FC_OTH_HH_E_SH", "FC_OTH_HH_E_WH"],
                     geo=list(COUNTRIES), sinceTimePeriod="2015")
    df["time"] = df["time"].astype(int)
    return df.groupby(["geo", "time"])["value"].sum() / 3600.0


def fit_scurve(shares, smax):
    """Speed k of a logistic with fixed saturation, by least squares on logit(s/smax) over the reported years."""
    import numpy as np
    s = shares[(shares > 0) & (shares < smax)]
    if len(s) < 3:
        return None, len(s)
    y = np.log(s / (smax - s))
    k = np.polyfit(np.array(s.index, dtype=float), y.values, 1)[0]
    return float(min(max(k, K_BOUNDS[0]), K_BOUNDS[1])), len(s)


def add_penetration(inp, sts, stock, pop, hhsize, hheat):
    """Household size, heat per household, base-year penetration, S-curve defaults; returns the history table."""
    hist = {}
    fitted = {}
    for geo in inp.index:
        size, ysz = hhsize.get(geo, (None, None))
        inp.loc[geo, "hhsize"] = size
        sts.loc[geo, "hhsize"] = f"SOURCED Eurostat ilc_lvph01 {ysz} (held flat)" if size else None
        base = int(inp.loc[geo, "base"])
        years = sorted(y for (g, y) in hheat.index if g == geo and hheat[(g, y)] > 0)
        yh = max([y for y in years if y <= base], default=max(years) if years else None)
        if yh and size:
            households = pop.loc[yh, geo] * 1e6 / size
            inp.loc[geo, "hhheat"] = hheat[(geo, yh)] * 1e6 / households
            sts.loc[geo, "hhheat"] = (f"DERIVED: Eurostat nrg_d_hhq {yh} household space + water heating (all fuels) / "
                                      "households (IMF population / household size); held flat")
        # penetration history = heat-pump heat / household heating need (hours filled where not reported)
        sh = {}
        for y in range(FIRST_HIST, base + 1):
            heat = 0.0
            for t, _ in TECHS:
                if (geo, t, y) in stock.index:
                    r = stock.loc[(geo, t, y)]
                    c = float(r.get("CAP_HEAT") or 0)
                    h = float(r.get("TIME_USE") or 0)
                    if not (SANE["TIME_USE"][0] <= h <= SANE["TIME_USE"][1]):
                        h = float(inp.loc[geo, f"{t}_hours"])
                    heat += c * h / 1e6
            need = hheat.get((geo, y))
            if need and need > 0:
                sh[y] = heat / need
        hist[geo] = pd.Series(sh, dtype=float)
        k, npts = fit_scurve(hist[geo], SMAX_DEFAULT)
        fitted[geo] = (k, npts)
    ks = [k for k, _ in fitted.values() if k is not None]
    k_med = float(pd.Series(ks).median()) if ks else 0.2
    for geo in inp.index:
        k, npts = fitted[geo]
        inp.loc[geo, "smax"] = SMAX_DEFAULT
        sts.loc[geo, "smax"] = ("ASSUMPTION: saturation of heat-pump heat / household heating need; Norway and Sweden "
                                "are at about 0.45 and still rising")
        if k is not None:
            inp.loc[geo, "k"] = round(k, 3)
            sts.loc[geo, "k"] = (f"DERIVED: least-squares fit of the logistic (saturation {SMAX_DEFAULT}) to {npts} reported "
                                 f"years of penetration, bounded {K_BOUNDS[0]}-{K_BOUNDS[1]}")
        else:
            inp.loc[geo, "k"] = round(k_med, 3)
            sts.loc[geo, "k"] = "ASSUMPTION: under 3 reported years to fit; median speed of the fitted countries"
        last = hist[geo].dropna()
        lastv = float(last.iloc[-1]) if len(last) else 0.0
        if lastv >= LINEAR_ABOVE:
            inp.loc[geo, "method"] = "Linear"
            sts.loc[geo, "method"] = (f"ASSUMPTION: heat-pump heat is already {lastv:.0%} of household heating need "
                                      "(Eurostat's heat pumps include services and reversible air conditioners), so the "
                                      "penetration metric is not meaningful; linear MW growth used")
        else:
            inp.loc[geo, "method"] = "S-curve"
            sts.loc[geo, "method"] = "Owner's choice (Nov 2026): S-curve of household penetration"
    histdf = pd.DataFrame(hist).T.reindex(columns=range(FIRST_HIST, 2025))
    return histdf


# ---------------------------------------------------------------- inputs per country
SANE = {"TIME_USE": (1, 8760), "SPERF_FACT": (1.5, 7)}   # outside these a published value is treated as not reported


def latest(series, col):
    lo, hi = SANE[col]
    s = series.dropna()
    s = s[(s >= lo) & (s <= hi)]
    return (s.index.max(), s.loc[s.index.max()]) if len(s) else (None, None)


def build_inputs(stock, hh):
    rows, status = [], []
    for geo, (name, wc) in COUNTRIES.items():
        caps = {t: stock.loc[(geo, t)]["CAP_HEAT"] if (geo, t) in stock.index.droplevel(2) else pd.Series(dtype=float)
                for t, _ in TECHS}
        base = max((y for t in caps for y in caps[t].dropna().index if caps[t][y] > 0), default=None)
        rec, st = {"geo": geo, "name": name, "wc": wc, "base": base}, {}
        for t, _ in TECHS:
            c = caps[t]
            cap = float(c.get(base, 0) or 0)
            rec[f"{t}_cap"] = cap
            st[f"{t}_cap"] = f"SOURCED Eurostat nrg_inf_hptc {base}" + (
                " (published 1,000 times too small - GW figures in the MW column - multiplied by 1,000: DERIVED)"
                if geo in GW_SCALED else "")
            # growth from the earliest year with capacity in the GROWTH_YEARS before the base year
            y0 = min((y for y in range(base - GROWTH_YEARS, base) if (c.get(y) or 0) > 0), default=None) if base else None
            if y0 is not None and cap > 0:
                rec[f"{t}_growth"] = max((cap - float(c[y0])) / (base - y0), 0.0)
                st[f"{t}_growth"] = (f"DERIVED: mean yearly change in capacity {y0}-{base} (floored at 0); "
                                     "ASSUMPTION that it continues")
            else:
                rec[f"{t}_growth"] = 0.0
                st[f"{t}_growth"] = "ASSUMPTION: no earlier year reported, so no growth - set your own"
            sub = stock.loc[(geo, t)] if (geo, t) in stock.index.droplevel(2) else pd.DataFrame()
            for k, col in (("hours", "TIME_USE"), ("spf", "SPERF_FACT")):
                y, v = latest(sub[col], col) if col in sub else (None, None)
                rec[f"{t}_{k}"] = v
                st[f"{t}_{k}"] = f"SOURCED Eurostat nrg_inf_hptc {y}" if v else None
        # household split (latest year with data)
        if geo in hh.index.get_level_values(0):
            g = hh.loc[geo]
            def col(s, b):
                return g[(s, b)] if (s, b) in g.columns else pd.Series(dtype=float)
            amb = (col("RA600", "FC_OTH_HH_E_SH") + col("RA600", "FC_OTH_HH_E_WH")).dropna()
            wh = col("RA600", "FC_OTH_HH_E_WH")
            ok = [y for y in amb.index if amb[y] > 0 and (wh.get(y) or 0) > 0]
            if ok:
                y = max(ok)
                rec["hw"], st["hw"] = wh[y] / amb[y], f"DERIVED Eurostat nrg_d_hhq {y}: ambient heat for water / (space + water)"
            tot = col("TOTAL", "FC_OTH_HH_E_SH")
            gas = col("G3000", "FC_OTH_HH_E_SH")
            ok = [y for y in tot.index if (tot[y] or 0) > 0 and pd.notna(gas.get(y))]
            if ok:
                y = max(ok)
                rec["gas"], st["gas"] = gas[y] / tot[y], f"DERIVED Eurostat nrg_d_hhq {y}: natural gas / all fuels, household space heating"
        rows.append(rec)
        status.append(st)
    inp = pd.DataFrame(rows).set_index("geo")
    sts = pd.DataFrame(status, index=inp.index)
    # fills: capacity-weighted average of the countries that report (labelled ASSUMPTION)
    for t, _ in TECHS:
        for k in ("hours", "spf"):
            have = inp[f"{t}_{k}"].notna() & (inp[f"{t}_cap"] > 0)
            avg = (inp.loc[have, f"{t}_{k}"] * inp.loc[have, f"{t}_cap"]).sum() / inp.loc[have, f"{t}_cap"].sum()
            miss = inp[f"{t}_{k}"].isna()
            inp.loc[miss, f"{t}_{k}"] = round(avg, 3)
            sts.loc[miss, f"{t}_{k}"] = ("ASSUMPTION: not reported (or outside a sane range); capacity-weighted average "
                                         "of the reporting countries")
    have = inp["hw"].notna() if "hw" in inp else pd.Series(False, index=inp.index)
    hw_avg = inp.loc[have, "hw"].mean()
    inp.loc[~have, "hw"] = round(hw_avg, 3)
    sts.loc[~have, "hw"] = "ASSUMPTION: water heating not split out; mean of the countries that report it"
    if "gas" not in inp:
        inp["gas"] = None
    miss = inp["gas"].isna()
    inp.loc[miss, "gas"] = round(inp.loc[~miss, "gas"].mean(), 3)
    sts.loc[miss, "gas"] = "ASSUMPTION: not in nrg_d_hhq; mean of the other countries"
    return inp, sts


def history_table(stock):
    rows = []
    for geo in COUNTRIES:
        for y in range(FIRST_HIST, 2025):
            r = {"geo": geo, "year": y}
            heat = elec = cap = 0.0
            for t, _ in TECHS:
                if (geo, t, y) in stock.index:
                    s = stock.loc[(geo, t, y)]
                    c, h, f = (float(s.get(k) or 0) for k in ("CAP_HEAT", "TIME_USE", "SPERF_FACT"))
                    cap += c
                    heat += c * h / 1e6
                    elec += c * h / f / 1e6 if f >= SANE["SPERF_FACT"][0] else 0
            r.update(cap_GW=cap / 1e3, heat_TWh=heat, elec_TWh=elec)
            rows.append(r)
    return pd.DataFrame(rows)


# ---------------------------------------------------------------- workbook
def months():
    d, res = FC_START, []
    while d <= FC_END:
        res.append(d)
        d = dt.date(d.year + (d.month == 12), d.month % 12 + 1, 1)
    return res


def header(ws, row, values, col=1):
    for i, v in enumerate(values):
        c = ws.cell(row, col + i, v)
        c.font = BOLD
        c.alignment = Alignment(wrap_text=True, vertical="top")


def cop_formula(t):
    """Excel expression: air-source COP (When2Heat curve x correction) at temperature cell t, radiator/floor mix."""
    d_rad, d_flo = f"((40-{t})-{t})", f"((30-0.5*{t})-{t})"
    cop_rad = f"cop_corr*(cop_a0+cop_a1*{d_rad}+cop_a2*{d_rad}^2)"
    cop_flo = f"cop_corr*(cop_a0+cop_a1*{d_flo}+cop_a2*{d_flo}^2)"
    return f"1/(rad_share/{cop_rad}+(1-rad_share)/{cop_flo})"


def line_chart(title, y_title, ws_data, cat_col, val_cols, first_row, last_row, header_row):
    ch = LineChart()
    ch.title, ch.y_axis.title = title, y_title
    ch.height, ch.width = 9, 18
    for c in val_cols:
        ch.add_data(Reference(ws_data, min_col=c, min_row=header_row, max_row=last_row), titles_from_data=True)
    ch.set_categories(Reference(ws_data, min_col=cat_col, min_row=first_row, max_row=last_row))
    ch.x_axis.number_format = "mmm/yy"
    for i, s in enumerate(ch.series):
        s.smooth = False
        s.graphicalProperties.line.width = 22000
    xlsx_charts.tidy_layout(ch)
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))          # no chart border
    ch.plot_area.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))  # no plot-area border
    return ch


def build(path, inp, sts, hist, cop, norm, act, upd, pop, penhist):
    wb = Workbook()
    codes = list(inp.index)
    n = len(codes)
    mlist = months()

    # ---- Notes
    ws = wb.active
    ws.title = "Notes"
    notes = [
        "Heat pump demand model - Europe, country by country (EU27 + United Kingdom + Norway)",
        "",
        "HOW TO USE",
        "Change any YELLOW cell on 'Global' or 'Countries'; every forecast figure is an Excel formula and updates at once.",
        "Pick a country on 'Dashboard' (cell B2) to see its monthly electricity demand and gas avoided.",
        "Forecast: monthly, Jan 2025 - Dec 2030, at the weather on 'Degree days' (actual months to the latest NASA POWER month, then 1991-2020 normals - paste your own HDD/CDD there). History (annual, Eurostat) on 'History'.",
        "",
        "METHOD (per country and month)",
        "Growth method per country ('Countries', yellow): S-CURVE (default) - household penetration = saturation / "
        "(1 + e^(-speed x (year - midpoint))), penetration being heat-pump heat / household heating need; heat-pump heat "
        "= penetration x households (population / household size, 'Population' sheet - paste your own) x heating need "
        "per household; each source keeps its last-year share of heat and capacity = heat / full-load hours. The midpoint "
        "is solved so the curve passes through the last Eurostat year. LINEAR - capacity = last Eurostat year + yearly "
        "addition x years since then (used where penetration is already above 60%: Portugal, Malta).",
        "Heat (TWh) = capacity x full-load hours x monthly share; monthly share = hot-water share / 12 + (1 - hot-water share) "
        "x HDD of the month ('Degree days') / normal HDD of the year, so a cold month gives more heat.",
        "Electricity (TWh) = heat / COP. Air source: COP of the month from its mean temperature (18 - (HDD - CDD) / days) with the When2Heat "
        "curve (radiator and underfloor sinks mixed by 'Global'), scaled so the heat-weighted year equals the Eurostat SPF. "
        "Ground and water source: the Eurostat SPF, flat. Shown as average GW (TWh / hours in the month).",
        "Gas avoided = heat x gas share of household space heating / gas boiler efficiency; Bcf/d via GWh per mcm and "
        "0.0353147 Bcf per mcm. It is the gas a gas boiler would have burned for the same heat, NOT a measured change.",
        "",
        "STATUS OF INPUTS ('Input status' sheet, one line per cell): SOURCED = as published; DERIVED = calculated from "
        "published figures; ASSUMPTION = our choice where nothing is published - change it if you have better.",
        "",
        "SIMPLIFICATIONS",
        "Monthly mean temperatures, not hourly: the curve is used for the monthly SHAPE only; the level is set by the "
        "Eurostat SPF. Hot water uses the same COP shape as space heating. Air-air units (incl. reversible air "
        "conditioners in Italy and Spain) are included at Eurostat's own full-load hours, which count heating use only.",
        "United Kingdom: Eurostat stops at 2019, so 2020-2024 are grown from 2019 at its 2016-2019 yearly addition "
        "(ASSUMPTION; MCS installation data would be better). France reports capacity but no hours or SPF.",
        "No back-test (owner rule). Policy targets (e.g. REPowerEU) are not in the forecast.",
        "",
        "SOURCES",
        f"Eurostat nrg_inf_hptc - Heat pumps, technical characteristics by technologies (updated {upd['hptc']}): "
        "https://ec.europa.eu/eurostat/databrowser/view/nrg_inf_hptc/default/table",
        f"Eurostat nrg_d_hhq - Disaggregated final energy consumption in households (updated {upd['hhq']}): "
        "https://ec.europa.eu/eurostat/databrowser/view/nrg_d_hhq/default/table",
        "Eurostat ilc_lvph01 - Average household size: https://ec.europa.eu/eurostat/databrowser/view/ilc_lvph01/default/table",
        "IMF World Economic Outlook population incl. projections (macro_drivers.xlsx 'Population_IMF_m').",
        "When2Heat COP parameters (Ruhnau, Hirth and Praktiknjo 2019, Scientific Data 6:189): "
        f"{COP_PARAMS_URL}; sink/source temperatures and the 0.85 correction from scripts/cop.py in the same repository.",
        "Weather: europe_degree_days.xlsx (NASA POWER daily T2M at the 4 most populous cities per country, "
        "population-weighted; normals 1991-2020).",
        f"Built {dt.date.today():%d %b %Y} by europe/HEAT_PUMP_MODEL.py.",
    ]
    for i, t in enumerate(notes, 1):
        c = ws.cell(i, 1, t)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if t.isupper() or i == 1:
            c.font = BOLD
    ws.column_dimensions["A"].width = 150

    # ---- Global
    g = wb.create_sheet("Global")
    header(g, 1, ["Input", "Value", "Unit", "Status", "Note"])
    glob_rows = [
        ("cop_a0", "Air-source COP curve c0", cop.loc[0, "air"], "-", "SOURCED", "When2Heat cop_parameters.csv"),
        ("cop_a1", "Air-source COP curve c1", cop.loc[1, "air"], "per K", "SOURCED", "When2Heat cop_parameters.csv"),
        ("cop_a2", "Air-source COP curve c2", cop.loc[2, "air"], "per K2", "SOURCED", "When2Heat cop_parameters.csv"),
        ("cop_corr", "COP correction factor", 0.85, "-", "SOURCED", "When2Heat scripts/cop.py (real-world correction)"),
        ("rad_share", "Share of space heating on radiators (rest underfloor)", 0.7, "share", "ASSUMPTION",
         "Radiator sink 40 - T, underfloor 30 - 0.5 T (When2Heat). Shape only: the level is the SPF."),
        ("boiler_eff", "Gas boiler efficiency (heat out / gas in)", 0.90, "share", "ASSUMPTION",
         "Condensing boiler, seasonal; lower it for older boilers"),
        ("gwh_per_mcm", "Gas energy per million m3", 10.55, "GWh/mcm", "ASSUMPTION",
         "About 38 MJ/m3 gross; ENTSOG and national TSOs use 10.5-11.2"),
        ("bcf_per_mcm", "Bcf per million m3", 0.0353147, "Bcf/mcm", "SOURCED", "Unit conversion used across the repo"),
    ]
    for i, (nm, label, v, unit, stat, note) in enumerate(glob_rows, 2):
        g.cell(i, 1, label)
        c = g.cell(i, 2, float(v))
        if stat == "ASSUMPTION":
            c.fill = YELLOW
        g.cell(i, 3, unit), g.cell(i, 4, stat), g.cell(i, 5, note)
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"Global!$B${i}")
    for col, w in zip("ABCDE", (52, 12, 10, 12, 90)):
        g.column_dimensions[col].width = w

    # ---- Countries (inputs) and Input status
    c_ws = wb.create_sheet("Countries")
    s_ws = wb.create_sheet("Input status")
    cols = ["Code", "Country", "Weather code", "Last Eurostat year"]
    keys = []
    for t, lab in TECHS:
        cols += [f"{lab} capacity, last year (MW)", f"{lab} yearly addition (MW/yr)", f"{lab} full-load hours (h)",
                 f"{lab} SPF"]
        keys += [f"{t}_cap", f"{t}_growth", f"{t}_hours", f"{t}_spf"]
    cols += ["Hot-water share of heat", "Gas share of household space heating",
             "Growth method (S-curve / Linear)", "S-curve saturation (share)", "S-curve speed (per year)",
             "Household size (persons)", "Heating need per household (MWh/yr)"]
    keys += ["hw", "gas", "method", "smax", "k", "hhsize", "hhheat"]
    calc_cols = ["Air COP scale (calc)", "Penetration, last Eurostat year (calc)", "S-curve midpoint year (calc)"]
    header(c_ws, 1, cols + calc_cols)
    header(s_ws, 1, ["Code", "Country"] + cols[4:])
    COL = {k: 5 + i for i, k in enumerate(keys)}          # column of each input
    COL["scale"], COL["sbase"], COL["t0"] = COL["hhheat"] + 1, COL["hhheat"] + 2, COL["hhheat"] + 3
    for r, geo in enumerate(codes, 2):
        rec = inp.loc[geo]
        c_ws.cell(r, 1, geo), c_ws.cell(r, 2, rec["name"]), c_ws.cell(r, 3, rec["wc"]), c_ws.cell(r, 4, int(rec["base"]))
        s_ws.cell(r, 1, geo), s_ws.cell(r, 2, rec["name"])
        for k in keys:
            v = rec[k]
            cell = c_ws.cell(r, COL[k], v if isinstance(v, str) else (round(float(v), 4) if pd.notna(v) else 0))
            stat = sts.loc[geo, k] if pd.notna(sts.loc[geo, k]) else ""
            if k.endswith("_growth") or stat.startswith("ASSUMPTION") or k in ("hw", "gas", "method", "smax", "k"):
                cell.fill = YELLOW
            s_ws.cell(r, COL[k] - 2, stat)
    last_c = 1 + n
    mdv = DataValidation(type="list", formula1='"S-curve,Linear"', allow_blank=False)
    c_ws.add_data_validation(mdv)
    mdv.add(f"{L(COL['method'])}2:{L(COL['method'])}{last_c}")
    # Population (annual, by country): IMF WEO, million - paste your own over it
    pws = wb.create_sheet("Population")
    pws.cell(1, 1, "Population, million (SOURCED: IMF World Economic Outlook via macro_drivers.xlsx 'Population_IMF_m'; "
                   "forecast years are the IMF's projections). Paste your own figures over the yellow cells - "
                   "same countries and years.").font = BOLD
    header(pws, 2, ["Code", "Country"] + POP_YEARS)
    for r, geo in enumerate(codes, 3):
        pws.cell(r, 1, geo), pws.cell(r, 2, inp.loc[geo, "name"])
        for j, y in enumerate(POP_YEARS, 3):
            v = pop.loc[y, geo]
            pws.cell(r, j, round(float(v), 4) if pd.notna(v) else None).fill = YELLOW
    pws.column_dimensions["B"].width = 16
    pop_first_col = 3
    POP = lambda cr, year_expr: (f"INDEX(Population!${L(pop_first_col)}${cr + 1}:${L(pop_first_col + len(POP_YEARS) - 1)}"
                                 f"${cr + 1},1,{year_expr}-{POP_YEARS[0] - 1})")
    # base-year penetration and S-curve midpoint (live formulas)
    for r, geo in enumerate(codes, 2):
        cc = lambda k: f"{L(COL[k])}{r}"
        hp_heat = "+".join(f"{cc(t + '_cap')}*{cc(t + '_hours')}" for t, _ in TECHS)
        need = f"({POP(r, f'$D{r}')}*{cc('hhheat')}/{cc('hhsize')})"
        c_ws.cell(r, COL["sbase"], f"=IFERROR(({hp_heat})/1000000/{need},0)").fill = GREY
        c_ws.cell(r, COL["t0"], f"=IFERROR($D{r}+LN({cc('smax')}/{cc('sbase')}-1)/{cc('k')},$D{r})").fill = GREY
    for col in range(1, COL["t0"] + 1):
        c_ws.column_dimensions[L(col)].width = 14
    c_ws.column_dimensions["B"].width = 16
    for col in range(3, s_ws.max_column + 1):
        s_ws.column_dimensions[L(col)].width = 40
    c_ws.freeze_panes = "C2"

    # ---- Weather (normals, COP, monthly share) - one column per country, same order as 'Countries'
    w = wb.create_sheet("Weather")
    blocks = {"hdd": 1, "cdd": 16, "temp": 31, "copraw": 46, "share": 61}
    titles = {"hdd": "Normal HDD 18C (1991-2020), SOURCED europe_degree_days.xlsx 'Normals month'",
              "cdd": "Normal CDD 18C (1991-2020), SOURCED europe_degree_days.xlsx 'Normals month'",
              "temp": "Normal mean temperature (C) = 18 - (HDD - CDD) / days (exact for 18C degree days)",
              "copraw": "Air-source COP from the When2Heat curve at the normal temperature (before scaling to the SPF)",
              "share": "Share of the year's heat in a normal month (hot water flat + space heating by HDD)"}
    for b, r0 in blocks.items():
        w.cell(r0, 1, titles[b]).font = BOLD
        header(w, r0 + 1, ["Month"] + codes)
        for m in range(1, 13):
            rr = r0 + 1 + m
            w.cell(rr, 1, m)
            for j, geo in enumerate(codes, 2):
                wc = inp.loc[geo, "wc"]
                cr = j  # Countries row of this geo is j (row 2 = first country) - same index
                if b in ("hdd", "cdd"):
                    w.cell(rr, j, round(float(norm.loc[m, f"{wc} {b.upper()} 18C"]), 2))
                elif b == "temp":
                    hd, cd = f"{L(j)}{blocks['hdd'] + 1 + m}", f"{L(j)}{blocks['cdd'] + 1 + m}"
                    w.cell(rr, j, f"=18-({hd}-{cd})/DAY(EOMONTH(DATE(2001,$A{rr},1),0))")
                elif b == "copraw":
                    w.cell(rr, j, "=" + cop_formula(f"{L(j)}{blocks['temp'] + 1 + m}"))
                else:
                    h0, h1 = blocks["hdd"] + 2, blocks["hdd"] + 13
                    hw = f"Countries!${L(COL['hw'])}${cr}"
                    w.cell(rr, j, f"={hw}/12+(1-{hw})*{L(j)}{blocks['hdd'] + 1 + m}/SUM({L(j)}${h0}:{L(j)}${h1})")
    w.column_dimensions["A"].width = 8
    # air COP scale factor on 'Countries': SPF x sum(share / COPraw)
    for j, geo in enumerate(codes, 2):
        sh = f"Weather!{L(j)}{blocks['share'] + 2}:{L(j)}{blocks['share'] + 13}"
        cp = f"Weather!{L(j)}{blocks['copraw'] + 2}:{L(j)}{blocks['copraw'] + 13}"
        c_ws.cell(j, COL["scale"], f"=IFERROR({L(COL['ATH_spf'])}{j}*SUMPRODUCT({sh}/{cp}),0)").fill = GREY

    # ---- Degree days for the forecast months (yellow; actual where the repo has the month, else normal)
    dd = wb.create_sheet("Degree days")
    last_act = max(d for d in act.index if pd.notna(act.loc[d, "DE HDD 18C"]))
    dd.cell(1, 1, f"Monthly degree days by country, base 18C, population-weighted (NASA POWER, europe_degree_days.xlsx). "
                  f"Defaults: ACTUAL months to {last_act:%b %Y}, then the 1991-2020 NORMAL. Paste your own over the "
                  "yellow cells (same countries and months). HDD drives the heat of the month; HDD and CDD together give "
                  "the month's mean temperature (18 - (HDD - CDD) / days), which sets the air-source COP.").font = BOLD
    HD0, CD0 = 3, 4 + n                     # first HDD and CDD columns
    header(dd, 2, ["Month", "Basis"] + [f"{c} HDD" for c in codes] + [""] + [f"{c} CDD" for c in codes])
    ddrow = {}
    for i, d in enumerate(mlist, 3):
        ddrow[d] = i
        dd.cell(i, 1, d).number_format = "mmm/yy"
        actual = d <= last_act
        dd.cell(i, 2, "actual" if actual else "normal 1991-2020")
        for j, geo in enumerate(codes):
            wc = inp.loc[geo, "wc"]
            for c0, kind in ((HD0, "HDD"), (CD0, "CDD")):
                v = act.loc[d, f"{wc} {kind} 18C"] if actual and d in act.index else None
                if v is None or pd.isna(v):
                    v = norm.loc[d.month, f"{wc} {kind} 18C"]
                dd.cell(i, c0 + j, round(float(v), 1)).fill = YELLOW
    dd.column_dimensions["B"].width = 16
    dd.freeze_panes = "C3"

    # ---- Calc (long: one row per country and month)
    calc = wb.create_sheet("Calc")
    ch = ["Code", "Month", "Year", "Days", "Months since last Eurostat year-end",
          "Air capacity (MW)", "Ground capacity (MW)", "Water capacity (MW)", "Share of year's heat",
          "Air heat (TWh)", "Ground heat (TWh)", "Water heat (TWh)", "Heat (TWh)", "Air COP",
          "Electricity (TWh)", "Electricity (average GW)", "Heat (average GW)", "Gas avoided (TWh)",
          "Gas avoided (Bcf/d)", "Penetration (S-curve)", "Household heating need (TWh/yr)",
          "Heat-pump heat, S-curve (TWh/yr)", "HDD", "CDD", "Mean temperature (C)"]
    header(calc, 1, ch)
    rowmap = {}
    r = 2
    for j, geo in enumerate(codes, 2):
        cr = j
        C = lambda k: f"Countries!${L(COL[k])}${cr}"
        for d in mlist:
            rowmap[(geo, d)] = r
            m = d.month
            calc.cell(r, 1, geo)
            calc.cell(r, 2, d).number_format = "mmm/yy"
            calc.cell(r, 3, f"=YEAR(B{r})")
            calc.cell(r, 4, f"=DAY(EOMONTH(B{r},0))")
            calc.cell(r, 5, f"=(C{r}-Countries!$D${cr})*12+MONTH(B{r})-0.5")
            # S-curve: penetration at mid-month x household heating need = heat-pump heat (annual rate)
            calc.cell(r, 20, f"=IFERROR({C('smax')}/(1+EXP(-{C('k')}*(C{r}+(MONTH(B{r})-6.5)/12-{C('t0')}))),0)")
            calc.cell(r, 21, f"=IFERROR({POP(cr, f'C{r}')}*{C('hhheat')}/{C('hhsize')},0)")
            calc.cell(r, 22, f"=T{r}*U{r}")
            base_heat = "+".join(f"{C(t + '_cap')}*{C(t + '_hours')}" for t, _ in TECHS)
            for i, (t, _) in enumerate(TECHS):
                linear = f"MAX(0,{C(t + '_cap')}+{C(t + '_growth')}*E{r}/12)"
                # S-curve capacity: tech keeps its last-year share of heat-pump heat; capacity = heat / hours
                scurve = f"IFERROR(V{r}*1000000*{C(t + '_cap')}/({base_heat}),0)"
                calc.cell(r, 6 + i, f"=IF({C('method')}=\"Linear\",{linear},{scurve})")
            jj = j - 2                                                  # country index on 'Degree days'
            calc.cell(r, 23, f"='Degree days'!{L(HD0 + jj)}{ddrow[d]}")
            calc.cell(r, 24, f"='Degree days'!{L(CD0 + jj)}{ddrow[d]}")
            calc.cell(r, 25, f"=18-(W{r}-X{r})/D{r}")
            hw = C("hw")
            hn = f"SUM(Weather!{L(j)}${blocks['hdd'] + 2}:{L(j)}${blocks['hdd'] + 13})"
            calc.cell(r, 9, f"={hw}/12+(1-{hw})*W{r}/{hn}")
            for i, (t, _) in enumerate(TECHS):
                calc.cell(r, 10 + i, f"={L(6 + i)}{r}*{C(t + '_hours')}*I{r}/1000000")
            calc.cell(r, 13, f"=SUM(J{r}:L{r})")
            calc.cell(r, 14, f"={C('scale')}*" + cop_formula(f"Y{r}"))
            calc.cell(r, 15, f"=IFERROR(J{r}/N{r},0)+IFERROR(K{r}/{C('GTH_spf')},0)+IFERROR(L{r}/{C('HTH_spf')},0)")
            calc.cell(r, 16, f"=O{r}*1000/(24*D{r})")
            calc.cell(r, 17, f"=M{r}*1000/(24*D{r})")
            calc.cell(r, 18, f"=M{r}*{C('gas')}/boiler_eff")
            calc.cell(r, 19, f"=R{r}*1000/gwh_per_mcm*bcf_per_mcm/D{r}")
            r += 1
    last_calc = r - 1
    for col in range(1, len(ch) + 1):
        calc.column_dimensions[L(col)].width = 13
    calc.freeze_panes = "C2"

    # ---- wide output sheets (month x country) + total
    wide = {}
    for name, col, unit in (("Electricity GW", 16, "Average GW"), ("Heat TWh", 13, "TWh per month"),
                            ("Gas avoided Bcf-d", 19, "Bcf/d")):
        ws = wb.create_sheet(name)
        header(ws, 1, ["Month"] + [inp.loc[c, "name"] for c in codes] + ["Total (EU27 + UK + Norway)"])
        for i, d in enumerate(mlist, 2):
            ws.cell(i, 1, d).number_format = "mmm/yy"
            for j, geo in enumerate(codes, 2):
                ws.cell(i, j, f"=Calc!{L(col)}{rowmap[(geo, d)]}").number_format = "0.000"
            ws.cell(i, n + 2, f"=SUM(B{i}:{L(n + 1)}{i})").number_format = "0.000"
        ws.cell(len(mlist) + 3, 1, f"Unit: {unit}. Formulas from 'Calc'.")
        ws.freeze_panes = "B2"
        ws.column_dimensions["A"].width = 10
        wide[name] = ws

    # ---- Annual summary (formulas)
    a = wb.create_sheet("Annual")
    years = sorted({d.year for d in mlist})
    a.cell(1, 1, "Annual totals by country (forecast; degree days as on 'Degree days') - TWh; gas avoided in bcm").font = BOLD
    for b, (label, col, scale) in enumerate((("Electricity (TWh)", "O", 1), ("Heat (TWh)", "M", 1),
                                             ("Gas avoided (bcm)", "R", None))):
        r0 = 3 + b * (n + 4)
        a.cell(r0, 1, label).font = BOLD
        header(a, r0 + 1, ["Country"] + years)
        for i, geo in enumerate(codes):
            rr = r0 + 2 + i
            a.cell(rr, 1, inp.loc[geo, "name"])
            for k, y in enumerate(years, 2):
                f = f"SUMIFS(Calc!${col}$2:${col}${last_calc},Calc!$A$2:$A${last_calc},\"{geo}\",Calc!$C$2:$C${last_calc},{y})"
                a.cell(rr, k, f"={f}" if scale else f"={f}*1000/gwh_per_mcm/1000").number_format = "0.00"
        rt = r0 + 2 + n
        a.cell(rt, 1, "Total").font = BOLD
        for k in range(2, len(years) + 2):
            a.cell(rt, k, f"=SUM({L(k)}{r0 + 2}:{L(k)}{rt - 1})").number_format = "0.00"
    a.column_dimensions["A"].width = 18

    # ---- Penetration (history DERIVED values, forecast formulas)
    pen = wb.create_sheet("Penetration")
    pen.cell(1, 1, "Household penetration (heat-equivalent) = heat-pump heat / household space + water heating need. "
                   "History DERIVED from Eurostat (values); forecast = annual heat-pump heat from 'Calc' / need "
                   "(formulas). Years between the last Eurostat year and 2025 are blank.").font = BOLD
    pyears = list(range(FIRST_HIST, FC_END.year + 1))
    header(pen, 2, ["Code", "Country", "Method"] + pyears)
    for r, geo in enumerate(codes, 3):
        cr = r - 1
        pen.cell(r, 1, geo), pen.cell(r, 2, inp.loc[geo, "name"])
        pen.cell(r, 3, f"=Countries!${L(COL['method'])}${cr}")
        base = int(inp.loc[geo, "base"])
        for j, y in enumerate(pyears, 4):
            if y <= base:
                v = penhist.loc[geo].get(y) if geo in penhist.index else None
                c = pen.cell(r, j, round(float(v), 4) if v is not None and pd.notna(v) else None)
            elif y >= FC_START.year:
                need = f"({POP(cr, str(y))}*Countries!${L(COL['hhheat'])}${cr}/Countries!${L(COL['hhsize'])}${cr})"
                c = pen.cell(r, j, f"=IFERROR(SUMIFS(Calc!$M$2:$M${last_calc},Calc!$A$2:$A${last_calc},\"{geo}\","
                                   f"Calc!$C$2:$C${last_calc},{y})/{need},0)")
            else:
                continue
            c.number_format = "0.0%"
    pen.column_dimensions["B"].width = 16
    pen.freeze_panes = "D3"

    # ---- History (values)
    h = wb.create_sheet("History")
    h.cell(1, 1, "Eurostat history (annual): capacity, heat = capacity x hours, electricity = heat / SPF "
                 "(DERIVED; countries without hours/SPF show 0)").font = BOLD
    hy = list(range(FIRST_HIST, 2025))
    for b, (label, key) in enumerate((("Installed capacity (GW)", "cap_GW"), ("Heat (TWh)", "heat_TWh"),
                                      ("Electricity (TWh)", "elec_TWh"))):
        r0 = 3 + b * (n + 4)
        h.cell(r0, 1, label).font = BOLD
        header(h, r0 + 1, ["Country"] + hy)
        for i, geo in enumerate(codes):
            h.cell(r0 + 2 + i, 1, inp.loc[geo, "name"])
            sub = hist[hist["geo"] == geo].set_index("year")
            for k, y in enumerate(hy, 2):
                h.cell(r0 + 2 + i, k, round(float(sub.loc[y, key]), 3))
    h.column_dimensions["A"].width = 18

    # ---- Dashboard: country picker + charts
    db = wb.create_sheet("Dashboard", 1)
    db.cell(1, 1, "Heat pump demand - pick a country in B2 (degree days on 'Degree days': actual, then normal)").font = Font(bold=True, size=12)
    db.cell(2, 1, "Country").font = BOLD
    pick = db.cell(2, 2, "Germany")
    pick.fill = YELLOW
    dv = DataValidation(type="list", formula1=f"=Countries!$B$2:$B${last_c}", allow_blank=False)
    db.add_data_validation(dv)
    dv.add("B2")
    header(db, 4, ["Month", "Electricity (avg GW) - selected", "Gas avoided (Bcf/d) - selected",
                   "Electricity (avg GW) - total", "Gas avoided (Bcf/d) - total"])
    ne, ng = wide["Electricity GW"], wide["Gas avoided Bcf-d"]
    for i, d in enumerate(mlist, 5):
        src = i - 3
        db.cell(i, 1, d).number_format = "mmm/yy"
        sel = f"MATCH($B$2,'Electricity GW'!$B$1:${L(n + 1)}$1,0)"
        db.cell(i, 2, f"=INDEX('Electricity GW'!$B{src}:${L(n + 1)}{src},1,{sel})").number_format = "0.00"
        db.cell(i, 3, f"=INDEX('Gas avoided Bcf-d'!$B{src}:${L(n + 1)}{src},1,{sel})").number_format = "0.000"
        db.cell(i, 4, f"='Electricity GW'!{L(n + 2)}{src}").number_format = "0.00"
        db.cell(i, 5, f"='Gas avoided Bcf-d'!{L(n + 2)}{src}").number_format = "0.000"
    lr = 4 + len(mlist)
    db.add_chart(line_chart("Heat pump electricity demand - selected country", "Average GW", db, 1, [2], 5, lr, 4), "G2")
    db.add_chart(line_chart("Gas a boiler would have burned - selected country", "Bcf/d", db, 1, [3], 5, lr, 4), "G21")
    db.add_chart(line_chart("Heat pump electricity demand - EU27 + UK + Norway", "Average GW", db, 1, [4], 5, lr, 4),
                 "S2")
    db.add_chart(line_chart("Gas a boiler would have burned - EU27 + UK + Norway", "Bcf/d", db, 1, [5], 5, lr, 4),
                 "S21")
    pr0 = lr + 3
    db.cell(pr0, 1, "Household penetration - selected country (history to the last Eurostat year, then forecast)").font = BOLD
    header(db, pr0 + 1, ["Year", "Penetration"])
    for i, y in enumerate(pyears):
        rr = pr0 + 2 + i
        db.cell(rr, 1, y)
        col = L(4 + i)
        db.cell(rr, 2, f"=IFERROR(1*INDEX(Penetration!${col}$3:${col}${2 + n},MATCH($B$2,Penetration!$B$3:$B${2 + n},0)),NA())"
                ).number_format = "0%"
    pch = line_chart("Heat pump household penetration - selected country", "Share of household heating", db, 1, [2],
                     pr0 + 2, pr0 + 1 + len(pyears), pr0 + 1)
    pch.x_axis.number_format = "0"
    pch.y_axis.number_format = "0%"
    db.add_chart(pch, "G40")
    for col, wdt in zip("ABCDE", (10, 20, 20, 20, 20)):
        db.column_dimensions[col].width = wdt

    wb.calculation.fullCalcOnLoad = True
    xlsx_charts.save_atomic(wb, path)
    out(f"Saved {path}: {n} countries, {len(mlist)} months, {last_calc - 1} Calc rows")


def make_pngs(path):
    """PNG copies of the dashboard totals and the 2030 country ranking, from a LibreOffice-recalculated copy (the
    workbook itself holds formulas only). Skipped when LibreOffice is not installed."""
    import shutil
    import subprocess
    import tempfile
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        out("LibreOffice not found: PNGs skipped")
        return
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, path],
                       check=True, capture_output=True, timeout=300)
        calc = os.path.join(tmp, os.path.basename(path))
        e = pd.read_excel(calc, sheet_name="Electricity GW").dropna(subset=["Month"])
        g = pd.read_excel(calc, sheet_name="Gas avoided Bcf-d").dropna(subset=["Month"])
        ann = pd.read_excel(calc, sheet_name="Annual", header=None)
    e = e[pd.to_datetime(e["Month"], errors="coerce").notna()]
    g = g[pd.to_datetime(g["Month"], errors="coerce").notna()]
    total = "Total (EU27 + UK + Norway)"
    pngdir = os.path.join(ROOT, "output", "PNG Charts")
    os.makedirs(pngdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(path))[0]
    for df, col, title, unit, name in (
            (e, total, "Heat pump electricity demand - EU27 + UK + Norway (degree days: actual, then normal)", "Average GW", "Electricity_total"),
            (g, total, "Gas a boiler would have burned - EU27 + UK + Norway (degree days: actual, then normal)", "Bcf/d", "Gas_avoided_total")):
        fig, ax = plt.subplots(figsize=(10, 4.5))
        ax.plot(pd.to_datetime(df["Month"]), df[col].astype(float), color="#1f77b4", lw=2)
        ax.set_title(title, fontsize=12, fontweight="bold")
        ax.set_ylabel(unit)
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b/%y"))
        ax.grid(axis="y", alpha=0.3)
        for sp in ax.spines.values():
            sp.set_visible(False)
        fig.text(0.01, 0.01, "Source: Eurostat nrg_inf_hptc and nrg_d_hhq, When2Heat COP curve, NASA POWER normals; "
                 "heat_pump_demand_model.xlsx", fontsize=7, color="grey")
        fig.tight_layout(rect=(0, 0.03, 1, 1))
        fig.savefig(os.path.join(pngdir, f"{stem}__{name}.png"), dpi=130)
        plt.close(fig)
    # 2030 electricity by country (first block of 'Annual')
    hdr = ann.index[ann[0] == "Country"][0]
    blk = ann.iloc[hdr + 1:]
    blk = blk[blk[0].notna() & (blk[0] != "Total")].iloc[:len(COUNTRIES)]
    years = [int(v) for v in ann.iloc[hdr, 1:].dropna()]
    v2030 = blk.set_index(0)[1 + years.index(2030)].astype(float).sort_values()
    fig, ax = plt.subplots(figsize=(9, 7))
    ax.barh(v2030.index, v2030.values, color="#1f77b4")
    ax.set_title("Heat pump electricity demand by country, 2030 (degree days: actual, then normal)", fontsize=12, fontweight="bold")
    ax.set_xlabel("TWh per year")
    for sp in ax.spines.values():
        sp.set_visible(False)
    fig.text(0.01, 0.01, "Source: Eurostat nrg_inf_hptc and nrg_d_hhq, When2Heat COP curve, NASA POWER normals; "
             "heat_pump_demand_model.xlsx", fontsize=7, color="grey")
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    fig.savefig(os.path.join(pngdir, f"{stem}__Electricity_by_country_2030.png"), dpi=130)
    plt.close(fig)
    out(f"PNGs written to {pngdir}")


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    stock, upd_hptc = heat_pump_stock()
    hh, upd_hhq = household_split()
    inp, sts = build_inputs(stock, hh)
    hist = history_table(stock)
    norm, act = weather_normals()
    pop = population()
    penhist = add_penetration(inp, sts, stock, pop, household_size(), household_heat())
    build(args.out, inp, sts, hist, cop_parameters(), norm, act, {"hptc": upd_hptc, "hhq": upd_hhq}, pop, penhist)
    make_pngs(args.out)


if __name__ == "__main__":
    main()
