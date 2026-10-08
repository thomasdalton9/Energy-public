"""
Long-term fundamentals page for the master workbooks: "Dashboard - Long-term".

Annual history per country and for the region, from two pulls:
  long_term_energy.xlsx  (LONG_TERM_ENERGY.py)  Power: Ember yearly (2000-), generation by fuel, demand, net imports,
                                                capacity, power-sector CO2. Fuels: Energy Institute Statistical Review
                                                (gas, oil, coal, primary energy, 1965-).
  macro_drivers.xlsx     (MACRO_DRIVERS.py)     World Bank GDP / population / industry / urbanisation / access,
                                                IMF WEO GDP growth and population incl. forecasts, and annual
                                                cooling / heating degree days (NASA POWER, population-weighted cities).
Both are long format (Power / Fuels: year, iso3, variable, value) or wide (macro: year x iso3) - see those scripts.

These are compiled annual statistics (Ember and the EI compile them from national statistics offices and grid
operators); they are the long consistent history the raw feeds on the other dashboards are too short to give. The page:
  - a summary table, one row per country plus the region: latest demand, 10-year demand and GDP growth, demand/GDP
    elasticity, demand per head, fuel shares and how wind + solar moved over 5 years, gas balance and import
    dependence, IMF GDP growth outlook, degree days, capacity and fleet utilisation;
  - charts: region generation by fuel and shares, demand vs GDP vs population (index), gas production vs consumption,
    demand and demand per head by country, wind + solar share by country, gas consumption by country, IMF GDP
    growth (history + forecast), degree days.

add_long_term_dashboard() is called by each master after its other dashboards.
"""
import os

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

import xlsx_charts

FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Other"]
DARK_BLUE = "17365D"
START_YEAR = 2000
TOP_N = 7   # lines per by-country chart (the palette has 8 slots; the rest fold into 'Other')

# master country names -> ISO3 (names as the masters use them)
ISO3 = {
    # South & Southeast Asia
    "India": "IND", "Pakistan": "PAK", "Bangladesh": "BGD", "Sri Lanka": "LKA", "Nepal": "NPL", "Bhutan": "BTN",
    "Thailand": "THA", "Vietnam": "VNM", "Philippines": "PHL", "Indonesia": "IDN", "Malaysia": "MYS",
    "Singapore": "SGP", "Myanmar": "MMR", "Cambodia": "KHM", "Laos": "LAO", "Brunei": "BRN", "Timor-Leste": "TLS",
    # South & Central America, Caribbean
    "Argentina": "ARG", "Bolivia": "BOL", "Brazil": "BRA", "Chile": "CHL", "Colombia": "COL", "Ecuador": "ECU",
    "Peru": "PER", "Uruguay": "URY", "Paraguay": "PRY", "Venezuela": "VEN", "Guyana": "GUY", "Suriname": "SUR",
    "Guatemala": "GTM", "El Salvador": "SLV", "Honduras": "HND", "Nicaragua": "NIC", "Costa Rica": "CRI",
    "Panama": "PAN", "Belize": "BLZ", "Trinidad and Tobago": "TTO", "Puerto Rico": "PRI", "Jamaica": "JAM",
    "Dominican Republic": "DOM",
    # North America
    "United States": "USA", "Canada": "CAN", "Mexico": "MEX",
    # Europe
    "Germany": "DEU", "France": "FRA", "Spain": "ESP", "Italy": "ITA", "Netherlands": "NLD", "Belgium": "BEL",
    "Poland": "POL", "Austria": "AUT", "Switzerland": "CHE", "Czechia": "CZE", "Slovakia": "SVK", "Hungary": "HUN",
    "Romania": "ROU", "Bulgaria": "BGR", "Greece": "GRC", "Portugal": "PRT", "Croatia": "HRV", "Slovenia": "SVN",
    "Denmark": "DNK", "Sweden": "SWE", "Norway": "NOR", "Finland": "FIN", "Ireland": "IRL", "Estonia": "EST",
    "Latvia": "LVA", "Lithuania": "LTU", "Serbia": "SRB", "Bosnia and Herzegovina": "BIH", "Montenegro": "MNE",
    "North Macedonia": "MKD", "Albania": "ALB", "Kosovo": "XKX", "United Kingdom": "GBR", "Turkey": "TUR",
    "Cyprus": "CYP", "Luxembourg": "LUX", "Malta": "MLT",
    # Australia & New Zealand
    "Australia": "AUS", "New Zealand": "NZL",
    # China
    "China": "CHN",
}

REGIONS = {
    "South & Southeast Asia": ["India", "Pakistan", "Bangladesh", "Sri Lanka", "Nepal", "Bhutan", "Thailand",
                               "Vietnam", "Philippines", "Indonesia", "Malaysia", "Singapore", "Myanmar", "Cambodia",
                               "Laos", "Brunei", "Timor-Leste"],
    "South & Central America and the Caribbean": [
        "Argentina", "Bolivia", "Brazil", "Chile", "Colombia", "Ecuador", "Peru", "Uruguay", "Paraguay", "Venezuela",
        "Guyana", "Suriname", "Guatemala", "El Salvador", "Honduras", "Nicaragua", "Costa Rica", "Panama", "Belize",
        "Trinidad and Tobago", "Puerto Rico", "Jamaica", "Dominican Republic"],
    "North America": ["United States", "Canada", "Mexico"],
    "Australia & New Zealand": ["Australia", "New Zealand"],
    "China": ["China"],
    "Europe": ["Germany", "France", "United Kingdom", "Italy", "Spain", "Poland", "Netherlands", "Turkey", "Belgium",
               "Austria", "Switzerland", "Czechia", "Slovakia", "Hungary", "Romania", "Bulgaria", "Greece", "Portugal",
               "Croatia", "Slovenia", "Denmark", "Sweden", "Norway", "Finland", "Ireland", "Estonia", "Latvia",
               "Lithuania", "Luxembourg", "Serbia", "Bosnia and Herzegovina", "Montenegro", "North Macedonia",
               "Albania", "Kosovo", "Cyprus", "Malta"],
}


# ----------------------------------------------------------------------------------------------------------- data
def _long(path, sheet):
    """Long sheet (year, iso3, variable, value) -> {variable: frame year x iso3}."""
    try:
        d = pd.read_excel(path, sheet_name=sheet)
    except Exception:  # noqa: BLE001  (file or sheet missing)
        return {}
    d["value"] = pd.to_numeric(d["value"], errors="coerce")
    d["year"] = pd.to_numeric(d["year"], errors="coerce")
    d = d.dropna(subset=["year", "value"])
    d["year"] = d["year"].astype(int)
    out = {}
    for v, g in d.groupby("variable"):
        f = g.pivot_table(index="year", columns="iso3", values="value", aggfunc="sum")
        # a lone zero between two real values is a hole in the source table (e.g. EI's Norway gas production 1998),
        # not a year of no output: left blank (no invented values; region totals skip that year)
        f = f.where(~((f == 0) & (f.shift(1) > 1) & (f.shift(-1) > 1)))
        out[v] = f
    return out


def _wide(path):
    """Wide macro sheets (year or date x iso3) -> {sheet: frame}."""
    out = {}
    try:
        x = pd.ExcelFile(path)
    except Exception:  # noqa: BLE001
        return out
    for s in x.sheet_names:
        if s in ("Units", "Countries", "City_T2M_daily") or s.startswith("Chart"):
            continue
        d = pd.read_excel(x, sheet_name=s)
        if d.empty:
            continue
        key = d.columns[0]
        if key == "year":
            d["year"] = pd.to_numeric(d["year"], errors="coerce")
            d = d.dropna(subset=["year"]).astype({"year": int}).set_index("year")
        elif key == "date":
            d = d.set_index(pd.to_datetime(d["date"], errors="coerce")).drop(columns="date")
        else:
            continue
        out[s] = d.apply(pd.to_numeric, errors="coerce")
    return out


def load(data_dir):
    lte = os.path.join(data_dir, "long_term_energy.xlsx")
    mac = os.path.join(data_dir, "macro_drivers.xlsx")
    data = {}
    data.update(_long(lte, "Power"))
    data.update(_long(lte, "Fuels"))
    data.update(_wide(mac))
    return data


def get(data, var, codes):
    """Frame year x iso3 for the given codes (missing codes / variable -> empty columns)."""
    f = data.get(var)
    if f is None:
        return pd.DataFrame(columns=codes, dtype=float)
    return f.reindex(columns=codes)


def region_sum(f, min_share=0.97):
    """Sum over countries, only years where the reporting countries hold >= min_share of the latest full year's total
    (a country that has not reported yet would otherwise pull the total down)."""
    if f.empty:
        return pd.Series(dtype=float)
    s = f.sum(axis=1, min_count=1)
    full = f.notna().sum(axis=1)
    if not len(full) or full.max() == 0:
        return pd.Series(dtype=float)
    # weight of each country = its latest value; a year counts if its reporting countries carry >= min_share of it
    w = f.ffill().iloc[-1].fillna(0).abs()   # abs: net imports can be negative (exporters)
    rep = f.notna().mul(w, axis=1).sum(axis=1)
    return s[rep >= min_share * w.sum()].dropna()


def cagr(s, years=10):
    s = s.dropna()
    if len(s) < 2:
        return None
    end = s.index.max()
    start = end - years
    if start not in s.index or s[start] <= 0 or s[end] <= 0:
        return None
    return ((s[end] / s[start]) ** (1 / years) - 1) * 100


def _num(v):
    return None if v is None or pd.isna(v) else float(v)


def _latest(s):
    s = s.dropna()
    return (int(s.index.max()), float(s.iloc[-1])) if len(s) else (None, None)


# ------------------------------------------------------------------------------------------------------- summary
SUMMARY_COLS = [
    # (header, number format, width)
    ("Country", None, 26), ("Year", "0", 7),
    ("Demand TWh", "#,##0", 10), ("Demand growth 10y %/yr", "0.0", 11), ("GDP growth 10y %/yr", "0.0", 10),
    ("Elasticity (demand / GDP growth)", "0.00", 12), ("Demand per head kWh", "#,##0", 10),
    ("GDP per head PPP $k", "0.0", 10), ("Coal %", "0", 7), ("Gas %", "0", 7), ("Wind + solar %", "0", 8),
    ("Wind + solar change 5y pp", "+0;-0;0", 9), ("Hydro %", "0", 7), ("Nuclear %", "0", 8),
    ("Net imports % of demand", "0.0", 9), ("Capacity GW", "#,##0", 9), ("Fleet utilisation %", "0", 9),
    ("Gas consumption bcm", "#,##0.0", 11), ("Gas production bcm", "#,##0.0", 10),
    ("Gas import dependence %", "0", 10), ("Gas demand growth 10y %/yr", "0.0", 11),
    ("IMF GDP growth next 5y %/yr", "0.0", 10), ("Population growth next 5y %/yr", "0.00", 11),
    ("CDD (18°C) 10y avg", "#,##0", 9), ("HDD (18°C) 10y avg", "#,##0", 9),
]


def summary_row(name, s):
    """s: dict of annual series for one country (or the region total)."""
    year, demand = _latest(s["demand"])
    gen = s["gen"]
    g = gen.dropna(how="all")
    shares = {}
    if len(g):
        last = g.iloc[-1]
        tot = last.sum()
        if tot > 0:
            shares = (last / tot * 100).to_dict()
        ws = (g[["Wind", "Solar"]].sum(axis=1) / g.sum(axis=1).replace(0, float("nan")) * 100).dropna()
        ws5 = ws[g.index.max()] - ws[g.index.max() - 5] if (g.index.max() - 5) in ws.index else None
    else:
        ws5 = None
    gdp_g = cagr(s["gdp"])
    dem_g = cagr(s["demand"])
    pop_y, pop = _latest(s["pop"].loc[:year] if year else s["pop"])
    ppp_y, ppp = _latest(s["ppp"].loc[:year] if year else s["ppp"])
    _, cap = _latest(s["cap"])
    _, gen_tot = _latest(g.sum(axis=1)) if len(g) else (None, None)
    cap_y = s["cap"].get(g.index.max()) if len(g) else None   # capacity of the same year as the generation
    util = gen_tot * 1000 / (cap_y * 8760) * 100 if gen_tot and cap_y and not pd.isna(cap_y) else None
    _, nimp = _latest(s["net_imports"])
    gy, gcons = _latest(s["gas_cons"])
    _, gprod = _latest(s["gas_prod"].loc[:gy] if gy else s["gas_prod"])
    dep = (gcons - gprod) / gcons * 100 if gcons and gprod is not None else None   # blank where the EI has no
    # production row for the country (it groups small producers into 'Other')
    # IMF: the five years after the latest actual GDP year
    imf = s["imf_gdp"].dropna()
    act = s["gdp"].dropna()
    fwd = imf[imf.index > (act.index.max() if len(act) else 9999)].head(5)
    popf = s["imf_pop"].dropna()
    pop_fwd = None
    if len(fwd) and len(popf):
        a, b = fwd.index.min() - 1, fwd.index.max()
        if a in popf.index and b in popf.index and popf[a] > 0:
            pop_fwd = ((popf[b] / popf[a]) ** (1 / (b - a)) - 1) * 100
    cdd = s["cdd"].dropna().tail(10)
    hdd = s["hdd"].dropna().tail(10)
    return [name, year, _num(demand), dem_g, gdp_g,
            dem_g / gdp_g if dem_g is not None and gdp_g and abs(gdp_g) > 0.5 else None,
            demand * 1000 / pop if demand and pop else None,   # TWh / million people = MWh/head; x1000 = kWh
            ppp / pop if ppp and pop else None,
            shares.get("Coal"), shares.get("Gas"),
            (shares.get("Wind", 0) + shares.get("Solar", 0)) if shares else None, ws5,
            shares.get("Hydro"), shares.get("Nuclear"),
            nimp / demand * 100 if nimp is not None and demand else None,
            cap, util, gcons, gprod, dep, cagr(s["gas_cons"]),
            fwd.mean() if len(fwd) else None, pop_fwd,
            cdd.mean() if len(cdd) else None, hdd.mean() if len(hdd) else None]


def country_series(data, codes):  # noqa: C901
    """Annual series per country code: dict iso3 -> dict of series / frames, plus the region total."""
    gen = {f: get(data, f"Gen_{f}_TWh", codes) for f in FUELS}
    one = {
        "demand": get(data, "Demand_TWh", codes), "net_imports": get(data, "Net_Imports_TWh", codes),
        "cap": get(data, "Cap_Total_GW", codes), "gas_cons": get(data, "Gas_Consumption_bcm", codes),
        "gas_prod": get(data, "Gas_Production_bcm", codes), "gdp": get(data, "GDP_real_USD_bn", codes),
        "ppp": get(data, "GDP_PPP_USD_bn", codes), "pop": get(data, "Population_m", codes),
        "imf_gdp": get(data, "GDP_growth_IMF_pct", codes), "imf_pop": get(data, "Population_IMF_m", codes),
        "cdd": get(data, "CDD_18", codes), "hdd": get(data, "HDD_18", codes),
    }
    per = {}
    for c in codes:
        per[c] = {k: v[c] for k, v in one.items()}
        per[c]["gen"] = pd.DataFrame({f: gen[f][c] for f in FUELS})
    reg = {k: region_sum(v) for k, v in one.items() if k not in ("imf_gdp", "imf_pop", "cdd", "hdd")}
    # region GDP growth outlook: GDP-weighted average of the countries' IMF growth; degree days: population-weighted
    w = one["gdp"].ffill().iloc[-1] if len(one["gdp"].dropna(how="all")) else pd.Series(dtype=float)
    reg["imf_gdp"] = (one["imf_gdp"].mul(w, axis=1).sum(axis=1, min_count=1) /
                      one["imf_gdp"].notna().mul(w, axis=1).sum(axis=1).replace(0, float("nan"))).dropna()
    reg["imf_pop"] = region_sum(one["imf_pop"])
    pw = one["pop"].ffill().iloc[-1] if len(one["pop"].dropna(how="all")) else pd.Series(dtype=float)
    for k in ("cdd", "hdd"):
        f = one[k]
        reg[k] = (f.mul(pw, axis=1).sum(axis=1, min_count=1) /
                  f.notna().mul(pw, axis=1).sum(axis=1).replace(0, float("nan"))).dropna()
    # region generation: the same years as the region's demand (every country reporting), each fuel summed
    yrs = reg["demand"].index
    reg["gen"] = pd.DataFrame({f: gen[f].reindex(yrs).sum(axis=1, min_count=1) for f in FUELS}, index=yrs)
    capf = {f: get(data, f"Cap_{f}_GW", codes) for f in FUELS}
    reg["cap_by_fuel"] = pd.DataFrame({f: capf[f].reindex(reg["cap"].index).sum(axis=1, min_count=1) for f in FUELS},
                                      index=reg["cap"].index)
    return per, reg, gen, one


# -------------------------------------------------------------------------------------------------------- charts
def _year_index(df):
    df = df.copy()
    df.index = pd.to_datetime(df.index.astype(int).astype(str), format="%Y")
    return df


def _top(f, n=TOP_N):
    """Columns with the largest latest values (keeps by-country charts readable)."""
    last = f.ffill().iloc[-1].dropna() if len(f) else pd.Series(dtype=float)
    return list(last.sort_values(ascending=False).index[:n])


def _chart(wb, used, sheet_name_fn, label, df, title, y_title, kind, chart_w, chart_h, note=None):
    df = df.dropna(how="all")
    if df.empty or df.shape[1] == 0:
        return None
    df, y_title = xlsx_charts.monthly_energy_to_gw(df, y_title, title)   # generation: TWh per year -> average GW
    ws = wb.create_sheet(sheet_name_fn(f"Long {label} data", used))
    d, n_bars = xlsx_charts.prepare(_year_index(df))
    if d.empty:
        return None
    xlsx_charts.write_table(ws, d, "%Y")
    if note:
        ws.cell(row=1, column=d.shape[1] + 4, value=note)
    ch = xlsx_charts.build_chart(ws, d, n_bars if kind != "line" else d.shape[1], title, y_title, kind,
                                 date_format="%Y", width=chart_w, height=chart_h, gridlines=False,
                                 inner=xlsx_charts.DASHBOARD_INNER)
    return ch, ws.title


def add_long_term_dashboard(wb, used, sheet_name_fn, region, data_dir, chart_w, chart_h, rows_per_chart=24,
                            cols=("B", "F"), countries=None, degree_days=("CDD_18",), index=None):
    """Writes 'Dashboard - Long-term' (summary table + charts) and its 'Long ... data' tabs. countries: master
    names (default REGIONS[region]). Returns a list of notes (missing inputs)."""
    data = load(data_dir)
    notes = []
    if not data:
        notes.append("long_term_energy.xlsx / macro_drivers.xlsx not found: long-term page not built")
        return notes
    names = [c for c in (countries or REGIONS[region]) if c in ISO3]
    codes = [ISO3[c] for c in names]
    label = {ISO3[c]: c for c in names}
    per, reg, gen, one = country_series(data, codes)

    dash = wb.create_sheet("Dashboard - Long-term", index) if index is not None else wb.create_sheet(
        "Dashboard - Long-term")
    used.add("Dashboard - Long-term")
    dash["B1"] = f"{region} - long-term fundamentals (annual)"
    dash["B1"].font = Font(bold=True, size=18, color=DARK_BLUE)
    dash["B2"] = ("Annual statistics compiled from national sources: power from Ember's yearly data (2000-), fuels "
                  "from the Energy Institute Statistical Review, GDP and population from the World Bank, outlook from "
                  "the IMF World Economic Outlook, degree days from NASA POWER (population-weighted cities). These "
                  "give the long consistent history; the raw monthly/daily feeds are on the other dashboards.")
    dash["B2"].font = Font(italic=True, color="6B6B6B")
    dash["B3"] = ("Growth = compound annual rate over 10 years to the latest year. Elasticity = demand growth / GDP "
                  "growth. Fleet utilisation = generation / (capacity x 8,760 h). Gas import dependence = "
                  "(consumption - production) / consumption. Region row: countries summed (years where all report).")
    dash["B3"].font = Font(italic=True, size=9, color="6B6B6B")

    rows = [summary_row(region + " (total)", reg)]
    for c in codes:
        r = summary_row(label[c], per[c])
        if r[1] is not None or r[17] is not None:
            rows.append(r)
    head_row = 5
    for j, (h, _, w) in enumerate(SUMMARY_COLS):
        cell = dash.cell(row=head_row, column=2 + j, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color=DARK_BLUE)
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        dash.column_dimensions[cell.column_letter].width = w
    dash.row_dimensions[head_row].height = 45
    rows = rows[:1] + sorted(rows[1:], key=lambda r: -(r[2] or 0))
    for i, r in enumerate(rows, start=head_row + 1):
        for j, ((_, fmt, _), v) in enumerate(zip(SUMMARY_COLS, r)):
            v = None if v is None or (isinstance(v, float) and pd.isna(v)) else v
            cell = dash.cell(row=i, column=2 + j, value=v)
            if fmt and v is not None:
                cell.number_format = fmt
            if i == head_row + 1:
                cell.font = Font(bold=True)
    dash.freeze_panes = dash.cell(row=head_row + 1, column=3)
    row = head_row + len(rows) + 2

    # ------------------------------------------------------------------------------------------------- charts
    charts = []
    yr = lambda f: f[f.index >= START_YEAR]   # noqa: E731
    src_power = "Ember yearly electricity data (compiled from national statistics)"
    g = yr(reg["gen"]).dropna(how="all")
    charts.append((_chart(wb, used, sheet_name_fn, "gen by fuel", g, f"{region} power generation by source",
                          "TWh per year", "stacked_bar", chart_w, chart_h), src_power))
    if len(g):
        sh = g.div(g.sum(axis=1), axis=0) * 100
        charts.append((_chart(wb, used, sheet_name_fn, "gen shares", sh.round(1), f"{region} generation mix",
                              "% of generation", "line", chart_w, chart_h), src_power))
    idx = pd.DataFrame({"Power demand": reg["demand"], "Real GDP": reg["gdp"], "Population": reg["pop"]})
    idx = yr(idx).dropna()
    if len(idx):
        idx = (idx / idx.iloc[0] * 100).round(1)
        charts.append((_chart(wb, used, sheet_name_fn, "demand vs GDP", idx,
                              f"{region} power demand vs real GDP and population", f"index, {idx.index.min()} = 100",
                              "line", chart_w, chart_h), "Ember (demand); World Bank (GDP at constant 2015 US$, "
                                                         "population)"))
    gas = pd.DataFrame({"Consumption": reg["gas_cons"], "Production": reg["gas_prod"]})
    gas = gas[gas.index >= 1990].dropna(how="all")
    charts.append((_chart(wb, used, sheet_name_fn, "gas balance", gas.round(1),
                          f"{region} natural gas production vs consumption (gap = net imports)", "bcm per year",
                          "line", chart_w, chart_h), "Energy Institute Statistical Review of World Energy"))
    dem = yr(one["demand"]).rename(columns=label)
    top = _top(dem)
    charts.append((_chart(wb, used, sheet_name_fn, "demand by country", dem[top],
                          f"Power demand by country (largest {len(top)})", "TWh per year", "line", chart_w, chart_h),
                   src_power))
    pc = (yr(one["demand"]) * 1e9 / (one["pop"].reindex(one["demand"].index) * 1e6)).rename(columns=label)
    charts.append((_chart(wb, used, sheet_name_fn, "demand per head", pc[[c for c in top if c in pc]].round(0),
                          "Power demand per head", "kWh per person per year", "line", chart_w, chart_h),
                   "Ember (demand); World Bank (population)"))
    wsol = (yr(gen["Wind"]).add(yr(gen["Solar"]), fill_value=0) /
            sum(yr(gen[f]).fillna(0) for f in FUELS).replace(0, float("nan")) * 100).rename(columns=label)
    charts.append((_chart(wb, used, sheet_name_fn, "wind solar share", wsol[[c for c in top if c in wsol]].round(1),
                          "Wind + solar share of generation", "% of generation", "line", chart_w, chart_h),
                   src_power))
    gc = one["gas_cons"]
    gc = gc[gc.index >= 1990].rename(columns=label)
    gtop = _top(gc)
    charts.append((_chart(wb, used, sheet_name_fn, "gas by country", gc[gtop].round(1),
                          f"Natural gas consumption by country (largest {len(gtop)})", "bcm per year", "line",
                          chart_w, chart_h), "Energy Institute Statistical Review of World Energy"))
    imf = one["imf_gdp"].rename(columns=label)
    imf = imf[imf.index >= 2010]
    gdp_top = _top(one["gdp"].rename(columns=label))
    reg_imf = reg["imf_gdp"][reg["imf_gdp"].index >= 2010]
    imf_df = pd.concat([reg_imf.rename(f"{region} (GDP-weighted)").to_frame(),
                        imf[[c for c in gdp_top[:TOP_N - 1] if c in imf]]], axis=1) if len(imf.columns) else \
        pd.DataFrame()
    charts.append((_chart(wb, used, sheet_name_fn, "GDP growth", imf_df.round(1),
                          "Real GDP growth, history and IMF forecast", "% per year", "line", chart_w, chart_h,
                          note="IMF World Economic Outlook: years after the latest actual are forecasts"),
                   "IMF World Economic Outlook (DataMapper)"))
    for dd in degree_days:
        f = data.get(dd)
        if f is None:
            notes.append(f"{dd}: macro_drivers.xlsx has no degree days")
            continue
        f = f.reindex(columns=codes).rename(columns=label)
        f = f[[c for c in top if c in f]]
        kind = "cooling" if dd.startswith("CDD") else "heating"
        charts.append((_chart(wb, used, sheet_name_fn, f"{dd[:3]}", f.round(0),
                              f"{kind.capitalize()} degree days (base 18°C), population-weighted cities",
                              "degree days per year", "line", chart_w, chart_h),
                       "NASA POWER daily temperature (MERRA-2), population-weighted largest cities"))
    charts.append((_chart(wb, used, sheet_name_fn, "capacity", yr(reg["cap_by_fuel"]).dropna(how="all"),
                          f"{region} installed capacity by source", "GW", "stacked_bar", chart_w, chart_h),
                   src_power))

    charts = [(c, s) for c, s in charts if c]
    missing = [k for k in ("Demand_TWh", "Gas_Consumption_bcm", "GDP_real_USD_bn", "GDP_growth_IMF_pct")
               if k not in data]
    if missing:
        notes.append("long-term inputs missing: " + ", ".join(missing))
        dash.cell(row=row - 1, column=2, value="Not available this run: " + ", ".join(missing)).font = Font(
            color="E34948")
    # second chart column: the first column whose left edge is past the first chart (summary widths differ from
    # the other dashboards, so the fixed 'F' would overlap)
    from openpyxl.utils import get_column_letter
    edge, j = 0.0, 2
    while edge < chart_w + 0.5:
        edge += (dash.column_dimensions[get_column_letter(j)].width or 8.43) * 0.19
        j += 1
    cols = (cols[0], get_column_letter(j))
    for k, ((ch, tab), src) in enumerate(charts):
        r = row + (k // 2) * rows_per_chart
        dash.add_chart(ch, f"{cols[k % 2]}{r}")
        note = dash[f"{cols[k % 2]}{r + rows_per_chart - 2}"]
        note.value = f"Source: {src} (data tab '{tab}')"
        note.font = Font(italic=True, size=8, color="6B6B6B")
    print(f"long-term page: {len(rows) - 1} countries, {len(charts)} charts" + (f"; {notes}" if notes else ""))
    return notes


# PNGs of the long-term charts (CLAUDE.md: every chart also as a PNG in output/PNG Charts)
PNG_TABS = [("gen by fuel", "power generation by source", "GW (annual average)", "stacked_bar"),
            ("gen shares", "generation mix", "% of generation", "line"),
            ("demand vs GDP", "power demand vs real GDP and population", "index", "line"),
            ("gas balance", "natural gas production vs consumption", "bcm per year", "line"),
            ("demand per head", "power demand per head", "kWh per person per year", "line"),
            ("GDP growth", "real GDP growth, history and IMF forecast", "% per year", "line")]


def render_pngs(master, out_dir=None):
    import png_charts
    out_dir = out_dir or png_charts.OUT
    names = pd.ExcelFile(master).sheet_names
    try:
        region = pd.read_excel(master, sheet_name="Dashboard - Long-term", header=None).iloc[0, 1]
        region = str(region).split(" - long-term")[0]
    except Exception:  # noqa: BLE001
        region = ""
    for label, title, units, kind in PNG_TABS:
        tab = next((n for n in names if n in (f"Long {label} data", f"LT {label} data")), None)
        if tab:
            png_charts.render_sheet(master, tab, f"{region} {title}".strip(), units, out_dir, kind)


if __name__ == "__main__":
    import sys
    for m in sys.argv[1:]:
        render_pngs(m)
