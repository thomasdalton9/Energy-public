"""
Add native Excel charts to a pulled workbook (see CLAUDE.md: every xlsx
we pull carries charts of its data). Run by each scheduled workflow
right after its pull:

    python3 add_charts.py "output/Data and Chart Outputs/colombia_gas_demand_by_sector.xlsx"

A registry keyed by file name says what to chart. Each entry returns a
list of chart specs; each spec becomes its own sheet ("Chart", "Chart -
<name>") holding the plotted table and a native Excel chart, rebuilt on
every run. Storage/level series use the AGSI-style water-year chart.
Files not in the registry get a generic line chart of their first data
sheet's numeric columns, so nothing is left without a chart.
"""
import os
import re
import sys
import warnings

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import water_year_chart  # noqa: E402
import xlsx_charts  # noqa: E402

warnings.filterwarnings("ignore")


def read(path, sheet):
    df = pd.read_excel(path, sheet_name=sheet)
    return df.drop(columns=[c for c in df.columns if str(c).startswith("Unnamed")])


def by_date(df, col):
    df = df.copy()
    df[col] = pd.to_datetime(df[col].astype(str), errors="coerce")
    return df.dropna(subset=[col]).set_index(col).sort_index()


def cols(df, *names):
    return [c for c in names if c in df.columns]


def spec(name, df, title, units, kind="line", date_format="%Y-%m", line_cols=()):
    return {"name": name, "df": df, "title": title, "units": units, "kind": kind, "date_format": date_format,
            "line_cols": line_cols}


def daily(df, start=None):
    """Daily series charted from `start` (keeps native charts responsive)."""
    return df[df.index >= start] if start else df


def monthly_mean(df, start=None):
    """Long daily series -> monthly averages (daily stacks over years are unreadable)."""
    d = daily(df, start)
    return d.resample("MS").mean().dropna(how="all")


def weekly_mean(df, start=None):
    d = daily(df, start)
    return d.resample("W-MON", label="left", closed="left").mean().dropna(how="all")


UNIT_SUFFIX = re.compile(r"_(Mt|Bcm|GWh|MWh|MW|TWh|10k_units|10k_kW|kt)$")


def label_and_unit(col):
    m = UNIT_SUFFIX.search(str(col))
    unit = {"10k_units": "10,000 units", "10k_kW": "10,000 kW"}.get(m.group(1), m.group(1)) if m else ""
    label = UNIT_SUFFIX.sub("", str(col)).replace("_", " ")
    return label, unit


# ------------------------------------------------------------------ registry

def argentina(p):
    d = by_date(read(p, "National"), "date")
    sectors = cols(d, "centrales_electricas", "industria", "residencial", "comercial", "gnc", "entes_oficiales")
    out = d[sectors].rename(columns={"centrales_electricas": "Power", "industria": "Industry", "residencial": "Residential",
                                     "comercial": "Commercial", "gnc": "Vehicle CNG", "entes_oficiales": "Government"})
    if "produccion_gas_natural" in d:
        out["Production"] = d["produccion_gas_natural"]
    return [spec("Demand", out, "Argentina gas demand by sector vs production", "million m3 per month", "stacked_bar",
                 line_cols=("Production",))]


def brazil(p):
    d = by_date(read(p, "Demand by segment"), "date")
    last = d.index.max().strftime("%b/%y")
    out = [spec("Demand", d[cols(d, "Power_Generation", "Industrial", "Residential", "Commercial", "Automotive",
                                 "Cogeneration", "Other_incl_CNG")].rename(
                    columns={"Power_Generation": "Power generation", "Automotive": "Vehicle CNG",
                             "Other_incl_CNG": "Other"}),
                f"Brazil gas demand by segment (MME, last published {last})", "million m3/day", "stacked_bar")]
    sheets = pd.ExcelFile(p).sheet_names
    if "Grid demand (ANP)" in sheets:   # pipeline-grid deliveries by consumer type, current (ANP)
        g = by_date(read(p, "Grid demand (ANP)"), "date")
        out.append(spec("Grid demand", g[cols(g, "Power_Generation", "Distributors", "Refineries", "Fertiliser", "Other")]
                        .rename(columns={"Power_Generation": "Power generation", "Distributors": "Distributors (city gates)"}),
                        "Brazil gas delivered from the transport pipelines, by consumer (ANP)", "million m3/day",
                        "stacked_bar"))
    if "Supply (ANP)" in sheets:
        s = by_date(read(p, "Supply (ANP)"), "date")
        s = s[cols(s, "Domestic_Available", "Bolivia_Pipeline", "Argentina_Pipeline", "LNG_Implied")].dropna(how="any")
        out.append(spec("Supply", s.rename(columns={"Domestic_Available": "Domestic (available)",
                                                    "Bolivia_Pipeline": "Bolivia (pipeline)",
                                                    "Argentina_Pipeline": "Argentina (pipeline)",
                                                    "LNG_Implied": "LNG (imports less pipeline)"}),
                        "Brazil gas supply by source (ANP)", "million m3/day", "stacked_bar"))
    return out


def bolivia(p):
    d = by_date(read(p, "Demand by sector"), "Month")
    c = [x for x in d.columns if x.endswith("_mcm_per_day") and not x.startswith("Total")]
    return [spec("Demand", d[c].rename(columns=lambda x: x.replace("_mcm_per_day", "").replace("_", " ")),
                 "Bolivia gas demand by sector", "million m3/day", "stacked_bar")]


def uruguay(p):
    d = by_date(read(p, "Demand by sector"), "Month")
    c = [f"{x}_mcm_per_day" for x in ("Residential", "Commercial", "Industrial", "Power", "Energy_own_use")]
    return [spec("Demand", d[cols(d, *c)].rename(columns=lambda x: x.replace("_mcm_per_day", "").replace("_", " ")),
                 "Uruguay gas demand by sector", "million m3/day", "stacked_bar")]


def chile_imports(p):
    d = by_date(read(p, "Gas imports"), "Month")
    return [spec("Imports", d[["Imports_mcm_per_day_approx"]].rename(columns={"Imports_mcm_per_day_approx": "Gas imports"}),
                 "Chile natural gas imports (LNG + Argentina)", "million m3/day (approx)", "stacked_bar")]


def power_mix(d):
    """Ember generation-by-type sheet (GWh columns) -> chart frame with a fixed fuel order."""
    g = pd.DataFrame({"Hydro": d.get("Hydro_GWh"), "Gas": d.get("Gas_GWh"), "Wind": d.get("Wind_GWh"),
                      "Solar": d.get("Solar_GWh"), "Coal": d.get("Coal_GWh"), "Nuclear": d.get("Nuclear_GWh"),
                      "Other": d[cols(d, "Bioenergy_GWh", "Other Fossil_GWh", "Other Renewables_GWh")].sum(axis=1)})
    # every chart carries the full fuel list (zeros where a country has none) so each fuel keeps its colour
    return g.fillna(0)


# Standard layout for raw grid-operator generation workbooks (one per country):
#   sheet "Daily": date, Hydro_MWh, Gas_MWh, Wind_MWh, Solar_MWh, Coal_MWh, Nuclear_MWh, Oil_MWh,
#                  Bioenergy_MWh, Other_MWh, Total_MWh   (MWh per day; absent fuels may be omitted)
# Charted as monthly GWh with the same fuel order/colours as power_mix.
def power_daily(title):
    def f(p):
        d = by_date(read(p, "Daily"), "date")
        m = d[[c for c in d.columns if str(c).endswith("_MWh") and c != "Total_MWh"]].resample("MS").sum(min_count=1) / 1000
        m = m[m.index >= "2021-01-01"].rename(columns=lambda c: c.replace("_MWh", "_GWh"))
        m = m.rename(columns={"Oil_GWh": "Other Fossil_GWh", "Other_GWh": "Other Renewables_GWh"})
        return [spec("Generation", power_mix(m), title, "GWh per month", "stacked_bar")]
    return f


def chile_power(p):
    d = by_date(read(p, "Generation by type"), "Month")
    return [spec("Generation", power_mix(d), "Chile power generation by type", "GWh per month", "stacked_bar")]


def sa_power(p):
    """One chart per country sheet of the Ember South America workbook."""
    out = []
    for sheet in pd.ExcelFile(p).sheet_names:
        if sheet.lower() in ("units", "notes") or sheet.startswith("Chart"):
            continue
        d = by_date(read(p, sheet), "Month")
        out.append(spec(sheet, power_mix(d), f"{sheet} power generation by type", "GWh per month", "stacked_bar"))
    return out


def colombia(p):
    d = by_date(read(p, "Demand by sector"), "Month")
    z = lambda *c: d[cols(d, *c)].sum(axis=1, min_count=1)  # noqa: E731
    g = pd.DataFrame({"Power": z("Power"), "Industrial (incl. oil sector)": z("Industrial", "Oil_sector"),
                      "Residential & commercial": z("Residential", "Commercial"),
                      "Refinery & petrochemical": z("Refinery", "Petrochemical"),
                      "Vehicle CNG & compressors": z("Vehicle_CNG", "Compressors")})
    return [spec("Demand", g, "Colombia gas demand by sector", "GBTUD", "stacked_bar")]


def ecuador(p):
    d = by_date(read(p, "Gas by use"), "Month")
    g = pd.DataFrame({"Power (Machala)": d.get("Power_MMBtu_per_day"),
                      "Industry (Bajo Alto LNG)": d.get("Industrial_LNG_MMBtu_per_day")})
    return [spec("Use", g, "Ecuador domestic gas (Amistad) by use", "MMBtu/day", "stacked_bar")]


def trinidad(p):
    out = []
    u = read(p, "Utilization by sector")
    u = u[u["sector"].astype(str).str.upper() != "TOTAL"]
    w = u.pivot_table(index=pd.to_datetime(u["date"]), columns="sector", values="mmscfd", aggfunc="sum")
    pick = lambda pat: w[[c for c in w.columns if re.search(pat, c, re.I)]].sum(axis=1, min_count=1)  # noqa: E731
    g = pd.DataFrame({"LNG": pick(r"^LNG"), "Methanol": pick("methanol"), "Ammonia & derivatives": pick("ammonia|urea"),
                      "Power": pick("power")})
    g["Other"] = w.sum(axis=1) - g.sum(axis=1)
    out.append(spec("Use", g, "Trinidad & Tobago gas use by sector", "MMscf/d", "stacked_bar"))
    pr = read(p, "Production by company")
    pr = pr[pr["company"].astype(str).str.upper() != "TOTAL"].copy()
    pr["company"] = pr["company"].astype(str).str.strip().str.upper().replace(
        {"BHP": "PERENCO (ex-BHP/Woodside)", "WOODSIDE": "PERENCO (ex-BHP/Woodside)", "PEREN": "PERENCO (ex-BHP/Woodside)",
         "PERENCO": "PERENCO (ex-BHP/Woodside)"})
    pw = pr.pivot_table(index=pd.to_datetime(pr["date"]), columns="company", values="mmscfd", aggfunc="sum")
    out.append(spec("Production", pw, "Trinidad & Tobago gas production by company", "MMscf/d", "stacked_bar"))
    return out


def ireland_demand(p):
    d = by_date(read(p, "Data"), "date")
    s = d[cols(d, "PowerGen_GWh", "LDM_ex_PowerGen_GWh", "DM_GWh", "NDM_GWh")]
    names = {"PowerGen": "Power", "LDM_ex_PowerGen": "Large industry (ex power)", "DM": "Daily metered",
             "NDM": "Non-daily metered (homes, small business)"}
    return [spec("Demand", monthly_mean(s, "2021-01-01").rename(columns=lambda c: names.get(c.replace("_GWh", ""), c)),
                 "Ireland gas demand by sector (monthly average)", "GWh/day", "stacked_bar"),
            {"name": "Seasonal", "water_year": d["Total_ROI_GWh"], "title": "Ireland total gas demand", "units": "GWh/day"}]


def ireland_combined(p):
    names = {"PowerGen": "Power", "LDM_ex_PowerGen": "Large industry (ex power)", "DM": "Daily metered",
             "NDM": "Non-daily metered (homes, small business)"}
    d = by_date(read(p, "Demand"), "date")
    dem = d[cols(d, "PowerGen_GWh", "LDM_ex_PowerGen_GWh", "DM_GWh", "NDM_GWh")]
    s = by_date(read(p, "Supply"), "date")
    sup = s[cols(s, "Corrib_Production_GWh", "Inch_Production_GWh", "Moffat_Imports_GWh")]
    return [spec("Demand", monthly_mean(dem, "2021-01-01").rename(columns=lambda c: names.get(c.replace("_GWh", ""), c)),
                 "Ireland gas demand by sector (monthly average)", "GWh/day", "stacked_bar"),
            spec("Supply", monthly_mean(sup, "2021-01-01").rename(
                     columns={"Corrib_Production_GWh": "Corrib", "Inch_Production_GWh": "Inch",
                              "Moffat_Imports_GWh": "Moffat imports (UK)"}),
                 "Ireland gas supply by source (monthly average)", "GWh/day", "stacked_bar"),
            {"name": "Seasonal", "water_year": d["Total_ROI_GWh"], "title": "Ireland total gas demand", "units": "GWh/day"}]


def ireland_supply(p):
    d = by_date(read(p, "Data"), "date")
    s = d[cols(d, "Corrib_Production_GWh", "Inch_Production_GWh", "Moffat_Imports_GWh")]
    return [spec("Supply", monthly_mean(s, "2021-01-01").rename(columns=lambda c: c.replace("_GWh", "").replace("_", " ")),
                 "Ireland gas supply by source (monthly average)", "GWh/day", "stacked_bar")]


def ireland_gni(p):
    e = by_date(read(p, "Entry flows"), "date")
    c = by_date(read(p, "Consumption by sector"), "date")
    tidy = lambda c: c.replace("_GWh", "").replace("ROI_Power_Gen", "Power").replace("_", " ")  # noqa: E731
    return [spec("Entry flows", e[cols(e, "Bellanaboy_GWh", "Gormanston_GWh", "Moffat_GWh")].rename(columns=tidy),
                 "Ireland gas entry flows (GNI)", "GWh/day", "stacked_area", "%Y-%m-%d"),
            spec("Consumption", c[cols(c, "ROI_Power_Gen_GWh", "DM_GWh", "NDM_GWh")].rename(columns=tidy),
                 "Ireland gas consumption by market sector (GNI)", "GWh/day", "stacked_area", "%Y-%m-%d")]


def ireland_smartgrid(p):
    d = by_date(read(p, "Data"), "datetime")
    dd = weekly_mean(d, "2021-01-01")
    return [spec("Demand and wind", dd[cols(dd, "Demand_Actual_MW", "Wind_Actual_MW")].rename(
                     columns={"Demand_Actual_MW": "Demand", "Wind_Actual_MW": "Wind generation"}),
                 "Ireland electricity demand and wind (weekly average)", "MW", "line", "%Y-%m-%d"),
            spec("CO2 intensity", dd[cols(dd, "CO2_Intensity_gCO2_per_kWh")].rename(
                     columns={"CO2_Intensity_gCO2_per_kWh": "CO2 intensity"}),
                 "Ireland grid CO2 intensity (weekly average)", "gCO2/kWh", "line", "%Y-%m-%d")]


def mexico(p):
    d = by_date(read(p, "Data"), "date")
    w = weekly_mean(d[["Total_MWh"]]).rename(columns={"Total_MWh": "National demand"})
    return [spec("Demand", w, "Mexico national electricity demand (weekly average)", "MWh/day", "line", "%Y-%m-%d")]


def fuel_mix(title, units, drop=("Total", "Imports", "Renewables_Share", "share", "unknown")):
    def f(p):
        d = by_date(read(p, "Data"), "date")
        c = [x for x in d.columns if not any(k.lower() in str(x).lower() for k in drop) and pd.api.types.is_numeric_dtype(d[x])]
        m = d[c].rename(columns=lambda x: str(x).replace("_MW", "").replace("_", " "))
        if (m.index.max() - m.index.min()).days > 400:
            return [spec("Mix", monthly_mean(m), f"{title} (monthly average)", units, "stacked_bar")]
        return [spec("Mix", m, title, units, "stacked_area", "%Y-%m-%d")]
    return f


def puertorico(p):
    d = by_date(read(p, "By fuel type"), "period")
    names = {"NG": "Natural gas", "RFO": "Residual fuel oil", "DFO": "Distillate oil", "COL": "Coal", "SUN": "Solar",
             "HYC": "Hydro", "WND": "Wind", "LFG": "Landfill gas", "BIO": "Biomass"}
    c = [x for x in d.columns if pd.api.types.is_numeric_dtype(d[x]) and not re.search(r"total|all", str(x), re.I)]
    return [spec("By fuel", d[c].rename(columns=lambda x: names.get(x, x)), "Puerto Rico generation by fuel",
                 "GWh per month", "stacked_bar")]


def china_nbs(p):
    d = read(p, "Data")
    dc = next(c for c in d.columns if "month" in str(c).lower() or "date" in str(c).lower())
    d = by_date(d, dc)
    gen = [c for c in d.columns if str(c).endswith("_Generation_GWh") and not str(c).startswith("Total")]
    out = []
    if gen:
        out.append(spec("Generation", d[gen].rename(columns=lambda c: c.replace("_Generation_GWh", "")),
                        "China power generation by source (NBS)", "GWh per month", "stacked_bar"))
    for c in [c for c in d.columns if c not in gen and pd.api.types.is_numeric_dtype(d[c])]:
        label, unit = label_and_unit(c)
        out.append(spec(label, d[[c]].rename(columns={c: label}), f"China {label.lower()} (NBS)",
                        f"{unit} per month" if unit else "per month", "line"))
    return out


def china_nbs_series(p):
    """China NBS workbooks: the pull writes a 'Series' sheet (column, label, unit, chart, kind); one chart per group."""
    d = read(p, "Data")
    d = by_date(d, d.columns[0])
    s = pd.read_excel(p, sheet_name="Series", index_col=0)
    out, used = [], set()
    for group in dict.fromkeys(s["chart"].dropna()):
        rows = s[s["chart"] == group]
        c = [x for x in rows.index if x in d.columns and d[x].notna().any()]
        if str(group).strip() and c:
            name = re.sub(r"[\\/*?:\[\]]", "-", str(group))[:23]   # sheet = "Chart - " + name, max 31 chars
            name = name if name not in used else f"{name[:20]} {len(used)}"
            used.add(name)
            out.append(spec(name, d[c].rename(columns=rows["label"].to_dict()),
                            f"China {group} (NBS)", rows["unit"].iloc[0], rows["kind"].iloc[0]))
    return out


def giignl(p):
    d = read(p, "Data")
    d = d.set_index(pd.to_datetime(d["report_year"].astype(int).astype(str) + "-01-01"))
    exporters = [c for c in d.columns if c.endswith("_Total_MT") and "Basin" not in c and "Middle East" not in c]
    top = d[exporters].mean().sort_values(ascending=False).index[:7]
    t = d[top].rename(columns=lambda c: c.replace("_Total_MT", ""))
    return [spec("Exports", t, "LNG exports by country (GIIGNL)", "million tonnes per year", "stacked_bar", "%Y")]


def south_africa(p):
    return fuel_mix("South Africa generation mix (Eskom)", "MW (daily mean)")(p)


def henry_hub(p):
    d = by_date(read(p, "Data"), "date")[["Henry_Hub_USD_per_MMBtu"]].rename(
        columns={"Henry_Hub_USD_per_MMBtu": "Henry Hub spot"})
    return [spec("Daily", daily(d, "2021-01-01"), "Henry Hub natural gas spot price (daily)", "USD/MMBtu"),
            spec("Monthly", monthly_mean(d), "Henry Hub natural gas spot price (monthly average)", "USD/MMBtu")]


def _sheet(p, name, date_col):
    try:
        return by_date(read(p, name), date_col)
    except ValueError:   # sheet not in this workbook
        return pd.DataFrame()


def _full_years(d):
    return d[d["Coverage"].astype(str).eq("Full year")] if "Coverage" in d else d


def singapore_power(p):
    out = []
    d = _sheet(p, "Daily demand", "Date")
    if not d.empty:
        dm = monthly_mean(d[cols(d, "System_Demand_Avg_MW", "System_Demand_Peak_MW")], "2021-01-01")
        out.append(spec("Demand", dm.rename(columns={"System_Demand_Avg_MW": "Average system demand",
                                                     "System_Demand_Peak_MW": "Daily peak"}),
                        "Singapore electricity system demand (monthly average of daily values)", "MW"))
    g = _sheet(p, "Daily generation by type", "Date")
    if not g.empty:
        gc = [c for c in g.columns if str(c).endswith("_GWh") and not str(c).startswith("Total")]
        out.append(spec("Generation", monthly_mean(g[gc].clip(lower=0), "2021-01-01").rename(
                            columns=lambda c: c.replace("_GWh", "").replace("_", " ")),
                        "Singapore metered generation by plant type (monthly average)", "GWh/day", "stacked_bar"))
    m = _sheet(p, "Monthly generation", "Month")
    if not m.empty:
        out.append(spec("Monthly generation", m.loc[m.index >= "2015-01-01", ["Electricity_Generation_GWh"]].rename(
                            columns={"Electricity_Generation_GWh": "Electricity generation"}),
                        "Singapore electricity generation (SingStat/EMA)", "GWh per month"))
    c = _full_years(_sheet(p, "Annual consumption", "Year"))
    if not c.empty:
        cc = [x for x in c.columns if str(x).endswith("_GWh") and not str(x).startswith("Total")]
        out.append(spec("Consumption", c[cc].rename(columns=lambda x: x.replace("_GWh", "").replace("_", " ")),
                        "Singapore electricity consumption by sector", "GWh per year", "stacked_bar", "%Y"))
    f = _full_years(_sheet(p, "Annual fuel mix", "Year"))
    if not f.empty:
        fc = [x for x in f.columns if str(x).endswith("_pct")]
        out.append(spec("Fuel mix", f[fc].rename(columns=lambda x: x.replace("_pct", "").replace("_", " ")),
                        "Singapore fuel mix for electricity generation", "% of generation", "stacked_bar", "%Y"))
    return out


def singapore_gas(p):
    out = []
    d = _full_years(_sheet(p, "Annual demand by sector", "Year"))
    if not d.empty:
        dc = cols(d, "Power_generation_TJ", "Industrial_TJ", "Commerce_Services_TJ", "Households_TJ", "Transport_TJ",
                  "Others_TJ")
        out.append(spec("Demand", (d[dc] / 1000).rename(columns=lambda x: x.replace("_TJ", "").replace("_", " ")),
                        "Singapore natural gas demand by sector", "PJ per year", "stacked_bar", "%Y"))
    i = _sheet(p, "Annual imports", "Year")
    if not i.empty:
        out.append(spec("Imports", (i[cols(i, "Pipeline_TJ", "LNG_TJ")] / 1000).rename(
                            columns={"Pipeline_TJ": "Pipeline gas", "LNG_TJ": "LNG"}),
                        "Singapore natural gas imports, pipeline vs LNG", "PJ per year", "stacked_bar", "%Y"))
    e = _sheet(p, "Power burn monthly (est)", "Month")
    if not e.empty:
        out.append(spec("Power burn", e[["Gas_for_power_mcm_per_day_est"]].rename(
                            columns={"Gas_for_power_mcm_per_day_est": "Gas for power (estimate)"}),
                        "Singapore gas burn for power, estimated from metered CCGT generation",
                        "mcm/day (approx)"))
    t = _sheet(p, "Town gas quarterly", "Quarter_start")
    if not t.empty:
        out.append(spec("Town gas", t.loc[t.index >= "2015-01-01", cols(t, "Domestic_GWh", "Non_domestic_GWh")].rename(
                            columns={"Domestic_GWh": "Domestic", "Non_domestic_GWh": "Non-domestic"}),
                        "Singapore town gas sales", "GWh per quarter", "stacked_bar"))
    return out


def brazil_hydro(p):
    d = by_date(read(p, "Daily"), "date")
    areas = [("SIN", "Brazil reservoirs, SIN (national)"),
             ("SE_CO", "Brazil reservoirs, SE/CO"),
             ("S", "Brazil reservoirs, South"),
             ("NE", "Brazil reservoirs, Northeast"),
             ("N", "Brazil reservoirs, North")]
    return [{"name": a, "water_year": d[f"{a}_pct"], "title": t, "units": "% full",
             "sheet": f"Water year - {a.replace('_', '-')}"} for a, t in areas if f"{a}_pct" in d]


def colombia_hydro(p):
    d = by_date(read(p, "Daily"), "date")
    return [{"name": "Storage", "water_year": d["Storage_pct"], "title": "Colombia reservoirs (national)",
             "units": "% full"}]


def generic(p):
    xl = pd.ExcelFile(p)
    for s in xl.sheet_names:
        if s.lower() in ("units", "notes") or s.lower().startswith("chart") or s == water_year_chart.SHEET:
            continue
        d = read(p, s)
        dc = next((c for c in d.columns if re.search(r"date|month|period|day|time", str(c), re.I)), None)
        if dc is None:
            continue
        d = by_date(d, dc)
        num = [c for c in d.columns if pd.api.types.is_numeric_dtype(d[c])][:8]
        if num and len(d) > 1:
            return [spec(s, d[num], f"{os.path.splitext(os.path.basename(p))[0]} - {s}", "", "line")]
    return []


REGISTRY = {
    "argentina_gas_monthly.xlsx": argentina,
    "brazil_gas_monthly.xlsx": brazil,
    "bolivia_gas_demand_by_sector.xlsx": bolivia,
    "uruguay_gas_demand_by_sector.xlsx": uruguay,
    "chile_gas_imports.xlsx": chile_imports,
    "chile_power_by_type.xlsx": chile_power,
    "brazil_power_generation_daily.xlsx": power_daily("Brazil power generation by type (ONS)"),
    "colombia_power_generation_daily.xlsx": power_daily("Colombia power generation by type (XM)"),
    "panama_power_generation_daily.xlsx": power_daily("Panama power generation by type (CND)"),
    "costa_rica_power_generation_daily.xlsx": power_daily("Costa Rica power generation by type (CENCE)"),
    "nicaragua_power_generation_daily.xlsx": power_daily("Nicaragua power generation by type (CNDC)"),
    "south_america_power_by_type.xlsx": sa_power,
    "argentina_power_generation_daily.xlsx": power_daily("Argentina power generation by type (CAMMESA)"),
    "uruguay_power_generation_daily.xlsx": power_daily("Uruguay power generation by type (ADME)"),
    "bolivia_power_generation_daily.xlsx": power_daily("Bolivia power generation by type (CNDC)"),
    "ecuador_power_generation_daily.xlsx": power_daily("Ecuador power generation by type (CENACE)"),
    "peru_power_generation_daily.xlsx": power_daily("Peru power generation by type (COES)"),
    "central_america_power_by_type.xlsx": sa_power,
    "guatemala_power_generation_daily.xlsx": power_daily("Guatemala power generation by type (AMM)"),
    "colombia_gas_demand_by_sector.xlsx": colombia,
    "ecuador_gas.xlsx": ecuador,
    "trinidad_gas.xlsx": trinidad,
    "ireland_gas_demand_daily.xlsx": ireland_demand,
    "ireland_gas_supply_daily.xlsx": ireland_supply,
    "ireland_gas_combined_daily.xlsx": ireland_combined,
    "ireland_gni_transparency_daily.xlsx": ireland_gni,
    "ireland_smartgrid_15min.xlsx": ireland_smartgrid,
    "mexico_demanda_nacional_daily.xlsx": mexico,
    "miso_fuel_mix_daily.xlsx": fuel_mix("MISO fuel mix", "MW (daily mean)"),
    "turkey_generation_mix_dashboard_daily.xlsx": fuel_mix("Turkey generation mix", "MW"),
    "south_africa_generation_mix_daily.xlsx": south_africa,
    "puertorico_generation_monthly.xlsx": puertorico,
    "china_nbs_clean_energy_products_monthly.xlsx": china_nbs_series,
    "china_nbs_energy_production_monthly.xlsx": china_nbs_series,
    "china_nbs_industrial_output_monthly.xlsx": china_nbs_series,
    "china_nbs_market_prices_10day.xlsx": china_nbs_series,
    "china_nbs_capacity_utilization_quarterly.xlsx": china_nbs_series,
    "china_nbs_ppi_monthly.xlsx": china_nbs_series,
    "giignl_contracted_vs_spot_annual.xlsx": giignl,
    "singapore_power.xlsx": singapore_power,
    "singapore_gas.xlsx": singapore_gas,
    "henry_hub_daily.xlsx": henry_hub,
    "brazil_hydro_reservoirs.xlsx": brazil_hydro,
    "colombia_hydro_reservoirs.xlsx": colombia_hydro,
    # these build their own charts in their pull scripts:
    "rhine_kaub_level_daily.xlsx": None,
    "gatun_lake_level.xlsx": None,
    "eia930_fuel_mix_daily.xlsx": None,
    "lng_feedgas_daily.xlsx": None,
}


def _drop_old_chart_sheets(path):
    """Charts are rebuilt from scratch each run; remove last run's chart sheets so renamed/retired ones don't linger."""
    from openpyxl import load_workbook
    wb = load_workbook(path)
    old = [n for n in wb.sheetnames if n == "Chart" or n.startswith("Chart - ")]
    if old:
        for n in old:
            del wb[n]
        root, ext = os.path.splitext(path)
        tmp = f"{root}.tmp{os.getpid()}{ext}"
        try:
            wb.save(tmp)
            os.replace(tmp, path)
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)


def _order_chart_sheets(path, names):
    """Chart sheets straight after the Units tab, in registry order (each add inserts at position 1)."""
    from openpyxl import load_workbook
    wb = load_workbook(path)
    present = [n for n in names if n in wb.sheetnames]
    if len(present) < 2:
        return
    sheets = {ws.title: ws for ws in wb._sheets}
    rest = [ws for ws in wb._sheets if ws.title not in present]
    head = rest[:1] if rest and rest[0].title.lower() in ("units", "notes") else []
    wb._sheets = head + [sheets[n] for n in present] + [ws for ws in rest if ws not in head]
    root, ext = os.path.splitext(path)
    tmp = f"{root}.tmp{os.getpid()}{ext}"
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def add_charts(path):
    name = os.path.basename(path)
    fn = REGISTRY.get(name, generic)
    if fn is None:
        print(f"{name}: charts are built by its own pull script")
        return 0
    specs = fn(path)
    _drop_old_chart_sheets(path)
    for i, s in enumerate(specs):
        sheet = "Chart" if i == 0 else f"Chart - {s['name']}"[:31]
        if "water_year" in s:
            water_year_chart.add_water_year_chart(path, s["water_year"], s["title"], s["units"],
                                                  sheet_name=s.get("sheet", water_year_chart.SHEET))
            continue
        df = s["df"].dropna(how="all")
        if df.empty:
            continue
        xlsx_charts.add_chart_sheet(path, df, s["title"], s["units"], kind=s["kind"], sheet_name=sheet,
                                    date_format=s["date_format"], line_cols=s.get("line_cols", ()))
    # one name per spec, in registry order; an optional "sheet" key names a water-year sheet
    # (several water-year charts in one workbook)
    _order_chart_sheets(path, [sp.get("sheet", water_year_chart.SHEET) if "water_year" in sp
                               else ("Chart" if i == 0 else f"Chart - {sp['name']}"[:31])
                               for i, sp in enumerate(specs)])
    print(f"{name}: {len(specs)} chart(s)")
    return len(specs)


if __name__ == "__main__":
    for p in sys.argv[1:]:
        if os.path.exists(p):
            add_charts(p)
        else:
            print(f"{p}: not found, skipped")
