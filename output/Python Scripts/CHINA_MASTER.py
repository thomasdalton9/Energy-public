"""
China master workbook: every China dataset this repo pulls in one file, with Dashboard front pages carrying all their
charts - the same layout and chart code as the other masters (south_america/SOUTH_AMERICA_MASTER.py).

  Dashboard - Power      power generation by source (average GW), total generation, solar cell / generator-set output;
                         NEA installed capacity by type (GW), NEA electricity consumption by sector (TWh) and y/y,
                         capacity factors (NBS generation / NEA capacity)
  Dashboard - Fuels      raw coal, coke, crude oil and refinery runs, natural gas output (NBS monthly)
  Dashboard - Industry   energy-intensive output (steel, cement, glass, non-ferrous, chemicals, vehicles) and capacity
                         utilisation by industry
  Dashboard - Prices     10-day producer-goods prices (coal, coke, LNG, fuels, steel, metals, chemicals, solar and
                         battery materials, building materials) in US$ per tonne (US$ per kg for polysilicon and
                         live hogs), converted from NBS's yuan at the Federal Reserve H.10 yuan/US$ rate of each price
                         date (last published rate on or before it), and PPI y/y by industry
  Dashboard - Long-term  annual summary and history from 2000 (fundamentals.py)
  <chart> data           the table each Dashboard chart plots
  <dataset> raw          the full data sheet(s) from each source workbook
  Sources                where each dataset comes from, units and notes, and the sources checked but not reachable

All series are China's National Bureau of Statistics (NBS) releases (asia/CHINA_NBS_*.py), the national statistics
office. NBS output covers industrial enterprises above designated size, so distributed solar and small plants are not
in the generation split, and 'Thermal' is NBS's own category (coal, gas, oil, biomass and waste). January and February
are published only combined, so the monthly series start in March 2021 and every January/February is a gap (the
combined figures are on the 'Jan-Feb' raw tabs). Ember is used only for the annual Long-term page.

Gas: only natural gas OUTPUT (NBS) is built. Gas consumption, imports (pipeline and LNG), storage and sector demand
have no official source reachable from GitHub Actions (see the Sources tab), so there is no gas balance and nothing
is estimated.

Reads (doesn't refetch) the workbooks the scheduled pulls write to "output/Data and Chart Outputs/". A missing input
is listed on its Dashboard and skipped. To add a dataset: append it to the lists below plus a SOURCES entry.

Usage: python3 china/CHINA_MASTER.py [--out "output/Data and Chart Outputs/Master Outputs/china_master.xlsx"]
"""
import argparse
import os
import sys
from types import SimpleNamespace

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
import add_charts  # noqa: E402
import fundamentals  # noqa: E402
import xlsx_charts  # noqa: E402
import SOUTH_AMERICA_MASTER as sam  # noqa: E402

DATA_DIR = sam.DATA_DIR

PROD = "china_nbs_energy_production_monthly.xlsx"
IND = "china_nbs_industrial_output_monthly.xlsx"
GAS = "china_ndrc_gas_monthly.xlsx"
IMPORTS = "china_gacc_energy_imports_monthly.xlsx"
PRICES = "china_nbs_market_prices_10day.xlsx"
PPI = "china_nbs_ppi_monthly.xlsx"
FX = "china_fx_usd_daily.xlsx"
CAPU = "china_nbs_capacity_utilization_quarterly.xlsx"
CAP = "china_nea_capacity_monthly.xlsx"
CONS = "china_nea_consumption_monthly.xlsx"
# china_nbs_clean_energy_products_monthly.xlsx is a subset of the industrial-output workbook (same Data columns), so
# it is not read again.

# (dashboard, [(code, country, workbook, raw sheets, short name, charts to show (None = all))])
POWER = [("CN", "China", PROD, ("Data", "Jan-Feb"), "energy production",
          {"power generation by sou", "electricity generation "}),
         ("CN", "China", IND, (), "industrial output", {"solar cell and power eq"}),
         ("CN", "China", CAP, ("Data", "Releases"), "NEA capacity", {"installed capacity by t"}),
         ("CN", "China", CONS, ("Data", "Jan-Feb", "YTD", "Releases"), "NEA consumption",
          {"electricity use by sect", "electricity use growth"})]
FUELS = [("CN", "China", PROD, (), "energy production",
          {"raw coal output", "coke output", "crude oil output and re"}),
         ("CN", "China", IMPORTS, ("Data",), "energy imports", {"crude oil and product i", "coal imports"})]
GAS_SETS = [("CN", "China", GAS, ("Data", "Jan-Feb"), "gas consumption", {"gas consumption y-y"}),
            ("CN", "China", PROD, (), "energy production", {"natural gas output"}),
            ("CN", "China", IMPORTS, (), "energy imports", {"natural gas imports"})]
INDUSTRY = [("CN", "China", IND, ("Data", "Jan-Feb"), "industrial output",
             {"iron and steel output", "cement output", "plate glass output", "non-ferrous metals outp",
              "chemicals output", "vehicle output", "cloth output", "machine tool output",
              "industrial robot output", "service robot output", "electronics output",
              "integrated circuit outp"}),
            ("CN", "China", CAPU, ("Data",), "capacity utilisation", None)]
# NBS 10-day prices are converted to US$ by price_usd_charts() below (charts named here); PPI is plain collect().
PRICE_CHARTS = {"coal and coke prices", "oil and gas product pri", "steel prices", "non-ferrous metal price",
                "basic chemical prices", "polysilicon prices", "lithium iron phosphate ", "polymer and fibre price",
                "building material price", "fertiliser and agrochem", "farm product prices", "live hog prices",
                "forest product prices"}
PRICE = [("CN", "China", PPI, ("Data",), "PPI", None)]
FED_URL = "https://www.federalreserve.gov/releases/h10/hist/dat00_ch.htm"

NBS = "National Bureau of Statistics of China (NBS)"
SOURCES = {
    CAP: ("National Energy Administration (NEA), monthly 'national power industry statistics' releases (全国电力工业统计数据; "
          "table read as HTML, Word attachment or, where NEA posts a picture, by OCR with sum and text checks): installed "
          "capacity by type", "https://www.nea.gov.cn/xwfb/"),
    CONS: ("National Energy Administration (NEA), monthly 'electricity consumption of the whole society' releases "
           "(全社会用电量): consumption by sector", "https://www.nea.gov.cn/xwfb/"),
    IMPORTS: ("General Administration of Customs of China (GACC), monthly bulletin table (14) Major Import Commodities in "
              "Quantity and Value (English site)", "http://english.customs.gov.cn/Statistics/Statistics?ColumnId=2"),
    GAS: ("National Development and Reform Commission (NDRC), Operation Bureau, national natural gas operation "
          "bulletin (全国天然气运行快报): apparent consumption", "https://www.ndrc.gov.cn/fggz/jjyxtj/"),
    PROD: (f"{NBS}, monthly 'industrial added value' release, table of output of major industrial products "
           "(industrial enterprises above designated size)", "https://www.stats.gov.cn/sj/zxfb/"),
    IND: (f"{NBS}, monthly 'industrial added value' release, table of output of major industrial products",
          "https://www.stats.gov.cn/sj/zxfb/"),
    PRICES: (f"{NBS}, 10-day market prices of important means of production in the circulation sector",
             "https://www.stats.gov.cn/sj/zxfb/"),
    FX: ("Board of Governors of the Federal Reserve System, H.10 Foreign Exchange Rates: noon buying rates in New York, "
         "Chinese renminbi per US dollar (the series FRED republishes as DEXCHUS); used to convert the NBS 10-day prices to US$",
         FED_URL),
    PPI: (f"{NBS}, monthly producer price (PPI) release, by industry", "https://www.stats.gov.cn/sj/zxfb/"),
    CAPU: (f"{NBS}, quarterly industrial capacity utilisation release", "https://www.stats.gov.cn/sj/zxfb/"),
}


def ppi_specs(path):
    """The PPI workbook's own chart groups, plus the series it leaves unplotted (consumer goods, light and equipment
    industries, other purchaser materials) and the headline month-on-month series, in groups of at most 8 lines."""
    specs = add_charts.china_nbs_series(path)
    d = add_charts.by_date(add_charts.read(path, "Data"), "month")
    lab = pd.read_excel(path, sheet_name="Series", index_col=0)["label"].to_dict()
    groups = [
        ("PPI consumer goods y/y", "% y/y", ["PPI_Food", "PPI_Clothing", "PPI_Daily_Goods", "PPI_Durables"], "_YoY_pct"),
        ("PPI other purchaser prices", "% y/y", ["PPIRM_Timber_Pulp", "PPIRM_Other_Materials", "PPIRM_Farm_Products",
                                                 "PPIRM_Textile_Materials"], "_YoY_pct"),
        ("PPI food and textile ind.", "% y/y", ["Agri_Food_Processing", "Food_Manufacturing", "Beverages", "Tobacco",
                                                "Textiles", "Apparel"], "_YoY_pct"),
        ("PPI light industries y/y", "% y/y", ["Wood_Products", "Paper", "Printing", "Pharmaceuticals", "Rubber_Plastics",
                                               "Non_Metallic_Mining", "Water_Supply"], "_YoY_pct"),
        ("PPI equipment industries", "% y/y", ["Metal_Products", "General_Equipment", "Automobiles",
                                               "Other_Transport_Equipment", "Electronics"], "_YoY_pct"),
        ("PPI headline m/m", "% m/m", ["PPI", "PPI_Producer_Goods", "PPI_Mining", "PPI_Raw_Materials", "PPI_Processing",
                                       "PPI_Consumer_Goods", "PPIRM"], "_MoM_pct"),
    ]
    for name, unit, stems, suffix in groups:
        cols = [st + suffix for st in stems if st + suffix in d.columns and d[st + suffix].notna().any()]
        if cols:
            title = {"PPI headline m/m": "China PPI headline m/m (NBS)"}.get(name, f"China {name} (NBS)")
            specs.append(add_charts.spec(name, d[cols].rename(columns=lab), title, unit, "line"))
    return specs


MASTER_SPECS = {PPI: ppi_specs}
HYDRO_DATASETS = []
HYDRO_EXTRA = {}
EMBER = set()
OPERATORS = {}
DASHBOARD_ONLY = {}

# Sources checked and NOT built (filled in from the probe runs, discovery_archive/china/ and results/china/).
NOT_AVAILABLE = [
    ("China", "Natural gas: production, imports, sector demand, storage, LNG terminals",
     "No official source reachable beyond the three built here", "Not built",
     "https://www.ndrc.gov.cn/fggz/jjyxtj/",
     "NDRC's bulletin gives apparent consumption only (no production, import or sector split); GACC gives imports as "
     "weight (tonnes), not volume. No gas balance is built and nothing is estimated or converted."),
    ("China", "NEA monthly capacity and consumption: what is NOT in the series",
     "National Energy Administration (NEA) releases, now built (china_nea_capacity_monthly.xlsx, china_nea_consumption_monthly.xlsx)",
     "Built, with gaps",
     "https://www.nea.gov.cn/xwfb/",
     "The list page is script-rendered, but the ds_*.json behind it holds the whole list, so the series run from Dec 2020 "
     "(capacity) and Mar 2021 (consumption). Gaps are shown, never filled: capacity has no January (NEA publishes "
     "January-February together) and no end-April 2024 or end-March 2026 stock (no release in NEA's list); consumption has "
     "no January/February month (combined figure on the Jan-Feb tab; 2021-23 and 2025 releases also give February alone), "
     "no December month (only the full year is released), and no May 2023 or May 2026 (no release in NEA's list). "
     "NEA's 2023 monthly capacity rows leave up to 26 GW outside the five types (shown as 'Other')."),
    ("China", "Power industry statistics (China Electricity Council)", "cec.org.cn", "Not reachable",
     "https://www.cec.org.cn/", "www.cec.org.cn timed out from GitHub Actions (connect timeout); english.cec.org.cn "
     "answers with a near-empty page."),
    ("China", "Customs trade statistics, Chinese site / query engine",
     "customs.gov.cn, stats.customs.gov.cn", "Not reachable",
     "http://stats.customs.gov.cn/",
     "HTTP 412 (JavaScript challenge) and a TLS certificate error from GitHub Actions. The English site works and is "
     "used (table 14)."),
    ("China", "NBS data portal (data.stats.gov.cn)", "National Bureau of Statistics", "Not reachable",
     "https://data.stats.gov.cn/", "HTTP 403 'UrlACL' for every automated request (discovery_archive/asia/"
     "CHINA_NBS_DISCOVERY*.py); the press releases on stats.gov.cn are used instead. The statistical yearbook pages "
     "(stats.gov.cn/sj/ndsj/) answer but are annual and are not read."),
    ("China", "CNPC ETRI, CNOOC, CCTD", "company / trade-body sites", "Not used",
     "https://www.cctd.com.cn/", "ETRI answered HTTP 504; CNOOC an empty page; CCTD's home page loads but its coal data "
     "is a commercial database."),
]


def collect(wb, datasets, data_dir, used, sources):
    """sam.collect per dataset so each can show a chosen subset of its charts (names = spec names; None = all)."""
    charts, rows, missing = [], [], []
    for code, country, fname, raw, short, only in datasets:
        cfg = SimpleNamespace(**{k: getattr(sys.modules[__name__], k) for k in (
            "SOURCES", "MASTER_SPECS", "HYDRO_DATASETS", "HYDRO_EXTRA", "EMBER", "OPERATORS")},
            DASHBOARD_ONLY={fname: only} if only else {}, GAS_BCFD=True)   # gas volumes in Bcf/d (owner, Oct 2026)
        before = set(wb.sheetnames)
        c, r, m = sam.collect(wb, [(code, country, fname, raw, short)], data_dir, used, sources, cfg=cfg)
        for ws in wb.worksheets:   # 'CN Data raw 2' -> 'CN industrial output Data raw': say which dataset it is
            if ws.title not in before and ws.title.endswith(" raw") or (ws.title not in before and " raw " in ws.title):
                sheet = ws.title.split(" ", 1)[1].rsplit(" raw", 1)[0]
                used.discard(ws.title)
                ws.title = sam.sheet_name(f"CN {short} {sheet} raw", used)
        charts += c
        rows += r
        missing += m
    return charts, rows, missing


def gas_combined(wb, used, data_dir):
    """Apparent consumption (NDRC) next to NBS output on one chart, over the months NDRC covers. The gap between the
    two is imports plus the output NBS does not count; it is not computed, because neither bulletin gives imports."""
    try:
        cons = add_charts.by_date(add_charts.read(os.path.join(data_dir, GAS), "Data"), "month")["Apparent_Consumption_Bcm"]
        out = add_charts.by_date(add_charts.read(os.path.join(data_dir, PROD), "Data"), "month")["Natural_Gas_Bcm"]
    except Exception as e:  # noqa: BLE001
        return None, f"gas consumption vs output ({type(e).__name__}: {e})"
    df = pd.DataFrame({"Apparent consumption (NDRC)": cons, "Output of enterprises above designated size (NBS)": out})
    df = df[df.iloc[:, 0].notna()]
    title = "China natural gas: apparent consumption (NDRC) and output (NBS)"
    df, gas_units = xlsx_charts.gas_volume_to_bcfd(df, "bcm per month", title)   # bcm per month -> Bcf/d
    df, n_bars = xlsx_charts.prepare(df, tuple(df.columns))
    ws = wb.create_sheet(sam.sheet_name("CN gas consumption vs output data", used))
    xlsx_charts.write_table(ws, df)
    src = ("NDRC national natural gas operation bulletin (consumption); NBS monthly industrial output release (output)",
           "https://www.ndrc.gov.cn/fggz/jjyxtj/")
    chart = xlsx_charts.build_chart(ws, df, n_bars, title, gas_units, "line", width=sam.CHART_W,
                                    height=sam.CHART_H, gridlines=False, inner=xlsx_charts.DASHBOARD_INNER)
    return (chart, src, ("China", title, df.index.max().strftime("%b/%y"), ws.title, *src)), None


def capacity_factor_chart(wb, used, data_dir):
    """NBS generation by type / (NEA installed capacity x hours in the month), months where both exist. NBS covers
    industrial enterprises above designated size (distributed solar and small plants are not in it) while NEA's capacity
    is the whole country's, so solar and wind factors are understated; the end-of-month stock is used, and months either
    series lacks (January/February NBS; January, Apr 2024 and Mar 2026 NEA) are gaps."""
    try:
        cap = add_charts.by_date(add_charts.read(os.path.join(data_dir, CAP), "Data"), "month")
        gen = add_charts.by_date(add_charts.read(os.path.join(data_dir, PROD), "Data"), "month")
    except Exception as e:  # noqa: BLE001
        return None, f"capacity factors ({type(e).__name__}: {e})"
    out = {}
    for t in ("Hydro", "Thermal", "Nuclear", "Wind", "Solar"):
        hours = pd.Series(cap.index.days_in_month * 24.0, index=cap.index)
        g = gen[f"{t}_Generation_TWh"].reindex(cap.index) * 1000.0
        out[t] = (g / (pd.to_numeric(cap[f"{t}_GW"], errors="coerce") * hours) * 100.0).round(1)
    df = pd.DataFrame(out).dropna(how="all")
    df = df[df.index >= "2021-03-01"]
    title = "China capacity factors: NBS generation / NEA installed capacity (NBS covers enterprises above designated size)"
    df, n_bars = xlsx_charts.prepare(df, tuple(df.columns))
    ws = wb.create_sheet(sam.sheet_name("CN capacity factors data", used))
    xlsx_charts.write_table(ws, df)
    src = ("NBS monthly industrial output release (generation); NEA national power industry statistics (capacity)",
           "https://www.nea.gov.cn/xwfb/")
    chart = xlsx_charts.build_chart(ws, df, n_bars, title, "% of installed capacity (monthly)", "line", width=sam.CHART_W,
                                    height=sam.CHART_H, gridlines=False, inner=xlsx_charts.DASHBOARD_INNER)
    return (chart, src, ("China", title, df.index.max().strftime("%b/%y"), ws.title, *src)), None


def fx_for(dates, fx):
    """Rate for each date = the last published rate on or before it (no other filling); returns (rate, rate date)."""
    left = pd.DataFrame({"d": pd.to_datetime(list(dates))}).sort_values("d")
    right = fx.rename("rate").rename_axis("rate_date").reset_index()
    m = pd.merge_asof(left, right, left_on="d", right_on="rate_date", direction="backward").set_index("d")
    return m["rate"], m["rate_date"]


def price_usd_charts(wb, used, data_dir, sources):
    """NBS 10-day prices (yuan per tonne or kg) in US$ per tonne / per kg: price / (yuan per US$), with the Fed H.10 rate
    of the price date (period start: the 1st, 11th or 21st) or, when the Fed published none that day (weekends, US
    holidays), the last rate published before it. Data tabs carry the rate, the date it is from and the yuan prices."""
    charts, rows, missing = [], [], []
    ppath, fpath = os.path.join(data_dir, PRICES), os.path.join(data_dir, FX)
    if not (os.path.exists(ppath) and os.path.exists(fpath)):
        return charts, rows, [f"China market prices in US$ ({PRICES} or {FX} missing)"]
    try:
        fx = add_charts.by_date(add_charts.read(fpath, "Data"), "date")["CNY_per_USD"].dropna()
        specs = [sp for sp in add_charts.china_nbs_series(ppath) if sp["name"] in PRICE_CHARTS]
    except Exception as e:  # noqa: BLE001
        return charts, rows, [f"China market prices in US$ ({type(e).__name__}: {e})"]
    src = (f"{NBS} (10-day circulation-sector prices); converted at the Federal Reserve H.10 yuan per US$ rate", FED_URL)
    for sp in specs:
        yuan = sp["df"].dropna(how="all")
        unit = "US$ per kg" if sp["units"].endswith("/kg") else "US$ per tonne"
        rate, rdate = fx_for(yuan.index, fx)
        usd = yuan.div(rate.reindex(yuan.index), axis=0)
        usd = usd[rate.reindex(yuan.index).notna().values]
        df, n_bars = xlsx_charts.prepare(usd.round(2), ())
        if df.empty:
            continue
        ws = wb.create_sheet(sam.sheet_name(f"CN {sp['name']} US$ data", used))
        xlsx_charts.write_table(ws, df, "%Y-%m")
        c0 = len(df.columns) + 3
        ws.cell(row=1, column=c0, value="Rate date used")
        ws.cell(row=1, column=c0 + 1, value="Yuan per US$ (Fed H.10)")
        for j, c in enumerate(yuan.columns):
            ws.cell(row=1, column=c0 + 2 + j, value=f"{c} ({sp['units']}, as published by NBS)")
        for i, ts in enumerate(df.index, start=2):
            ws.cell(row=i, column=c0, value=rdate[ts].to_pydatetime()).number_format = "dd/mm/yy"
            ws.cell(row=i, column=c0 + 1, value=float(rate[ts]))
            for j, c in enumerate(yuan.columns):
                v = yuan.at[ts, c]
                ws.cell(row=i, column=c0 + 2 + j, value=None if pd.isna(v) else float(v))
        for k in range(c0, c0 + 2 + len(yuan.columns)):
            ws.cell(row=1, column=k).font = Font(bold=True)
            ws.column_dimensions[ws.cell(row=1, column=k).column_letter].width = 24
        title = sp["title"]
        charts.append((xlsx_charts.build_chart(ws, df, n_bars, title, unit, "line", sp["date_format"], width=sam.CHART_W,
                                               height=sam.CHART_H, gridlines=False, inner=xlsx_charts.DASHBOARD_INNER), src))
        rows.append(("China", title, df.index.max().strftime("%b/%y"), ws.title, *src))
    # raw tabs (yuan as published) and the exchange-rate workbook
    sam.write_frame(wb.create_sheet(sam.sheet_name("CN market prices Data raw", used)), add_charts.read(ppath, "Data"))
    sam.write_frame(wb.create_sheet(sam.sheet_name("CN FX Data raw", used)), add_charts.read(fpath, "Data"))
    sources.append(("China", "market prices", PRICES, *SOURCES[PRICES], sam.notes_text(ppath)))
    sources.append(("China", "yuan per US dollar", FX, *SOURCES[FX], sam.notes_text(fpath)))
    ws = wb.create_sheet(sam.sheet_name("Conversion factors", used))
    last = fx.index.max()
    for line in ("Currency conversion of the NBS 10-day prices (the only conversion applied)",
                 "Prices: NBS 10-day market prices of important means of production, yuan per tonne (yuan per kg for polysilicon and live hogs), as published; the yuan series are on the 'CN market prices Data raw' tab.",
                 "US$ price = yuan price / (yuan per US$). Units: US$ per tonne, US$ per kg. No other unit conversion (no energy-content, volume or barrel factors) is made.",
                 "Exchange rate: Board of Governors of the Federal Reserve System, H.10 Foreign Exchange Rates, Historical Rates for the Chinese Renminbi (noon buying rates in New York, yuan per US dollar; the series FRED republishes as DEXCHUS): " + FED_URL,
                 "Rate used for each price: the rate of the price's date (each NBS 10-day period is dated its first day: the 1st, 11th or 21st), or, where the Fed published none that day (weekends, US holidays), the last rate published before it. Nothing else is used to fill days.",
                 "The rate and the date it is from are beside each chart's data on the 'US$ data' tabs, next to the yuan prices.",
                 f"Rate series: {fx.index.min():%d %b %Y} to {last:%d %b %Y}, latest {fx.iloc[-1]:.4f} yuan per US$ (a price dated after the latest Fed date would use that last rate).",
                 "NBS circulation-sector prices are as reported by NBS (tax treatment not stated by NBS).",
                 "Workbook: china_fx_usd_daily.xlsx (asia/CHINA_FX_USD.py, 1st and 15th); an ECB cross-rate column in it is validation only."):
        ws.append([line])
    ws["A1"].font = Font(bold=True, size=14)
    ws.column_dimensions["A"].width = 180
    return charts, rows, missing


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(DATA_DIR, "Master Outputs", "china_master.xlsx"))
    ap.add_argument("--data-dir", default=DATA_DIR)
    args = ap.parse_args()

    wb = Workbook()
    dash_power = wb.active
    dash_power.title = "Dashboard - Power"
    names = ["Dashboard - Gas", "Dashboard - Fuels", "Dashboard - Industry", "Dashboard - Prices"]
    dashes = [wb.create_sheet(n) for n in names]
    used = {"Dashboard - Power", *names, "Dashboard - Long-term", "Sources"}
    sources = []

    sections = [(dash_power, "China - power generation (NBS, average GW)", POWER),
                (dashes[0], "China - natural gas: apparent consumption (NDRC) and output (NBS)", GAS_SETS),
                (dashes[1], "China - fuel output: coal, coke and crude oil (NBS)", FUELS),
                (dashes[2], "China - energy-intensive industrial output and capacity utilisation (NBS)", INDUSTRY),
                (dashes[3], "China - producer-goods prices in US$ (NBS, converted at the Federal Reserve H.10 rate) and PPI (NBS)", PRICE)]
    counts = []
    for dash, heading, datasets in sections:
        charts, rows, missing = collect(wb, datasets, args.data_dir, used, sources)
        if datasets is GAS_SETS:
            both, why = gas_combined(wb, used, args.data_dir)
            if both:
                charts.insert(0, both[:2])
                rows.insert(0, both[2])
            else:
                missing.append(why)
        if datasets is POWER:
            cf, why = capacity_factor_chart(wb, used, args.data_dir)
            if cf:
                charts.append(cf[:2])
                rows.append(cf[2])
            else:
                missing.append(why)
        if datasets is PRICE:
            pc, pr, pm = price_usd_charts(wb, used, args.data_dir, sources)
            charts, rows, missing = pc + charts, pr + rows, pm + missing
        sam.draw_dashboard(dash, heading, charts, rows, missing)
        counts.append(len(charts))
        if missing:
            print(f"missing ({heading}):", missing)

    fundamentals.add_long_term_dashboard(wb, used, sam.sheet_name, "China", args.data_dir, sam.CHART_W, sam.CHART_H,
                                         rows_per_chart=sam.ROWS_PER_CHART, degree_days=("CDD_18", "HDD_18"),
                                         index=sum(s.startswith("Dashboard") for s in wb.sheetnames))

    src = wb.create_sheet("Sources")
    src.append(["Country", "Dataset", "Workbook", "Publisher", "Link", "Units and notes (from the source workbook)"])
    for cell in src[1]:
        cell.font = Font(bold=True)
    seen = set()
    for s in sources:
        if s[2] in seen:   # a workbook read for two dashboards is listed once
            continue
        seen.add(s[2])
        src.append(list(s))
        if s[4]:
            src.cell(row=src.max_row, column=5).hyperlink = s[4]
    if NOT_AVAILABLE:
        src.append([])
        src.append(["Checked, NOT built", "Dataset", "Publisher / page", "Result from GitHub Actions", "Link", "Notes"])
        for cell in src[src.max_row]:
            cell.font = Font(bold=True)
        for row in NOT_AVAILABLE:
            src.append(list(row))
            if row[4]:
                src.cell(row=src.max_row, column=5).hyperlink = row[4]
    for col, w in (("A", 18), ("B", 22), ("C", 44), ("D", 60), ("E", 50), ("F", 160)):
        src.column_dimensions[col].width = w

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_charts.save_atomic(wb, args.out)
    print(f"Saved {args.out}: charts per dashboard {dict(zip(['power', *names], counts))}; tabs {wb.sheetnames}")


if __name__ == "__main__":
    main()
