"""
South Korea master workbook: every South Korea dataset this repo pulls in one file, with Dashboard front pages carrying all
their charts - the same layout and chart code as the other masters (south_america/SOUTH_AMERICA_MASTER.py).

  Dashboard - Power      generation by fuel (average GW, electricity traded on the KPX power market), installed capacity by
                         fuel (GW), capacity factors (generation / capacity), peak and minimum demand (GW), system
                         marginal price (SMP) monthly and daily in US$/MWh, hours each fuel set the SMP
  Dashboard - Gas        LNG imports by origin region (Mt per month) and import unit price (KOGAS), LNG burned by power
                         generators (annual, KPX EPSIS)
  Dashboard - Long-term  annual summary and history from 2000 (fundamentals.py)
  <chart> data           the table each Dashboard chart plots
  <dataset> raw          the full data sheet(s) from each source workbook
  Sources                where each dataset comes from, units and notes, and the sources checked but not built

Power data are Korea Power Exchange (KPX) statistics from EPSIS (south_korea/SOUTH_KOREA_EPSIS.py). The fuel split covers the
electricity traded on the KPX power market (547 TWh in 2025, 92% of the 593.6 TWh EPSIS reports for the business generators;
PPA volumes, 25 TWh in 2025, are not split by fuel in the source and are not charted) and leaves out self-generation and small
behind-the-meter systems (so solar is understated). The SMP is converted from KRW/kWh to US$/MWh
with the Federal Reserve H.10 won per US$ rate (south_korea/KOREA_FX_USD.py): daily SMP at the rate of the day (or the last
rate published before it), monthly SMP at the mean of the month's published rates.

Gas: LNG imports by origin region are the Korea Gas Corporation (KOGAS) file dataset on the Public Data Portal (monthly,
tonnes, about an 8-month lag); LNG burned for power is EPSIS's annual table. KOGAS's own site, KESIS, the customs portals and
Petronet publish no tables reachable from GitHub Actions (see the Sources tab), so there is no sector demand, storage or
terminal-inventory series and no gas balance; nothing is estimated and no energy or volume conversion of LNG is made.

Reads (doesn't refetch) the workbooks the scheduled pulls write to "output/Data and Chart Outputs/". A missing input is listed on
its Dashboard and skipped.

Usage: python3 south_korea/SOUTH_KOREA_MASTER.py [--out "output/Data and Chart Outputs/Master Outputs/south_korea_master.xlsx"]
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

GEN = "south_korea_power_generation_monthly.xlsx"
CAP = "south_korea_power_capacity.xlsx"
SMP = "south_korea_smp.xlsx"
DEM = "south_korea_demand_daily.xlsx"
GAS = "south_korea_gas_power_use.xlsx"
FX = "south_korea_fx_usd_daily.xlsx"

# (code, country, workbook, raw sheets, short name, charts to show (None = all))
POWER = [("KR", "South Korea", GEN, ("Data", "Annual"), "generation", {"Generation"}),
         ("KR", "South Korea", CAP, ("Monthly",), "capacity", None),
         ("KR", "South Korea", DEM, ("Daily",), "demand", None),
         ("KR", "South Korea", SMP, ("Monthly", "Marginal fuel", "Daily"), "SMP", {"Marginal fuel"})]
LNG = "south_korea_lng_imports_monthly.xlsx"
GAS_SETS = [("KR", "South Korea", LNG, ("Data",), "LNG imports", None),
            ("KR", "South Korea", GAS, ("Annual",), "gas use for power", None)]

KPX = "Korea Power Exchange (KPX), Electric Power Statistics Information System (EPSIS)"
FED_URL = "https://www.federalreserve.gov/releases/h10/hist/dat00_ko.htm"
SOURCES = {
    GEN: (f"{KPX}: electricity trading volume by fuel (KPX power market; PPA volume unsplit), monthly; annual generation table",
          "https://epsis.kpx.or.kr/epsisnew/selectEkmaPtdBftChart.do?menuId=040501"),
    CAP: (f"{KPX}: installed capacity by fuel", "https://epsis.kpx.or.kr/epsisnew/selectEkpoBftChart.do?menuId=020100"),
    DEM: (f"{KPX}: power supply and demand results (daily peak and minimum demand)",
          "https://epsis.kpx.or.kr/epsisnew/selectEkgeEpsMepChart.do?menuId=030100"),
    SMP: (f"{KPX}: system marginal price (monthly, hourly) and SMP-setting fuel",
          "https://epsis.kpx.or.kr/epsisnew/selectEkmaSmpSmpChart.do?menuId=040201"),
    GAS: (f"{KPX}: generation fuel consumption (annual)", "https://epsis.kpx.or.kr/epsisnew/selectEkgeFfuChart.do?menuId=060200"),
    LNG: ("Korea Gas Corporation (KOGAS), 'Korea's natural gas imports by continent', Public Data Portal (data.go.kr) file dataset "
          "15088508 (CSV, free download, monthly from 1988)", "https://www.data.go.kr/data/15088508/fileData.do"),
    FX: ("Board of Governors of the Federal Reserve System, H.10 Foreign Exchange Rates: noon buying rates in New York, "
         "South Korean won per US dollar (the series FRED republishes as DEXKOUS); used to convert the SMP to US$", FED_URL),
}
MASTER_SPECS = {}
HYDRO_DATASETS = []
HYDRO_EXTRA = {}
EMBER = set()
OPERATORS = {}
DASHBOARD_ONLY = {}

NOT_AVAILABLE = [
    ("South Korea", "Natural gas / LNG: KOGAS sales and supply by sector, inventory, terminal send-out",
     "Korea Gas Corporation (kogas.or.kr) and its Public Data Portal file datasets", "Site reachable, no data tables; LNG imports by origin built",
     "https://www.kogas.or.kr/site/koGas/1030302000000",
     "kogas.or.kr answers (HTTP 200) but its import/transport, sales and production pages are descriptive text; its boards list "
     "quarterly 'Gas Industry' PDFs and articles only. The free file datasets on data.go.kr give LNG imports by continent (built) "
     "and monthly domestic production (15049906: a count with no unit stated in the file, so not charted) and Japan/China/Taiwan "
     "import prices to Jun 2023 (15117762, not charted). No sector sales, storage or terminal inventory table found "
     "(searches in discovery_archive/results/south_korea/probe13.txt, probe15.txt)."),
    ("South Korea", "Energy statistics monthly / national energy balance (KEEI)", "KESIS, kesis.net", "Reachable, script-rendered",
     "https://www.kesis.net/menu.es?mid=a10101000000",
     "The statistics tables are loaded by script from /stat/list/*.es; plain requests to those endpoints returned an HTTP 500 or "
     "empty answers (discovery_archive/results/south_korea/probe13.txt), and the monthly bulletin is a board of publications "
     "with no attachment links for an anonymous visitor. Not built."),
    ("South Korea", "LNG and crude imports by origin (customs)", "Korea Customs Service UNI-PASS trade statistics (unipass.customs.go.kr), "
     "KITA K-stat (stat.kita.net), data.go.kr", "Reachable, query tools need a browser session or key",
     "https://unipass.customs.go.kr/ets/index_eng.do",
     "UNI-PASS and K-stat answer HTTP 200 but their statistics queries are script-driven (K-stat's query pages return a service-error "
     "page to a plain request); the data.go.kr API route needs an API key. Not built."),
    ("South Korea", "Crude and product imports, stocks, refinery runs", "KNOC Petronet (petronet.co.kr)",
     "Reachable, member service", "https://www.petronet.co.kr/v4/eng/main.jsp",
     "Prices on the home page are public; the import, production and stock tables are inside a menu form (POST) that returns a "
     "page shell without tables, and the site lists them as a fee-based information service. Not built."),
    ("South Korea", "Hourly generation by fuel (to roll up to daily and monthly)", "KPX data portal / data.go.kr, EPSIS",
     "No keyless hourly series for all fuels", "https://www.data.go.kr/data/15069337/fileData.do",
     "Opened from GitHub Actions on 10 Oct 2026 (discovery_archive/results/south_korea/hourly_probe.txt, hourly_probe2.txt): "
     "the open APIs (generation by source, 15106744 / 15113384 / 15142651) need a service key this repo does not hold. The free "
     "file datasets are (a) 15065387 hourly traded volume, TOTAL only (no fuel), 2017-2021, one-off; (b) 15069337 hourly "
     "renewables by fuel (13 renewable categories) with capacity, calendar 2025 only, replaced once a year; (c) 15065269 hourly "
     "solar and wind by region, calendar 2025 only; (d) 15127502 Jeju solar/wind 2019-2023 (zip). None has hourly nuclear, coal, "
     "gas, oil or hydro. EPSIS offers year/month only on the trading-volume page (040501), day only on the demand page "
     "(030100, no fuel), and its real-time page (030300) shows the current day with no history. A rolled-up daily/monthly "
     "generation by fuel series for all fuels therefore cannot be built; the monthly EPSIS series is used."),
    ("South Korea", "Ember monthly electricity data", "ember-energy.org", "Blocked (Cloudflare challenge, HTTP 403)",
     "https://ember-energy.org/data/monthly-electricity-data/",
     "Not needed: KPX provides the raw series. Ember appears only through the Long-term page."),
]


def collect(wb, datasets, data_dir, used, sources):
    """sam.collect per dataset so each can show a chosen subset of its charts (names = spec names; None = all)."""
    charts, rows, missing = [], [], []
    for code, country, fname, raw, short, only in datasets:
        cfg = SimpleNamespace(**{k: getattr(sys.modules[__name__], k) for k in (
            "SOURCES", "MASTER_SPECS", "HYDRO_DATASETS", "HYDRO_EXTRA", "EMBER", "OPERATORS")},
            DASHBOARD_ONLY={fname: only} if only else {}, GAS_BCFD=True)
        before = set(wb.sheetnames)
        c, r, m = sam.collect(wb, [(code, country, fname, raw, short)], data_dir, used, sources, cfg=cfg)
        for ws in wb.worksheets:   # 'KR Monthly raw 2' -> 'KR SMP Monthly raw': say which dataset it is
            if ws.title not in before and (ws.title.endswith(" raw") or " raw " in ws.title):
                sheet = ws.title.split(" ", 1)[1].rsplit(" raw", 1)[0]
                used.discard(ws.title)
                ws.title = sam.sheet_name(f"KR {short} {sheet} raw", used)
        charts += c
        rows += r
        missing += m
    return charts, rows, missing


def fx_for(dates, fx):
    """Rate for each date = the last published rate on or before it (no other filling)."""
    left = pd.DataFrame({"d": pd.to_datetime(list(dates))}).sort_values("d")
    right = fx.rename("rate").rename_axis("rate_date").reset_index()
    m = pd.merge_asof(left, right, left_on="d", right_on="rate_date", direction="backward").set_index("d")
    return m["rate"], m["rate_date"]


def _chart(wb, used, name, df, n_bars, title, units, kind, date_format, src):
    ws = wb.create_sheet(sam.sheet_name(name, used))
    xlsx_charts.write_table(ws, df, date_format)
    return ws, xlsx_charts.build_chart(ws, df, n_bars, title, units, kind, date_format, width=sam.CHART_W, height=sam.CHART_H,
                                       gridlines=False, inner=xlsx_charts.DASHBOARD_INNER)


def smp_usd_charts(wb, used, data_dir, sources):
    """KPX SMP (KRW/kWh) in US$/MWh = KRW per kWh x 1000 / (KRW per US$). Daily: the Fed H.10 rate of the day (or the last
    published before it); monthly: the mean of the month's published daily rates. The data tabs carry the won price, the rate
    and (daily) the date the rate is from."""
    charts, rows, missing = [], [], []
    spath, fpath = os.path.join(data_dir, SMP), os.path.join(data_dir, FX)
    if not (os.path.exists(spath) and os.path.exists(fpath)):
        return charts, rows, [f"South Korea SMP in US$ ({SMP} or {FX} missing)"]
    try:
        fx = add_charts.by_date(add_charts.read(fpath, "Data"), "date")["KRW_per_USD"].dropna()
        mon = add_charts.by_date(add_charts.read(spath, "Monthly"), "month")
        day = add_charts.by_date(add_charts.read(spath, "Daily"), "date")
    except Exception as e:  # noqa: BLE001
        return charts, rows, [f"South Korea SMP in US$ ({type(e).__name__}: {e})"]
    src = (f"{KPX} (SMP); converted at the Federal Reserve H.10 won per US$ rate", FED_URL)
    # monthly
    krw = mon[["SMP_Land_KRW_per_kWh", "SMP_Jeju_KRW_per_kWh"]].dropna(how="all")
    rate_m = fx.resample("MS").mean().reindex(krw.index)
    usd = krw.mul(1000.0).div(rate_m, axis=0).dropna(how="all")
    usd.columns = ["Mainland", "Jeju"]
    df, n = xlsx_charts.prepare(usd.round(2), ())
    title = "South Korea system marginal price, monthly (KPX EPSIS)"
    ws, ch = _chart(wb, used, "KR SMP monthly US$ data", df, n, title, "US$ per MWh", "line", "%Y-%m", src)
    ws.cell(row=1, column=len(df.columns) + 3, value="Won per US$ (mean of month)")
    for i, ts in enumerate(df.index, start=2):
        ws.cell(row=i, column=len(df.columns) + 3, value=float(rate_m[ts]))
    ws.cell(row=1, column=len(df.columns) + 4, value="Mainland SMP (KRW per kWh)")
    for i, ts in enumerate(df.index, start=2):
        v = krw.at[ts, "SMP_Land_KRW_per_kWh"]
        ws.cell(row=i, column=len(df.columns) + 4, value=None if pd.isna(v) else float(v))
    charts.append((ch, src))
    rows.append(("South Korea", title, df.index.max().strftime("%b/%y"), ws.title, *src))
    # daily weighted average
    w = day["Weighted_avg"].dropna()
    rate_d, rdate = fx_for(w.index, fx)
    usd_d = (w * 1000.0 / rate_d.reindex(w.index)).dropna().round(2)
    d_df = pd.DataFrame({"Mainland SMP, daily weighted average": usd_d})
    d_df, n = xlsx_charts.prepare(d_df, ())
    if not d_df.empty:
        title = "South Korea mainland SMP, daily weighted average (KPX EPSIS)"
        ws, ch = _chart(wb, used, "KR SMP daily US$ data", d_df, n, title, "US$ per MWh", "line", "%b/%y", src)
        c0 = 3
        for j, h in enumerate(("Rate date used", "Won per US$ (Fed H.10)", "Weighted average SMP (KRW per kWh)")):
            ws.cell(row=1, column=c0 + j, value=h).font = Font(bold=True)
        for i, ts in enumerate(d_df.index, start=2):
            ws.cell(row=i, column=c0, value=rdate[ts].to_pydatetime()).number_format = "dd/mm/yy"
            ws.cell(row=i, column=c0 + 1, value=float(rate_d[ts]))
            ws.cell(row=i, column=c0 + 2, value=float(w[ts]))
        charts.append((ch, src))
        rows.append(("South Korea", title, d_df.index.max().strftime("%d/%m/%y"), ws.title, *src))
    sam.write_frame(wb.create_sheet(sam.sheet_name("KR FX Data raw", used)), add_charts.read(fpath, "Data"))
    sources.append(("South Korea", "won per US dollar", FX, *SOURCES[FX], sam.notes_text(fpath)))
    ws = wb.create_sheet(sam.sheet_name("Conversion factors", used))
    last = fx.index.max()
    for line in ("Currency conversion of the KPX system marginal price (the only conversion applied)",
                 "SMP as published by KPX: KRW per kWh (south_korea_smp.xlsx; raw tabs 'KR Monthly raw', 'KR Daily raw'). "
                 "US$ per MWh = KRW per kWh x 1,000 / (KRW per US$).",
                 "Exchange rate: Board of Governors of the Federal Reserve System, H.10 Foreign Exchange Rates, Historical Rates for "
                 "the South Korean Won (noon buying rates in New York, won per US dollar; the series FRED republishes as DEXKOUS): " + FED_URL,
                 "Daily chart: the rate of the SMP's day or, where the Fed published none that day (weekends, US holidays), the last "
                 "rate published before it; the rate and the date it is from are beside the data. Monthly chart: the mean of the "
                 "daily rates the Fed published in the month. Nothing else is used to fill days.",
                 f"Rate series: {fx.index.min():%d %b %Y} to {last:%d %b %Y}, latest {fx.iloc[-1]:.2f} won per US$.",
                 "Workbook: south_korea_fx_usd_daily.xlsx (south_korea/KOREA_FX_USD.py, 1st and 15th); the ECB cross-rate column in it "
                 "is validation only.",
                 "Generation: GWh per month traded on the KPX power market, shown as average GW = GWh / hours in the month. The market "
                 "volume by fuel was 92% of the 2025 generation EPSIS reports for the business generators (547 of 593.6 TWh); the "
                 "difference is mainly PPA volumes (25 TWh in 2025, reported to Jun 2026 and not split by fuel, in the "
                 "'KR generation Data raw' tab) and energy not traded through the market. Solar and wind are understated because "
                 "small and behind-the-meter systems are outside the market data.",
                 "Capacity factor = traded generation / (end-of-month installed capacity x hours in the month); the same coverage "
                 "caveat applies, so the solar and wind factors are understated."):
        ws.append([line])
    ws["A1"].font = Font(bold=True, size=14)
    ws.column_dimensions["A"].width = 180
    return charts, rows, missing


def capacity_factor_chart(wb, used, data_dir):
    try:
        cap = add_charts.by_date(add_charts.read(os.path.join(data_dir, CAP), "Monthly"), "date")
        gen = add_charts.by_date(add_charts.read(os.path.join(data_dir, GEN), "Data"), "month")
    except Exception as e:  # noqa: BLE001
        return None, f"capacity factors ({type(e).__name__}: {e})"
    cap = cap[cap.index >= "2018-01-01"]
    hours = pd.Series(cap.index.days_in_month * 24.0, index=cap.index)
    out = {}
    for name, g, c in (("Nuclear", "Nuclear_GWh", "Nuclear_MW"), ("Coal", "Coal_GWh", "Coal_MW"), ("Gas", "Gas_GWh", "Gas_MW"),
                       ("Solar", "Solar_GWh", "Solar_MW"), ("Wind", "Wind_GWh", "Wind_MW")):
        out[name] = (gen[g].reindex(cap.index) * 1000.0 / (pd.to_numeric(cap[c], errors="coerce") * hours) * 100.0).round(1)
    df = pd.DataFrame(out).dropna(how="all")
    df, n = xlsx_charts.prepare(df, tuple(df.columns))
    title = "South Korea capacity factors: KPX market generation / installed capacity (KPX EPSIS)"
    src = (f"{KPX}: generation (trading volume) and installed capacity", SOURCES[CAP][1])
    ws, ch = _chart(wb, used, "KR capacity factors data", df, n, title, "% of installed capacity (monthly)", "line", "%Y-%m", src)
    return (ch, src, ("South Korea", title, df.index.max().strftime("%b/%y"), ws.title, *src)), None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(DATA_DIR, "Master Outputs", "south_korea_master.xlsx"))
    ap.add_argument("--data-dir", default=DATA_DIR)
    args = ap.parse_args()

    wb = Workbook()
    dash_power = wb.active
    dash_power.title = "Dashboard - Power"
    dash_gas = wb.create_sheet("Dashboard - Gas")
    used = {"Dashboard - Power", "Dashboard - Gas", "Dashboard - Long-term", "Sources"}
    sources = []
    counts = []
    for dash, heading, datasets in (
            (dash_power, "South Korea - power: generation, capacity, demand and system marginal price (KPX EPSIS)", POWER),
            (dash_gas, "South Korea - gas: LNG imports (KOGAS) and LNG burned for power (KPX EPSIS)", GAS_SETS)):
        charts, rows, missing = collect(wb, datasets, args.data_dir, used, sources)
        if datasets is POWER:
            cf, why = capacity_factor_chart(wb, used, args.data_dir)
            if cf:
                charts.append(cf[:2])
                rows.append(cf[2])
            else:
                missing.append(why)
            pc, pr, pm = smp_usd_charts(wb, used, args.data_dir, sources)
            charts, rows, missing = charts + pc, rows + pr, missing + pm
        sam.draw_dashboard(dash, heading, charts, rows, missing)
        counts.append(len(charts))
        if missing:
            print(f"missing ({heading}):", missing)

    fundamentals.add_long_term_dashboard(wb, used, sam.sheet_name, "South Korea", args.data_dir, sam.CHART_W, sam.CHART_H,
                                         rows_per_chart=sam.ROWS_PER_CHART, degree_days=("CDD_18", "HDD_18"),
                                         index=sum(s.startswith("Dashboard") for s in wb.sheetnames))

    src = wb.create_sheet("Sources")
    src.append(["Country", "Dataset", "Workbook", "Publisher", "Link", "Units and notes (from the source workbook)"])
    for cell in src[1]:
        cell.font = Font(bold=True)
    seen = set()
    for s in sources:
        if s[2] in seen:
            continue
        seen.add(s[2])
        src.append(list(s))
        if s[4]:
            src.cell(row=src.max_row, column=5).hyperlink = s[4]
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
    print(f"Saved {args.out}: charts per dashboard {dict(zip(['power', 'gas'], counts))}; tabs {wb.sheetnames}")


if __name__ == "__main__":
    main()
