"""
Daily US LNG feedgas by plant, in Bcf/d, from pipeline EBB postings -
scheduled quantities at the meters where interstate pipelines hand gas
to each LNG terminal (or to the terminal's own dedicated lateral).

How the points were found: KM_OPAVAIL_SCAN.py (Kinder Morgan platform -
NGPL, Tennessee Gas, Kinder Morgan Louisiana, Elba Express) and
ENBRIDGE_OA_SCAN.py (Texas Eastern) pulled every point on each
pipeline's "Operationally Available Capacity" posting; the POINTS table
below is the terminal-side meters picked out of those files by name,
checked against FEEDGAS_PIPES_DISCOVERY.py's terminal -> pipeline map.
Only terminal-side meters are used, never pipeline-to-pipeline
interconnects upstream of them (e.g. Texas Eastern's delivery into
Kinder Morgan Louisiana is NOT counted - KMLP's own delivery into Sabine
Pass already includes it), so nothing is double counted.

COVERAGE IS PARTIAL for most plants - see COVERAGE_NOTES, also written
into the workbook's Units tab. Pipes not yet scraped (Trunkline and
Gulf Run on Energy Transfer, Transco on Williams, BIG/Gulf South for
Freeport, Cameron Interstate's own postings) and intrastate pipes that
publish nothing (the Permian lines into Corpus Christi, Venture
Global's laterals) are the gap. Treat plant numbers as "feedgas seen on
covered pipes", a floor rather than the full figure, until noted as
complete.

Gas day: yesterday (US Central), "best available" cycle - by the
morning after, that's the final intraday cycle. The run upserts that
gas day into the archive, so re-running is harmless.

Units: pipelines post Dth (= MMBtu). Bcf/d = Dth / 1,037,000, using
1,037 Btu per cubic foot, EIA's average heat content of US dry gas.

Run by .github/workflows/lng_feedgas_daily.yml every morning; needs
Playwright + Chromium. Output: lng_feedgas_daily.xlsx next to this
script.
"""

import argparse
import io
import os
import re
import sys
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))  # repo root

import xlsx_notes

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Playwright isn't installed: pip install playwright && playwright install chromium", file=sys.stderr)
    sys.exit(1)

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT = os.path.join(HERE, "lng_feedgas_daily.xlsx")
BTU_PER_CF = 1037
DTH_PER_BCF = BTU_PER_CF * 1_000_000 / 1000  # 1 Bcf = 1e9 cf * 1,037 Btu = 1.037e12 Btu = 1.037e6 Dth

# (plant, platform, pipeline, loc id, label). Platforms: "km" = Kinder
# Morgan pipeline code, "enbridge" = Enbridge LINK business unit,
# "gasnom" = gasnom.com pipeline, "et" = Energy Transfer asset code,
# "cheniere" = Cheniere LNG Connection pipeline number (200 Creole
# Trail, 400 Corpus Christi Pipeline), "tc" = TC Energy pipeline (ANR,
# CGT = Columbia Gulf), "bhe" = BHE GT&S pipeline path ("cpl" = Cove Point).
POINTS = [
    ("Sabine Pass", "km", "KMLP", "49448", "KMLP -> SP Liquefaction, Cameron Par."),
    ("Sabine Pass", "km", "NGPL", "46622", "NGPL -> Sabine Pass Liquefaction"),
    # Creole Trail's own delivery meter into the plant. Its receipts
    # include the Texas Eastern (25,000) and Trunkline (259,976) Creole
    # Trail deliveries previously counted here, plus LEAP and Acadian -
    # those upstream meters are dropped so nothing is counted twice.
    ("Sabine Pass", "cheniere", "200", "CT200111", "Creole Trail -> Sabine Pass Liquefaction"),
    ("Sabine Pass", "cheniere", "200", "SPLNGD", "Creole Trail -> Sabine Pass LNG"),
    ("Plaquemines", "km", "TGP", "55833", "Tennessee Gas -> VG Gator Express, Evangeline Pass"),
    ("Plaquemines", "enbridge", "TE", "74530", "Texas Eastern -> Gator Express"),
    # Wilkinson Bayou is in Plaquemines Parish - Columbia Gulf's feed into
    # Gator Express (Venture Global's filings name Columbia Gulf, TGP and
    # Texas Eastern as the plant's supply pipes). Counterparty not yet
    # confirmed from Columbia Gulf's location file (it timed out).
    ("Plaquemines", "tc", "CGT", "4267", "Columbia Gulf -> Wilkinson Bayou (Gator Express)"),
    # Cameron Interstate's own delivery meter into the terminal - its
    # receipts include the Tennessee Gas (TENN-CIP) and Texas Eastern
    # (TETCO-CIP) quantities seen on those pipes, plus LEAP, LEG, NG3 and
    # Gillis Hub, so this one meter is the whole plant.
    ("Cameron", "gasnom", "CAMERON", "772300", "Cameron Interstate -> Cameron LNG"),
    # A second, direct feed: Cameron Interstate's receipts don't include
    # Columbia Gulf, and the two together (~2.0 Bcf/d) match Cameron's
    # EIA-implied feedgas.
    ("Cameron", "tc", "CGT", "4246", "Columbia Gulf -> Cameron LNG"),
    ("Calcasieu Pass", "enbridge", "TE", "74529", "Texas Eastern -> TransCameron, Oak Grove"),
    # ANR's location file: meter 523094 is an interconnect in Cameron
    # Parish whose counterparty is TransCameron Pipeline, LLC (TC Energy's
    # Grand Chenier XPress).
    ("Calcasieu Pass", "tc", "ANR", "523094", "ANR -> TransCameron (Grand Chenier XPress), Mermentau River"),
    # Cheniere Corpus Christi Pipeline's delivery into the plant - its
    # receipts include the intrastate Permian supply (Kinder Morgan
    # Tejas, Enterprise...) as well as the NGPL/TGP Sinton meters
    # previously summed here, which are dropped to avoid double counting.
    ("Corpus Christi", "cheniere", "400", "CC200221", "Corpus Christi Pipeline -> Corpus Christi Liquefaction"),
    ("Freeport", "enbridge", "TE", "79999", "Texas Eastern -> Stratton Ridge"),
    ("Freeport", "enbridge", "TE", "73912", "Texas Eastern -> BIG Pipeline, Angleton"),
    ("Elba Island", "km", "EEC", "660700", "Elba Express -> Elba Liquefaction, Chatham"),
    ("Golden Pass", "et", "GR", "808311", "Gulf Run -> Golden Pass Pipeline"),
    # Cove Point pipeline's delivery into the plant (BHE GT&S posting).
    ("Cove Point", "bhe", "cpl", "10001", "Cove Point pipeline -> Cove Point plant"),
]

# Approximate first-feedgas month per plant (commissioning, a little
# before first cargo) - blanks before these are "not operating yet", not
# missing data. Used to keep pre-start days out of the progress log's
# "missing" list, out of the calibration (plus RAMP_MONTHS of start-up
# after), and to stop an unmetered plant's EIA placeholder being applied
# before it existed. Edit if you have better dates.
PLANT_START = {
    "Sabine Pass": date(2016, 1, 1),
    "Cove Point": date(2018, 2, 1),
    "Corpus Christi": date(2018, 11, 1),
    "Cameron": date(2019, 5, 1),
    "Freeport": date(2019, 8, 1),
    "Elba Island": date(2019, 10, 1),
    "Calcasieu Pass": date(2022, 1, 1),
    "Plaquemines": date(2024, 12, 1),
    "Golden Pass": date(2026, 1, 1),
}
RAMP_MONTHS = 3
# Meters are treated as live this many months BEFORE PLANT_START, so any
# pre-commissioning feedgas is captured (and reported missing if absent);
# the calibration still waits for PLANT_START + RAMP_MONTHS.
PRE_START_MONTHS = 6


def live_from(plant):
    start = PLANT_START.get(plant)
    return (pd.Timestamp(start) - pd.DateOffset(months=PRE_START_MONTHS)).date() if start else date.min


# Everything --dump pulls: all delivery points on these pipelines, to
# find more terminal meters for POINTS.
DUMP_PIPELINES = [
    ("km", "KMLP"), ("km", "NGPL"), ("km", "TGP"), ("km", "EEC"), ("km", "SNG"),
    ("enbridge", "TE"), ("gasnom", "CAMERON"), ("cheniere", "200"), ("cheniere", "400"),
    ("et", "TGC"), ("et", "GR"), ("et", "TGR"), ("et", "LCLNG"), ("et", "FGT"),
    ("tc", "ANR"), ("tc", "CGT"), ("bhe", "cpl"),
]

COVERAGE_NOTES = {
    "Sabine Pass": "Near complete - the plant's three feed pipes' delivery meters: Creole Trail, Kinder Morgan Louisiana and NGPL.",
    "Plaquemines": "Complete - Gator Express's three feeds: Tennessee Gas, Texas Eastern and Columbia Gulf.",
    "Cameron": "Complete - Cameron Interstate's delivery meter into the terminal plus Columbia Gulf's direct feed.",
    "Calcasieu Pass": "Mostly complete - ANR (Grand Chenier XPress) and Texas Eastern deliveries into TransCameron; Sabine Pipe Line's, if any, not yet seen.",
    "Corpus Christi": "Mostly complete - Cheniere Corpus Christi Pipeline's delivery into the plant, which includes intrastate Permian gas it receives; gas delivered straight to the plant by the intrastate ADCC pipeline is not seen.",
    "Freeport": "Partial - Texas Eastern's deliveries at Stratton Ridge and into BIG Pipeline only; Texas intrastate supply is invisible, so use the calibrated estimate. Stratton Ridge also serves Dow's Freeport complex.",
    "Elba Island": "Elba Liquefaction meter on Elba Express.",
    "Golden Pass": "Partial - Gulf Run's delivery into Golden Pass Pipeline; Permian gas via Kinder Morgan's Trident (intrastate) is not seen.",
    "Cove Point": "Complete - Cove Point pipeline's delivery meter into the plant.",
}

KM_URL = "https://pipeline2.kindermorgan.com/Capacity/OpAvailPoint.aspx?code={code}"
KM = "#WebSplitter1_tmpl1_ContentPlaceHolder1_"
KM_DATE_PICKER = "WebSplitter1_tmpl1_ContentPlaceHolder1_dtePickerBegin"
KM_SET_DATE_JS = """([id, y, m, d]) => { const p = $find(id); if (!p) return null;
    p.set_value(new Date(y, m - 1, d)); return String(p.get_value()); }"""
ENBRIDGE_URL = "https://rtba.enbridge.com/InformationalPosting/Default.aspx?bu={bu}&Type=OA"
ENBRIDGE_DATE_INPUT = "#ctl00_MainContent_ctl01_oaDefault_ucDate_rdpDate_dateInput"
GASNOM_URL = "https://www.gasnom.com/ip/{pipe}/oauc.cfm?dt={day:%m/%d/%Y}&type=1"
ET_PAGE_URL = "https://tgcmessenger.energytransfer.com/ipost/capacity/operationally-available-by-location?asset={asset}"
ET_CSV_URL = ET_PAGE_URL + "&f=csv&extension=csv&gasDay={day:%m}%2F{day:%d}%2F{day:%Y}&cycleDesc=Final&pointCd=&name="
CHENIERE_API = "https://lngconnectionapi.cheniere.com/api/Capacity/GetCapacity?tspNo={tsp}&beginDate={begin}&cycleId=null&locationId=0"
TABLE_ROWS_JS = """() => [...document.querySelectorAll('tr')].map(tr =>
    [...tr.querySelectorAll('td,th')].map(c => c.innerText.replace(/\\s+/g, ' ').trim()))"""


def to_number(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "").str.strip(), errors="coerce")


def fetch_km(page, code, gas_day):
    """Delivery points for one KM pipeline and gas day, via the page's
    Excel export. Returns {loc id: scheduled Dth}."""
    page.goto(KM_URL.format(code=code), wait_until="networkidle", timeout=60000)
    page.click(KM + "rbDelivery")
    page.wait_for_load_state("networkidle", timeout=30000)
    if page.evaluate(KM_SET_DATE_JS, [KM_DATE_PICKER, gas_day.year, gas_day.month, gas_day.day]) is None:
        raise RuntimeError("gas day picker not found")
    page.click(KM + "HeaderBTN1_btnRetrieve")
    page.wait_for_function(
        """sel => { const g = document.querySelector(sel);
                    return g && /Row Count:\\s*\\d|No record/.test(g.innerText); }""",
        arg=KM + "DGOpAvail", timeout=180000,
    )
    with page.expect_download(timeout=60000) as info:
        page.click(KM + "HeaderBTN1_btnDownload")
    path = info.value.path()
    meta = pd.read_excel(path, header=None, nrows=2)
    df = pd.read_excel(path, header=3)
    df = df[to_number(df["Loc"]).notna()]
    print(f"  KM {code}: {len(df)} delivery points, {meta.iloc[1, 2]}, cycle {meta.iloc[1, 3]}", flush=True)
    return points_frame(df["Loc"].astype(str).str.split(".").str[0], df["Loc Name"], df["Total Scheduled Quantity"])


ENBRIDGE_PICKER_ID = "ctl00_MainContent_ctl01_oaDefault_ucDate_rdpDate"
ENBRIDGE_SET_DATE_JS = """([id, y, m, d]) => { const p = $find(id); if (!p) return false;
    p.set_selectedDate(new Date(y, m - 1, d)); return true; }"""


def fetch_enbridge(page, bu, gas_day):
    """Delivery points for one Enbridge business unit. The page opens on
    yesterday's gas day; for any other day, set its Telerik date picker
    (which posts back and reloads the cycle list) and pick the latest
    cycle posted for that day. The CSV's own Eff_Gas_Day is checked."""
    page.goto(ENBRIDGE_URL.format(bu=bu), wait_until="networkidle", timeout=60000)
    want = f"{gas_day.month}/{gas_day.day}/{gas_day.year}"
    if page.locator(ENBRIDGE_DATE_INPUT).input_value() != want:
        if not page.evaluate(ENBRIDGE_SET_DATE_JS, [ENBRIDGE_PICKER_ID, gas_day.year, gas_day.month, gas_day.day]):
            raise RuntimeError("Enbridge date picker not found")
        page.wait_for_function(f"() => document.querySelector('{ENBRIDGE_DATE_INPUT}').value === '{want}'", timeout=60000)
        page.wait_for_load_state("networkidle", timeout=90000)
        page.wait_for_timeout(2000)
        latest = page.locator("#ddlSelector option").last.get_attribute("value")
        if page.locator("#ddlSelector").input_value() != latest:
            page.locator("#ddlSelector").select_option(latest)
            page.wait_for_load_state("networkidle", timeout=90000)
            page.wait_for_timeout(2000)
    with page.expect_download(timeout=90000) as info:
        page.locator("a:has-text('Downloadable Format')").first.click()
    with open(info.value.path(), encoding="utf-8", errors="replace") as f:
        df = pd.read_csv(io.StringIO(f.read()))
    days = set(pd.to_datetime(df["Eff_Gas_Day"], format="%m-%d-%Y").dt.date)
    if days != {gas_day}:
        raise RuntimeError(f"Enbridge {bu} CSV is for gas day(s) {days}, wanted {gas_day}")
    df = df[df["Flow_Ind_Desc"] == "Delivery"]
    print(f"  Enbridge {bu}: {len(df)} delivery points, gas day {gas_day}, cycle {df['Cycle_Desc'].iloc[0]}", flush=True)
    return points_frame(df["Loc"].astype(str), df["Loc_Name"], df["Total_Scheduled_Quantity"])


def fetch_gasnom(page, pipe, gas_day):
    """gasnom.com OAC page: a plain HTML table per gas day. Rows with a
    blank name continue the location above (its other flow direction)."""
    page.goto(GASNOM_URL.format(pipe=pipe, day=gas_day), wait_until="networkidle", timeout=60000)
    text = page.inner_text("body")
    wanted = f"{gas_day:%B} {gas_day.day}, {gas_day:%Y}"
    if wanted not in text.split("Eff Gas Day")[-1][:80]:
        raise RuntimeError(f"page isn't for gas day {wanted}: {text[:300]!r}")
    rows = page.evaluate(TABLE_ROWS_JS)
    header = next(r for r in rows if "Loc" in r and "TSQ" in r)
    i_name, i_loc, i_flow, i_tsq = (header.index(c) for c in ("Location Name", "Loc", "Flow Ind", "TSQ"))
    data, last_name = [], ""
    for r in rows:
        if len(r) < len(header) or r == header or not r[i_loc].isdigit():
            continue
        last_name = r[i_name] or last_name
        if r[i_flow] == "D":
            data.append((r[i_loc], last_name, r[i_tsq]))
    df = pd.DataFrame(data, columns=["loc", "name", "tsq"])
    cycle = text.split("Cycle Indicator Description:")[-1].split("Capacity Type")[0].strip()[:30]
    print(f"  gasnom {pipe}: {len(df)} delivery points, gas day {wanted}, cycle {cycle!r}", flush=True)
    return points_frame(df["loc"], df["name"], df["tsq"])


def fetch_et(page, asset, gas_day):
    """Energy Transfer Messenger+ OA-by-location CSV export for the gas
    day's Final cycle. The HTML page is loaded first so the export
    request carries the site's cookies."""
    page.goto(ET_PAGE_URL.format(asset=asset), wait_until="networkidle", timeout=60000)
    response = page.context.request.get(ET_CSV_URL.format(asset=asset, day=gas_day), timeout=90000)
    if not response.ok:
        raise RuntimeError(f"CSV export HTTP {response.status}")
    df = pd.read_csv(io.StringIO(response.text()))
    cols = {c.lower().strip(): c for c in df.columns}
    def col(*names):
        for n in names:
            if n in cols:
                return cols[n]
        raise KeyError(f"none of {names} in columns {list(df.columns)}")
    flow = col("flow ind", "flow indicator", "flow_ind")
    df = df[df[flow].astype(str).str.upper().str.startswith("D")]
    print(f"  ET {asset}: {len(df)} delivery points, gas day {gas_day}, Final cycle", flush=True)
    return points_frame(df[col("loc", "location")].astype(str).str.split(".").str[0],
                        df[col("loc name", "location name")], df[col("tsq", "total scheduled quantity")])


def fetch_cheniere(page, tsp, gas_day):
    """Cheniere LNG Connection's JSON capacity feed (what its site calls
    behind the scenes; operator-hosted, no browser needed). Asks for the
    gas day explicitly, falls back to the default (current gas day's
    latest cycle), and checks the effective date either way."""
    import requests
    for begin in (f"{gas_day:%m/%d/%Y}", "null"):
        r = requests.get(CHENIERE_API.format(tsp=tsp, begin=begin), timeout=60,
                         headers={"User-Agent": "Mozilla/5.0", "Accept": "application/json"})
        r.raise_for_status()
        report = r.json().get("report") or []
        if not report:
            continue
        eff = datetime.fromisoformat(report[0]["avaiL_CAP_EFF_DT_TIME"])
        eff_day = (eff - timedelta(hours=9)).date()  # gas day runs 9am-9am Central
        if eff_day == gas_day:
            break
    else:
        raise RuntimeError(f"no Cheniere {tsp} posting for gas day {gas_day}")
    if eff_day != gas_day:
        raise RuntimeError(f"Cheniere {tsp} posting is for gas day {eff_day}, wanted {gas_day}")
    df = pd.DataFrame(report)
    keys = {k.lower(): k for k in df.columns}
    tsq = next(keys[k] for k in keys if "sch" in k and "q" in k)
    flow = next(keys[k] for k in keys if "purp" in k)
    df = df[df[flow].astype(str).str.contains("Deliver", case=False)]
    print(f"  Cheniere {tsp}: {len(df)} delivery points, gas day {eff_day}, cycle {report[0].get('cycle')}", flush=True)
    return points_frame(df[keys["loc"]], df[keys["loc_name"]], df[tsq])


TC_BASE = "https://ebb.tceconnects.com/infopost/"
TC_REPORTS = {"ANR": ("OperationallyAvailableCapacityANR", 3005), "CGT": ("OperationallyAvailableCapacity", 14)}
BHE_OA_LIST = "https://infopost.bhegts.com/{pipe}/postings/capacity-operationally-available"


# TC's report viewer can't change gas day for anyone (its session
# keep-alive redirects to TC's error page, so the viewer shows "ASP.NET
# session has expired" and the date box does nothing). The report URL with
# rs:Format=CSV does work, but only ever serves TC's latest posting - by
# the morning run, the NEXT gas day's Timely cycle. So each run caches that
# posting ("TC next day (timely)" sheet) and the following run, pulling
# that gas day, takes TC's meters from the cache: TC's values are the
# Timely nominations, not the final intraday cycle, and there's no TC
# history before the cache started.
TC_AHEAD = {}  # gas day -> {point column: Dth}, filled by fetch_tc
TC_CACHE_SHEET = "TC next day (timely)"
_TC_LATEST = {}  # pipe -> (gas day, frame): fetched once per run


def fetch_tc(page, pipe, gas_day):
    if pipe not in _TC_LATEST:
        _TC_LATEST[pipe] = fetch_tc_latest(page, pipe)
    posted, frame = _TC_LATEST[pipe]
    if posted == gas_day:
        return frame
    for point in POINTS:
        if point[1] == "tc" and point[2] == pipe and point[3] in frame.index:
            TC_AHEAD.setdefault(posted, {})[point_column(point)] = frame.at[point[3], "scheduled_dth"]
    raise RuntimeError(f"TC only serves its latest posting ({posted}) - cached for that day's run; "
                       f"{gas_day} comes from the cache if an earlier run saw it")


def fetch_tc_latest(page, pipe):
    report, asset = TC_REPORTS[pipe]
    page.goto(TC_BASE + f"TCeConnects.aspx?v=1.3&SID=67&info=Y&assetid={asset}", wait_until="domcontentloaded", timeout=90000)
    page.wait_for_timeout(2000)
    url = f"{TC_BASE}ReportViewer.aspx?/InfoPost/{report}&pAssetNbr={asset}&rs:Format=CSV"
    text = page.context.request.get(url, timeout=180000).text().lstrip("\ufeff")
    if not text.startswith("TSPName"):
        snippet = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", text))[:200]
        raise RuntimeError(f"CSV export didn't return the OA report: {snippet!r}")
    df = pd.read_csv(io.StringIO(text), dtype=str)
    days = set(pd.to_datetime(df["EffGasDay"]).dt.date)
    if len(days) != 1:
        raise RuntimeError(f"TC {pipe} export spans gas days {sorted(days)}")
    posted = days.pop()
    tsq = next(c for c in df.columns if c.startswith("TotalSched"))
    df = df[df["LocPurpDesc"].astype(str).str.startswith("Delivery")]
    frame = points_frame(df["Location"], df["LocationName"], df[tsq])
    print(f"  TC {pipe}: {len(frame)} delivery points, latest posting gas day {posted}, cycle {df['Cycle'].iloc[0]}", flush=True)
    return posted, frame


def fetch_bhe(page, pipe, gas_day):
    """BHE GT&S OA postings: the listing page links a CSV per cycle; take
    the latest-posted one for the gas day (listing is newest first)."""
    page.goto(BHE_OA_LIST.format(pipe=pipe), wait_until="networkidle", timeout=90000)
    page.wait_for_timeout(3000)
    rows = page.evaluate("""() => [...document.querySelectorAll('[role=row]')].map(r => ({
        text: r.innerText.replace(/\\s+/g, ' '), csv: (r.querySelector('a[href$=".csv"]') || {}).href}))""")
    wanted = f"Capacity Available {gas_day:%m/%d/%Y}"
    row = next((r for r in rows if r.get("csv") and wanted in r["text"]), None)
    if row is None:
        raise RuntimeError(f"no BHE {pipe} posting for {gas_day}")
    df = pd.read_csv(io.StringIO(page.context.request.get(row["csv"], timeout=60000).text()), dtype=str)
    df = df[df["Loc Purp Desc"].astype(str).str.startswith("Delivery")]
    print(f"  BHE {pipe}: {len(df)} delivery points, gas day {gas_day}, cycle {df['CycleDesc'].iloc[0]}", flush=True)
    return points_frame(df["Loc"], df["Loc Name"], df["Total Scheduled Quantity"])


def points_frame(locs, names, tsq):
    """Common shape for every fetcher: loc id -> name, scheduled Dth."""
    df = pd.DataFrame({"loc": locs.astype(str).str.strip().values,
                       "name": names.astype(str).map(lambda s: " ".join(s.split())).values,
                       "scheduled_dth": to_number(pd.Series(tsq.values))})
    return df.drop_duplicates("loc", keep="last").set_index("loc")


def point_column(point):
    plant, platform, pipeline, loc, label = point
    return f"{plant} | {label} ({pipeline} {loc})"


FETCHERS = {"km": fetch_km, "enbridge": fetch_enbridge, "gasnom": fetch_gasnom, "et": fetch_et, "cheniere": fetch_cheniere,
            "tc": fetch_tc, "bhe": fetch_bhe}


def fetch_all(pipelines, gas_day):
    """{(platform, pipeline): points frame}; a pipeline that fails is
    logged and left out."""
    results = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(ignore_https_errors=True, accept_downloads=True, viewport={"width": 1600, "height": 1200})
        page = context.new_page()
        for platform, pipeline in pipelines:
            try:
                results[(platform, pipeline)] = FETCHERS[platform](page, pipeline, gas_day)
            except Exception as e:
                print(f"  WARNING: {platform} {pipeline} failed ({type(e).__name__}: {str(e)[:300]})", file=sys.stderr, flush=True)
        browser.close()
    return results


def dump(gas_day):
    """Every delivery point with scheduled gas on DUMP_PIPELINES, to a
    CSV next to this script - for finding terminal meters to add."""
    results = fetch_all(DUMP_PIPELINES, gas_day)
    frames = [df.assign(platform=pl, pipeline=pipe) for (pl, pipe), df in results.items()]
    allpts = pd.concat(frames).reset_index().sort_values("scheduled_dth", ascending=False)
    path = os.path.join(HERE, f"feedgas_point_dump_{gas_day}.csv")
    allpts.to_csv(path, index=False)
    print(f"\nSaved {path} ({len(allpts)} delivery points)")
    for (pl, pipe), df in results.items():
        print(f"\n== {pl} {pipe}: top delivery points by scheduled Dth")
        print(df.sort_values("scheduled_dth", ascending=False).head(15).to_string())


def pull(gas_day, tc_cache=None):
    """One row of point-level scheduled Dth for the gas day. Points on a
    pipeline that failed to load are left blank, not zero."""
    needed = sorted({(p[1], p[2]) for p in POINTS})
    results = fetch_all(needed, gas_day)
    row = {}
    for point in POINTS:
        plant, platform, pipeline, loc, label = point
        values = results.get((platform, pipeline))
        if values is None:
            row[point_column(point)] = float("nan")
        elif loc not in values.index:
            print(f"  WARNING: {pipeline} point {loc} ({label}) not in today's posting", file=sys.stderr, flush=True)
            row[point_column(point)] = float("nan")
        else:
            row[point_column(point)] = values.at[loc, "scheduled_dth"]
    # TC meters: from the Timely posting an earlier run cached for this day
    cached = {**(tc_cache.loc[gas_day].dropna().to_dict() if tc_cache is not None and gas_day in tc_cache.index else {}),
              **TC_AHEAD.get(gas_day, {})}
    for col, value in cached.items():
        if col in row and pd.isna(row[col]):
            row[col] = value
            print(f"  {col}: {value:,.0f} Dth from the cached TC Timely posting", flush=True)
    return pd.DataFrame([row], index=pd.Index([gas_day], name="gas_day"))


def load_tc_cache(path):
    try:
        df = pd.read_excel(path, sheet_name=TC_CACHE_SHEET, index_col=0)
    except (FileNotFoundError, ValueError):
        return None
    df.index = pd.to_datetime(df.index).date
    df.index.name = "gas_day"
    return df


def plants_bcfd(points_dth):
    """Plant totals in Bcf/d from the point table. A plant with some
    points missing sums what it has; all missing -> blank."""
    out = {}
    for plant in dict.fromkeys(p[0] for p in POINTS):
        cols = [point_column(p) for p in POINTS if p[0] == plant and point_column(p) in points_dth.columns]
        out[plant] = points_dth[cols].sum(axis=1, min_count=1) / DTH_PER_BCF
    df = pd.DataFrame(out, index=points_dth.index)
    df["Total"] = df.sum(axis=1, min_count=1)
    return df.round(3)


# ---- Calibration against EIA monthly LNG exports by terminal ----------
# EIA's "U.S. Liquefied Natural Gas Exports by Point of Exit" (monthly,
# ~2-3 months behind, free, no key) gives what each terminal actually
# exported. Feedgas is larger than exports by the plant's own fuel and
# shrinkage - FEEDGAS_PER_EXPORT is a round, stated assumption (~9%),
# not a measured figure. Comparing monthly-average metered feedgas with
# EIA-implied feedgas gives each plant's "seen share" of its supply;
# dividing daily metered by the latest seen share gives a calibrated
# estimate - the fix for plants whose supply is partly invisible
# (Freeport's Texas intrastate gas above all).
EIA_EXPORTS_XLS = "https://www.eia.gov/dnav/ng/xls/NG_MOVE_POE2_A_EPG0_ENG_MMCF_M.xls"
FEEDGAS_PER_EXPORT = 1.09
MIN_DAYS_FOR_MONTH = 20
PORT_TO_PLANT = [  # first match wins; Calcasieu before Cameron (both in Cameron Parish)
    (r"calcasieu", "Calcasieu Pass"), (r"sabine", "Sabine Pass"), (r"cameron", "Cameron"),
    (r"freeport", "Freeport"), (r"corpus", "Corpus Christi"), (r"cove point", "Cove Point"),
    (r"elba", "Elba Island"), (r"plaquemines", "Plaquemines"), (r"golden pass", "Golden Pass"),
]


def fetch_eia_exports():
    """Monthly EIA exports by plant, as feedgas-equivalent Bcf/d.

    The workbook splits each point of exit by destination country
    ("Sabine Pass, LA Liquefied Natural Gas Exports to Japan") across
    several "Data N" sheets. Use a port's own total column where one
    exists ("Sabine Pass, LA Liquefied Natural Gas Exports (MMcf)"),
    otherwise sum all of its country columns across every sheet - the
    first version read only "Data 1", which stops part-way through the
    alphabet of countries, and undercounted every terminal."""
    import re
    import requests
    r = requests.get(EIA_EXPORTS_XLS, timeout=120, headers={"User-Agent": "Mozilla/5.0"})
    r.raise_for_status()
    book = pd.ExcelFile(io.BytesIO(r.content))
    frames = []
    for name in [n for n in book.sheet_names if n.lower().startswith("data")]:
        sheet = pd.read_excel(book, sheet_name=name, header=2)
        sheet = sheet.rename(columns={sheet.columns[0]: "month"}).dropna(subset=["month"])
        sheet["month"] = pd.to_datetime(sheet["month"]).dt.to_period("M")
        frames.append(sheet.set_index("month"))
    raw = pd.concat(frames, axis=1)
    raw = raw.loc[:, ~raw.columns.duplicated()]
    totals, by_country = {}, {}
    for col in raw.columns:
        c = str(col)
        plant = next((pl for pat, pl in PORT_TO_PLANT if re.search(pat, c.split(" Liquefied")[0], re.I)), None)
        if plant is None:
            continue
        series = pd.to_numeric(raw[col], errors="coerce")
        # "Exports to All Countries" is the terminal's total, not a country
        # (adding it to the country columns doubled Plaquemines, Calcasieu
        # Pass, Elba and Golden Pass).
        if re.search(r"Exports to ", c) and not re.search(r"Exports to All Countries", c, re.I):
            # The same terminal-country pair appears under more than one
            # label (MMcf vs Million Cubic Feet, repeated across sheets) -
            # summing every column doubled Plaquemines, Calcasieu Pass and
            # Elba. Keep one series per (terminal, country), the larger.
            country = re.sub(r"\s*\(.*$", "", c.split("Exports to ", 1)[1]).strip().lower()
            key = (plant, country)
            by_country[key] = series if key not in by_country else pd.concat([by_country[key], series], axis=1).max(axis=1)
        else:
            # Some terminals have the same total repeated under two labels
            # (units spelled differently, or on two sheets) - summing them
            # doubled Plaquemines, Calcasieu Pass and Elba. Take the larger
            # value month by month instead of the sum.
            totals.setdefault(plant, []).append((c, series))
    for plant, cols in totals.items():
        print(f"  EIA exports: {plant} total from {[c for c, _ in cols]}", flush=True)
    totals = {pl: pd.concat([s for _, s in cols], axis=1).max(axis=1) for pl, cols in totals.items()}
    country_sums = {}
    for (plant, _), series in by_country.items():
        country_sums[plant] = country_sums[plant].add(series, fill_value=0) if plant in country_sums else series
    by_country = country_sums
    # a total where EIA gives one, the country sum for months it doesn't
    mmcf = {pl: totals[pl].combine_first(by_country[pl]) if pl in totals and pl in by_country
            else totals.get(pl, by_country.get(pl)) for pl in set(totals) | set(by_country)}
    days = raw.index.days_in_month
    df = pd.DataFrame({pl: v / 1000 / days * FEEDGAS_PER_EXPORT for pl, v in mmcf.items()}, index=raw.index)
    df.index = df.index.astype(str)
    df.index.name = "month"
    df = df.dropna(how="all").round(3)
    print(f"  EIA exports: {len(df)} months, latest {df.index[-1]}; port totals for {sorted(totals)}, "
          f"country sums for {sorted(set(by_country) - set(totals))}", flush=True)
    return df


def monthly_check(plants_daily, eia):
    """Metered monthly average vs EIA-implied feedgas, per plant."""
    daily = plants_daily.drop(columns=["Total"], errors="ignore").copy()
    daily.index = pd.to_datetime(daily.index)
    months = daily.index.to_period("M").astype(str)
    counts = daily.groupby(months).count()
    metered = daily.groupby(months).mean().where(counts >= MIN_DAYS_FOR_MONTH)
    rows = []
    for month in sorted(set(metered.index) | set(eia.index)):
        for plant in sorted(set(daily.columns) | set(eia.columns)):
            m = metered.at[month, plant] if month in metered.index and plant in metered.columns else float("nan")
            e = eia.at[month, plant] if month in eia.index and plant in eia.columns else float("nan")
            if pd.isna(m) and pd.isna(e):
                continue
            rows.append({"month": month, "plant": plant, "eia_implied_feedgas_bcfd": e,
                         "metered_bcfd": round(m, 3) if pd.notna(m) else m,
                         "seen_share": round(m / e, 3) if pd.notna(m) and pd.notna(e) and e > 0 else float("nan")})
    return pd.DataFrame(rows).set_index(["month", "plant"]) if rows else pd.DataFrame()


def estimated(plants_daily, check, eia=None):
    """Daily estimate = metered / latest seen share, per plant with a
    share; plants without one yet carry their metered value. Plants in
    EIA's data with no meter at all (e.g. Cove Point until it's covered)
    carry their latest EIA month - stale, but better than leaving them
    out of the estimated total."""
    daily = plants_daily.drop(columns=["Total"], errors="ignore").copy()
    metered_total = daily.sum(axis=1, min_count=1)
    if eia is not None and not eia.empty:
        for plant in eia.columns:
            if plant not in daily.columns or daily[plant].isna().all():
                latest = eia[plant].dropna()
                if not latest.empty:
                    start = live_from(plant)
                    after_start = pd.Series([d >= start for d in daily.index], index=daily.index)
                    daily[f"{plant} (EIA {latest.index[-1]}, unmetered)"] = latest.iloc[-1] * after_start.where(after_start)
                    daily = daily.drop(columns=[plant], errors="ignore")
    share = {}
    if not check.empty:
        s = check["seen_share"].dropna()
        for (month, plant), v in s.sort_index().items():
            start = PLANT_START.get(plant)
            if start is not None:
                settled = (pd.Period(start, "M") + RAMP_MONTHS).strftime("%Y-%m")
                if month < settled:  # before start-up, or still ramping up
                    continue
            if 0.05 < v < 2:  # ignore nonsense ratios (outages, turnarounds)
                share[plant] = v
    est = pd.DataFrame({p: (daily[p] / share[p]) if p in share else daily[p] for p in daily.columns}, index=daily.index)
    est["Estimated total"] = est.sum(axis=1, min_count=1)
    est["Metered total"] = metered_total
    return est.round(3), share


def load_points(path):
    try:
        df = pd.read_excel(path, sheet_name="Points (Dth)", index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    df.index = pd.to_datetime(df.index).date
    df.index.name = "gas_day"
    return df


def notes_lines():
    lines = [
        "UNITS",
        f"'Bcfd by plant': billion cubic feet per day = scheduled Dth / {DTH_PER_BCF:,.0f} ({BTU_PER_CF:,} Btu/cf, EIA average heat content).",
        "'Points (Dth)': scheduled quantity at each terminal-side meter, in dekatherms (= MMBtu), as posted by the pipeline.",
        "",
        "WHAT THIS IS",
        "Pipeline-posted scheduled deliveries into each LNG terminal, for the gas day (9am-9am Central), best-available "
        "cycle - by the next morning that is the final intraday nomination cycle. Nominations, not metered flow.",
        "Only terminal-side meters are counted, so gas passed between pipelines upstream is not double counted.",
        "Total = sum of the plants shown. It is a floor on US feedgas while coverage is partial.",
        "",
        "COVERAGE BY PLANT",
    ]
    lines += [f"{plant}: {note}" for plant, note in COVERAGE_NOTES.items()]
    lines += ["", "PLANT START DATES (approximate first feedgas; blanks before these = not operating yet)"]
    lines += [f"{plant}: {start:%b %Y} (meters treated as live from {live_from(plant):%b %Y}, "
              f"{PRE_START_MONTHS} months earlier, to catch pre-commissioning feedgas)" for plant, start in PLANT_START.items()]
    lines += ["", "CALIBRATION",
              f"'EIA exports (Bcfd)': EIA monthly LNG exports by terminal (point of exit), converted to Bcf/d and grossed up by "
              f"{FEEDGAS_PER_EXPORT:.2f} for liquefaction fuel and shrinkage (an assumed ~{(FEEDGAS_PER_EXPORT - 1) * 100:.0f}%) - "
              "i.e. the feedgas each plant actually needed. About 2-3 months behind.",
              f"'Monthly check': for each month with {MIN_DAYS_FOR_MONTH}+ metered days, metered average vs EIA-implied feedgas; "
              "seen_share = the part of each plant's supply our meters see.",
              "'Estimated (calibrated)': daily metered / the plant's latest seen_share - fills the gas we can't see "
              "(Freeport's Texas intrastate supply above all). Plants with no share yet (the first 2-3 months, until "
              "EIA catches up with our history) carry their metered value. 'Estimated total' vs 'Metered total' "
              "shows how much is inferred.",
              "", "TC ENERGY METERS (ANR, COLUMBIA GULF)",
              "TC eConnects only serves its latest posting - by each morning run, the next gas day's Timely cycle. "
              f"Each run caches it ('{TC_CACHE_SHEET}') and the next run uses it for that day, so TC's meters "
              "(Calcasieu Pass via ANR; Cameron and Plaquemines via Columbia Gulf) are Timely nominations, not the final "
              "cycle, and have no history before the cache started. A missed run leaves that day's TC meters blank.",
              "", "POINTS", "Each 'Points (Dth)' column is 'Plant | meter (pipeline code, location id)'.", "",
              "SOURCE", "Kinder Morgan (pipeline2.kindermorgan.com) and Enbridge LINK (rtba.enbridge.com) "
              "Operationally Available Capacity postings, Total Scheduled Quantity column; gasnom.com (Cameron Interstate), "
              "Energy Transfer Messenger+, Cheniere LNG Connection, TC eConnects and BHE GT&S equivalents - all operator-hosted."]
    return lines


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--out", default=DEFAULT_OUT)
    parser.add_argument("--gas-day", help="YYYY-MM-DD (default: yesterday, US Central)")
    parser.add_argument("--dump", action="store_true", help="list every delivery point on DUMP_PIPELINES instead of updating the workbook")
    parser.add_argument("--probe", help="comma-separated gas days (YYYY-MM-DD,...): report which pipelines still "
                                        "serve each day, without touching the workbook - to see how far back a backfill can go")
    parser.add_argument("--from", dest="date_from", help="backfill: first gas day, YYYY-MM-DD (with --to)")
    parser.add_argument("--to", dest="date_to", help="backfill: last gas day, YYYY-MM-DD (default: yesterday)")
    args = parser.parse_args()
    yesterday = datetime.now(ZoneInfo("America/Chicago")).date() - timedelta(days=1)
    gas_day = date.fromisoformat(args.gas_day) if args.gas_day else yesterday
    if args.dump:
        dump(gas_day)
        return
    if args.probe:
        probe([date.fromisoformat(d.strip()) for d in args.probe.split(",")])
        return

    if args.date_from:
        first, last = date.fromisoformat(args.date_from), date.fromisoformat(args.date_to) if args.date_to else yesterday
        days = [first + timedelta(days=i) for i in range((last - first).days + 1)]
    else:
        days = [gas_day]

    # Backfills save every SAVE_EVERY days, so an interrupted run keeps
    # what it has (re-running skips nothing - it just overwrites).
    SAVE_EVERY = 7
    pending, retrieved = [], 0
    tc_cache = load_tc_cache(args.out)
    started = datetime.now()
    progress_log = os.path.splitext(args.out)[0] + "_progress.log"
    for i, day in enumerate(days, 1):
        print(f"Pulling LNG feedgas points for gas day {day} ({i}/{len(days)})...", flush=True)
        row = pull(day, tc_cache)
        if row.notna().sum(axis=1).iloc[0] == 0:
            print(f"  No points retrieved for {day} - skipped.", file=sys.stderr)
        else:
            pending.append(row)
            retrieved += 1
        progress_line(progress_log, i, len(days), day, row, started)
        if pending and (i % SAVE_EVERY == 0 or i == len(days)):
            save(args.out, pd.concat(pending))
            pending = []
    if retrieved == 0:
        print("No points retrieved at all - leaving the archive untouched.", file=sys.stderr)
        sys.exit(1)


def progress_line(path, i, n, day, row, started):
    """One summary line per gas day - printed, and appended to a
    <workbook>_progress.log next to the workbook - so a long backfill can
    be followed day by day: meters found, metered total, which pipes
    were missing, time elapsed and an estimate of time left."""
    values = row.iloc[0]
    live = [p for p in POINTS if day >= live_from(p[0])]
    got = int(sum(pd.notna(values.get(point_column(p))) for p in live))
    missing = sorted({f"{p[0]} ({p[2]})" for p in live if pd.isna(values.get(point_column(p)))})
    total = values.sum(skipna=True) / DTH_PER_BCF
    elapsed = datetime.now() - started
    left = elapsed / i * (n - i)
    fmt = lambda td: f"{int(td.total_seconds() // 3600)}:{int(td.total_seconds() % 3600 // 60):02d}"
    mark = "✓" if got == len(live) else ("~" if got else "✗")
    line = (f"[{i:>4}/{n}] {day}  {mark} {got}/{len(live)} meters  metered {total:5.2f} Bcf/d"
            + (f"  missing: {', '.join(missing)}" if missing else "")
            + f"  |  elapsed {fmt(elapsed)}  ~{fmt(left)} left")
    print(line, flush=True)
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now():%Y-%m-%d %H:%M}  {line}\n")


def probe(days):
    """For each gas day, which pipelines return a posting - pipelines are
    only required to keep ~3 years of postings, so this shows how far
    back a backfill will actually find data. Nothing is saved."""
    needed = sorted({(p[1], p[2]) for p in POINTS})
    table = {}
    for day in days:
        print(f"\nProbing {day}...", flush=True)
        got = fetch_all(needed, day)
        table[str(day)] = {f"{pl} {pipe}": ("ok" if (pl, pipe) in got else "-") for pl, pipe in needed}
    print("\n=== PROBE: which pipelines still serve each gas day (ok / -) ===")
    print(pd.DataFrame(table).to_string())


_EIA_CACHE = {}


def save(out, new):
    """Upsert new gas-day rows into the workbook and rewrite every sheet."""
    points = load_points(out)
    points = new if points.empty else new.combine_first(points)
    # Columns follow POINTS exactly: a meter dropped from POINTS (e.g. the
    # three upstream Cameron interconnects replaced by Cameron Interstate's
    # own terminal meter) drops out of the sheet rather than lingering.
    points = points.reindex(columns=[point_column(p) for p in POINTS])
    points = points.sort_index()
    points.index.name = "gas_day"
    plants = plants_bcfd(points)
    args = argparse.Namespace(out=out)

    try:
        if "eia" not in _EIA_CACHE:  # once per run, not at every backfill save
            _EIA_CACHE["eia"] = fetch_eia_exports()
        eia = _EIA_CACHE["eia"]
    except Exception as e:
        print(f"  WARNING: EIA exports fetch failed ({type(e).__name__}: {e}) - reusing the workbook's copy", file=sys.stderr)
        try:
            eia = pd.read_excel(args.out, sheet_name="EIA exports (Bcfd)", index_col=0)
            eia.index = eia.index.astype(str)
        except (FileNotFoundError, ValueError):
            eia = pd.DataFrame()
    check = monthly_check(plants, eia) if not eia.empty else pd.DataFrame()
    est, share = estimated(plants, check, eia)
    if share:
        print(f"  seen shares used: {share}")

    sheets = {"Bcfd by plant": plants, "Estimated (calibrated)": est, "Points (Dth)": points}
    cache = load_tc_cache(out)
    if TC_AHEAD:
        ahead = pd.DataFrame.from_dict(TC_AHEAD, orient="index")
        cache = ahead if cache is None else ahead.combine_first(cache)
    if cache is not None and not cache.empty:
        cache = cache.sort_index().tail(14)
        cache.index.name = "gas_day"
        sheets[TC_CACHE_SHEET] = cache
    if not check.empty:
        sheets["Monthly check"] = check
    if not eia.empty:
        sheets["EIA exports (Bcfd)"] = eia
    xlsx_notes.write_workbook(args.out, sheets, notes_lines(),
                              {"UNITS", "WHAT THIS IS", "COVERAGE BY PLANT", "CALIBRATION", "POINTS", "SOURCE",
                               "TC ENERGY METERS (ANR, COLUMBIA GULF)",
                               "PLANT START DATES (approximate first feedgas; blanks before these = not operating yet)"})
    import openpyxl
    wb = openpyxl.load_workbook(args.out)
    for name in ("Bcfd by plant", "Estimated (calibrated)", "Points (Dth)", TC_CACHE_SHEET):
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        for (cell,) in ws.iter_rows(min_row=2, max_col=1):
            cell.number_format = "dd-mmm-yyyy"
        ws.column_dimensions["A"].width = 14
    wb.save(args.out)

    print(f"\nSaved {args.out} - {len(points)} gas day(s).")
    print(plants.tail().to_string())


if __name__ == "__main__":
    main()
