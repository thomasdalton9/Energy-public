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
# "gasnom" = gasnom.com pipeline, "et" = Energy Transfer asset code.
POINTS = [
    ("Sabine Pass", "km", "KMLP", "49448", "KMLP -> SP Liquefaction, Cameron Par."),
    ("Sabine Pass", "km", "NGPL", "46622", "NGPL -> Sabine Pass Liquefaction"),
    ("Sabine Pass", "enbridge", "TE", "75866", "Texas Eastern -> Cheniere (Creole Trail), Beauregard Par."),
    ("Plaquemines", "km", "TGP", "55833", "Tennessee Gas -> VG Gator Express, Evangeline Pass"),
    ("Plaquemines", "enbridge", "TE", "74530", "Texas Eastern -> Gator Express"),
    # Cameron Interstate's own delivery meter into the terminal - its
    # receipts include the Tennessee Gas (TENN-CIP) and Texas Eastern
    # (TETCO-CIP) quantities seen on those pipes, plus LEAP, LEG, NG3 and
    # Gillis Hub, so this one meter is the whole plant.
    ("Cameron", "gasnom", "CAMERON", "772300", "Cameron Interstate -> Cameron LNG"),
    ("Calcasieu Pass", "enbridge", "TE", "74529", "Texas Eastern -> TransCameron, Oak Grove"),
    ("Corpus Christi", "km", "NGPL", "48934", "NGPL -> Cheniere Corpus Christi Pipeline, Sinton"),
    ("Corpus Christi", "km", "TGP", "49861", "Tennessee Gas -> Cheniere Corpus Christi Pipeline, Sinton"),
    ("Freeport", "enbridge", "TE", "79999", "Texas Eastern -> Stratton Ridge"),
    ("Elba Island", "km", "EEC", "660700", "Elba Express -> Elba Liquefaction, Chatham"),
]

# Everything --dump pulls: all delivery points on these pipelines, to
# find more terminal meters for POINTS.
DUMP_PIPELINES = [
    ("km", "KMLP"), ("km", "NGPL"), ("km", "TGP"), ("km", "EEC"), ("km", "SNG"),
    ("enbridge", "TE"), ("gasnom", "CAMERON"),
    ("et", "TGC"), ("et", "GR"), ("et", "TGR"), ("et", "LCLNG"), ("et", "FGT"),
]

COVERAGE_NOTES = {
    "Sabine Pass": "Partial - missing Trunkline/Creole Trail (Energy Transfer) and Transco (Williams).",
    "Plaquemines": "Near complete - Gator Express is fed by Tennessee Gas and Texas Eastern.",
    "Cameron": "Complete - Cameron Interstate's delivery meter into the terminal.",
    "Calcasieu Pass": "Partial - TransCameron is intrastate; only Texas Eastern's delivery into it is seen.",
    "Corpus Christi": "Partial - Permian supply arrives on intrastate pipes (GCX, Whistler, ADCC) with no public data; only NGPL/TGP deliveries into Cheniere's Corpus Christi Pipeline are seen.",
    "Freeport": "Partial - Stratton Ridge only; BIG Pipeline and Gulf South meters not yet covered. Stratton Ridge is a hub, so check against Freeport's reported output.",
    "Elba Island": "Elba Liquefaction meter on Elba Express.",
    "Golden Pass": "Not yet covered - fed by Gulf Run (Energy Transfer) and Trident (intrastate).",
    "Cove Point": "Not yet covered - Transco, Columbia and DETI.",
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


def fetch_enbridge(page, bu, gas_day):
    """Delivery points for one Enbridge business unit. The page defaults
    to yesterday's gas day and the latest cycle - only that default is
    used, so it must match the requested gas day."""
    page.goto(ENBRIDGE_URL.format(bu=bu), wait_until="networkidle", timeout=60000)
    shown = page.locator(ENBRIDGE_DATE_INPUT).input_value()
    if datetime.strptime(shown, "%m/%d/%Y").date() != gas_day:
        raise RuntimeError(f"page shows gas day {shown}, wanted {gas_day} (only the default day is supported)")
    with page.expect_download(timeout=90000) as info:
        page.locator("a:has-text('Downloadable Format')").first.click()
    with open(info.value.path(), encoding="utf-8", errors="replace") as f:
        df = pd.read_csv(io.StringIO(f.read()))
    df = df[df["Flow_Ind_Desc"] == "Delivery"]
    print(f"  Enbridge {bu}: {len(df)} delivery points, gas day {shown}, cycle {df['Cycle_Desc'].iloc[0]}", flush=True)
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


def points_frame(locs, names, tsq):
    """Common shape for every fetcher: loc id -> name, scheduled Dth."""
    df = pd.DataFrame({"loc": locs.astype(str).str.strip().values,
                       "name": names.astype(str).map(lambda s: " ".join(s.split())).values,
                       "scheduled_dth": to_number(pd.Series(tsq.values))})
    return df.drop_duplicates("loc", keep="last").set_index("loc")


def point_column(point):
    plant, platform, pipeline, loc, label = point
    return f"{plant} | {label} ({pipeline} {loc})"


FETCHERS = {"km": fetch_km, "enbridge": fetch_enbridge, "gasnom": fetch_gasnom, "et": fetch_et}


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


def pull(gas_day):
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
    return pd.DataFrame([row], index=pd.Index([gas_day], name="gas_day"))


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
    lines += ["", "POINTS", "Each 'Points (Dth)' column is 'Plant | meter (pipeline code, location id)'.", "",
              "SOURCE", "Kinder Morgan (pipeline2.kindermorgan.com) and Enbridge LINK (rtba.enbridge.com) "
              "Operationally Available Capacity postings, Total Scheduled Quantity column; gasnom.com (Cameron Interstate) "
              "and Energy Transfer Messenger+ equivalents."]
    return lines


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--out", default=DEFAULT_OUT)
    parser.add_argument("--gas-day", help="YYYY-MM-DD (default: yesterday, US Central)")
    parser.add_argument("--dump", action="store_true", help="list every delivery point on DUMP_PIPELINES instead of updating the workbook")
    args = parser.parse_args()
    gas_day = date.fromisoformat(args.gas_day) if args.gas_day else datetime.now(ZoneInfo("America/Chicago")).date() - timedelta(days=1)
    if args.dump:
        dump(gas_day)
        return

    print(f"Pulling LNG feedgas points for gas day {gas_day}...", flush=True)
    new = pull(gas_day)
    if new.notna().sum(axis=1).iloc[0] == 0:
        print("No points retrieved at all - leaving the archive untouched.", file=sys.stderr)
        sys.exit(1)

    points = load_points(args.out)
    points = new if points.empty else new.combine_first(points)
    # Columns follow POINTS exactly: a meter dropped from POINTS (e.g. the
    # three upstream Cameron interconnects replaced by Cameron Interstate's
    # own terminal meter) drops out of the sheet rather than lingering.
    points = points.reindex(columns=[point_column(p) for p in POINTS])
    points = points.sort_index()
    points.index.name = "gas_day"
    plants = plants_bcfd(points)

    xlsx_notes.write_workbook(args.out, {"Bcfd by plant": plants, "Points (Dth)": points}, notes_lines(),
                              {"UNITS", "WHAT THIS IS", "COVERAGE BY PLANT", "POINTS", "SOURCE"})
    import openpyxl
    wb = openpyxl.load_workbook(args.out)
    for name in ("Bcfd by plant", "Points (Dth)"):
        ws = wb[name]
        for (cell,) in ws.iter_rows(min_row=2, max_col=1):
            cell.number_format = "dd-mmm-yyyy"
        ws.column_dimensions["A"].width = 14
    wb.save(args.out)

    print(f"\nSaved {args.out} - {len(points)} gas day(s).")
    print(plants.tail().to_string())


if __name__ == "__main__":
    main()
