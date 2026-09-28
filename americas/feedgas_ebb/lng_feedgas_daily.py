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

# (plant, platform, pipeline, loc id, label). Platform "km" = Kinder
# Morgan pipeline code; "enbridge" = Enbridge LINK business unit.
POINTS = [
    ("Sabine Pass", "km", "KMLP", "49448", "KMLP -> SP Liquefaction, Cameron Par."),
    ("Sabine Pass", "km", "NGPL", "46622", "NGPL -> Sabine Pass Liquefaction"),
    ("Sabine Pass", "enbridge", "TE", "75866", "Texas Eastern -> Cheniere (Creole Trail), Beauregard Par."),
    ("Plaquemines", "km", "TGP", "55833", "Tennessee Gas -> VG Gator Express, Evangeline Pass"),
    ("Plaquemines", "enbridge", "TE", "74530", "Texas Eastern -> Gator Express"),
    ("Cameron", "km", "TGP", "49446", "Tennessee Gas -> Cameron Interstate, Banken Rd"),
    ("Cameron", "enbridge", "TE", "73882", "Texas Eastern -> Sempra/Cameron"),
    ("Cameron", "enbridge", "TE", "75533", "Texas Eastern -> Cameron Interstate, Beauregard Par."),
    ("Calcasieu Pass", "enbridge", "TE", "74529", "Texas Eastern -> TransCameron, Oak Grove"),
    ("Corpus Christi", "km", "NGPL", "48934", "NGPL -> Cheniere Corpus Christi Pipeline, Sinton"),
    ("Corpus Christi", "km", "TGP", "49861", "Tennessee Gas -> Cheniere Corpus Christi Pipeline, Sinton"),
    ("Freeport", "enbridge", "TE", "79999", "Texas Eastern -> Stratton Ridge"),
    ("Elba Island", "km", "EEC", "660700", "Elba Express -> Elba Liquefaction, Chatham"),
]

COVERAGE_NOTES = {
    "Sabine Pass": "Partial - missing Trunkline/Creole Trail (Energy Transfer) and Transco (Williams).",
    "Plaquemines": "Near complete - Gator Express is fed by Tennessee Gas and Texas Eastern.",
    "Cameron": "Partial - Cameron Interstate's other supply interconnects not yet covered.",
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
    return dict(zip(df["Loc"].astype(str).str.split(".").str[0], to_number(df["Total Scheduled Quantity"])))


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
    return dict(zip(df["Loc"].astype(str), to_number(df["Total_Scheduled_Quantity"])))


def point_column(point):
    plant, platform, pipeline, loc, label = point
    return f"{plant} | {label} ({pipeline} {loc})"


def pull(gas_day):
    """One row of point-level scheduled Dth for the gas day. Points on a
    pipeline that failed to load are left blank, not zero."""
    fetchers = {"km": fetch_km, "enbridge": fetch_enbridge}
    needed = sorted({(p[1], p[2]) for p in POINTS})
    results = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(ignore_https_errors=True, accept_downloads=True, viewport={"width": 1600, "height": 1200})
        page = context.new_page()
        for platform, pipeline in needed:
            try:
                results[(platform, pipeline)] = fetchers[platform](page, pipeline, gas_day)
            except Exception as e:
                print(f"  WARNING: {platform} {pipeline} failed ({type(e).__name__}: {str(e)[:200]})", file=sys.stderr, flush=True)
        browser.close()
    row = {}
    for point in POINTS:
        plant, platform, pipeline, loc, label = point
        values = results.get((platform, pipeline))
        if values is None:
            row[point_column(point)] = float("nan")
        elif loc not in values:
            print(f"  WARNING: {pipeline} point {loc} ({label}) not in today's posting", file=sys.stderr, flush=True)
            row[point_column(point)] = float("nan")
        else:
            row[point_column(point)] = values[loc]
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
              "Operationally Available Capacity postings, Total Scheduled Quantity column."]
    return lines


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--out", default=DEFAULT_OUT)
    parser.add_argument("--gas-day", help="YYYY-MM-DD (default: yesterday, US Central)")
    args = parser.parse_args()
    gas_day = date.fromisoformat(args.gas_day) if args.gas_day else datetime.now(ZoneInfo("America/Chicago")).date() - timedelta(days=1)

    print(f"Pulling LNG feedgas points for gas day {gas_day}...", flush=True)
    new = pull(gas_day)
    if new.notna().sum(axis=1).iloc[0] == 0:
        print("No points retrieved at all - leaving the archive untouched.", file=sys.stderr)
        sys.exit(1)

    points = load_points(args.out)
    points = new if points.empty else new.combine_first(points)
    points = points.reindex(columns=[point_column(p) for p in POINTS] + [c for c in points.columns if c not in {point_column(p) for p in POINTS}])
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
