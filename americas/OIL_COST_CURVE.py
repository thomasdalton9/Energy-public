"""
APPROXIMATE US shale OIL supply cost curve (breakeven WTI $/bbl vs cumulative Mb/d) -> output/Data and Chart Outputs/oil_cost_curve.xlsx.
STANDALONE: reads, never writes, brent_wti_daily.xlsx and brent_forward_curve.xlsx.   Manual workflow oil_cost_curve.yml.

LABEL: APPROXIMATE - built from company disclosures and the Dallas Fed Energy Survey, US shale only, WTI basis.

What PUBLIC sources support (see 'Sources' and 'Not reachable'):
  * BREAKEVENS: no operator publishes a numeric new-well breakeven in text we could open for more than a handful of assets. The one public, numeric,
    basin-by-basin source is the Dallas Fed Energy Survey (annual Q1 special question, first-quarter 2026 fielded 11-19 Mar 2026, 130 firms): the WTI
    price a firm needs (a) to cover operating expenses on existing wells (shut-in breakeven) and (b) to profitably drill a new well, averages by basin
    with min and max. These are executive survey answers, not engineering curves; they are the breakeven inputs here (SOURCED).
  * VOLUMES: EIA STEO regional tables (API v2 route steo): crude oil production by region (Permian, Bakken, Eagle Ford), tight oil by formation,
    and the one-year change of EXISTING-well production (base decline, kb/d a year). SOURCED, read at run time.
  * INVENTORY: Texas Pacific Land 2Q 2026 deck (citing Enverus): remaining locations with <$60/bbl breakeven by basin; with STEO completions this gives
    years of inventory at the current pace (DERIVED).
  * OPERATOR CROSS-CHECKS (not curve inputs): D&C cost per lateral foot (Permian Resources, Matador, Diamondback, Chord), Chord EUR by lateral length.

Curve construction (every number carries a status in the workbook):
  Each basin is split in two tiers: BASE = production that survives one year of natural decline, which stays on at the shut-in breakeven; and
  REPLACEMENT = the barrels that new wells must add each year just to hold output flat (EIA's existing-well decline), which need the new-well breakeven.
  Cumulative volume ranks them by breakeven. Total volume = current production. This answers: what output does each WTI price support 12 months ahead
  if drilling is only maintenance? Growth barrels above maintenance are NOT on the curve (no public source gives their volume at a price).
Mapping: which tiers are in the money at spot, at the 12-month average, at the Brent forward M12 / M24 (converted to WTI with an editable historical
spread) and at +/- an editable shift.

Editable (yellow, survive reruns): everything on 'Inputs' incl. override columns for pasting an own (paid) breakeven set.
Usage: python3 OIL_COST_CURVE.py [--out PATH] [--steo-tsv FILE]   (EIA_API_KEY env var; fallback = 'STEO raw' tab of the committed workbook)
"""
import argparse
import datetime as dt
import os
import sys

import numpy as np
import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402

OUT_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
DEFAULT_OUT = os.path.join(OUT_DIR, "oil_cost_curve.xlsx")
WTI_XLSX = os.path.join(OUT_DIR, "brent_wti_daily.xlsx")
FWD_XLSX = os.path.join(OUT_DIR, "brent_forward_curve.xlsx")

FILL_IN = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
FILL_HEAD = PatternFill("solid", start_color="DDEBF7", end_color="DDEBF7")
BOLD = Font(bold=True)
WRAP = Alignment(wrap_text=True, vertical="top")
LABEL_BANNER = "APPROXIMATE - built from company disclosures and the Dallas Fed Energy Survey, US shale only, WTI basis"

STEO_IDS = ["COPRPM", "COPRBK", "COPREF", "COPRR48", "TOPRPM", "TOPRBK", "TOPREF", "TOPRNI", "TOPRAC", "TOPRMP", "TOPRWF", "TOPRR48", "TOPRL48",
            "COEOPPM", "COEOPBK", "COEOPEF", "CONWPM", "CONWBK", "CONWEF", "NWCPM", "NWCBK", "NWCEF", "RIGSPM", "RIGSBK", "RIGSEF"]

# ------------------------------------------------------------------------------------------------------------------ sources
U_DALLAS_BE = "https://www.dallasfed.org/research/surveys/des/data/breakeven"
U_DALLAS_XLSX = "https://www.dallasfed.org/-/media/documents/research/surveys/des/documents/breakeven.xlsx"
U_DALLAS_Q1 = "https://www.dallasfed.org/research/surveys/des/2026/2601"
U_DALLAS_Q1X = "https://www.dallasfed.org/-/media/Documents/research/surveys/DES/2026/2601/des26q1_charts.xlsx"
U_DALLAS_Q3 = "https://www.dallasfed.org/research/surveys/des/2026/2603"
U_STEO = "https://api.eia.gov/v2/steo/data (EIA Short-Term Energy Outlook, regional series)"
U_TPL = "https://d1io3yog0oux5.cloudfront.net/_44c0a1327e37db57b36e014028684610/texaspacific/db/706/6716/pdf/2Q+2026+TPL+Investor+Presentation+vF.pdf"
U_PR1 = "https://www.sec.gov/Archives/edgar/data/1658566/000165856626000069/ex991prq12026earningsrelea.htm"
U_PR2 = "https://www.sec.gov/Archives/edgar/data/1658566/000165856626000094/ex991prq22026earningsrelea.htm"
U_MTDR = "https://www.sec.gov/Archives/edgar/data/1520006/000152000626000036/a20260630mtdr8ker-ex991.htm"
U_FANG1 = "https://www.sec.gov/Archives/edgar/data/1539838/000153983826000073/diamondbackex992-5x4x26.htm"
U_FANG2 = "https://www.sec.gov/Archives/edgar/data/1539838/000153983826000137/diamondbackex991-8x3x26.htm"
U_CHRD = "https://ir.chordenergy.com/image/CHRD-2Q26-Earnings-Presentation.pdf"
U_NOG = "https://www.sec.gov/Archives/edgar/data/1104485/000119312526237800/d11850dex991.htm"
U_EIA_UP = "https://www.eia.gov/analysis/studies/drilling/pdf/upstream.pdf"
DAL = "Federal Reserve Bank of Dallas"
DOC_Q1 = "Dallas Fed Energy Survey, Q1 2026, special questions (fielded 11-19 Mar 2026, 130 firms)"

# (ID, type, item, value, unit, as-of, publisher, document, url, status, used where)
SOURCES = [
    ("D1", "breakeven", "New-well breakeven, Permian Basin - Midland (average of responses)", 68.57, "USD/bbl WTI", "Q1 2026", DAL, DOC_Q1 + "; 14 responses", U_DALLAS_Q1X, "SOURCED", "Inputs (Permian average uses 'Permian all responses')"),
    ("D2", "breakeven", "New-well breakeven, Permian Basin - Delaware", 63.21, "USD/bbl WTI", "Q1 2026", DAL, DOC_Q1 + "; 19 responses", U_DALLAS_Q1X, "SOURCED", "memo (range)"),
    ("D3", "breakeven", "New-well breakeven, Permian Basin - Other", 70.0, "USD/bbl WTI", "Q1 2026", DAL, DOC_Q1 + "; 12 responses", U_DALLAS_Q1X, "SOURCED", "memo (range)"),
    ("D4", "breakeven", "New-well breakeven, Permian (all responses)", 67, "USD/bbl WTI", "Q1 2026", DAL, "Break-even prices for U.S. oil producers - major shale plays, 'New Wells' sheet, update 25 Mar 2026", U_DALLAS_XLSX, "SOURCED", "Inputs: Permian new-well breakeven"),
    ("D5", "breakeven", "New-well breakeven, Eagle Ford", 63, "USD/bbl WTI", "Q1 2026", DAL, DOC_Q1 + "; 5 responses (small sample)", U_DALLAS_Q1X, "SOURCED", "Inputs: Eagle Ford"),
    ("D6", "breakeven", "New-well breakeven, Other U.S. (shale)", 62.38, "USD/bbl WTI", "Q1 2026", DAL, DOC_Q1 + "; 17 responses", U_DALLAS_Q1X, "SOURCED", "Inputs: Bakken (MAPPING ASSUMPTION - no Bakken line published) and Other shale"),
    ("D7", "breakeven", "New-well breakeven, all responses / large firms (>=10,000 b/d) / small firms", "66 / 59 / 68", "USD/bbl WTI", "Q1 2026", DAL, "Break-even page, 'New Wells' sheet", U_DALLAS_XLSX, "SOURCED", "Cross-checks"),
    ("D8", "breakeven", "Shut-in (operating cost) breakeven: Permian all 39, Eagle Ford 40, Other shale 44.12, Midland 41.54, Delaware 33.82", "39 / 40 / 44.12", "USD/bbl WTI", "Q1 2026", DAL, "Break-even page, 'Existing Wells' sheet (WTI needed to cover operating expenses for existing wells)", U_DALLAS_XLSX, "SOURCED", "Inputs: shut-in breakevens"),
    ("D9", "breakeven", "New-well breakeven range of responses (min-max): Midland 50-90, Delaware 50-85, Permian other 50-85, Eagle Ford 50-70, Other shale 35-85", "see item", "USD/bbl WTI", "Q1 2026", DAL, DOC_Q1 + ", 'Special Question 2'", U_DALLAS_Q1X, "SOURCED", "Inputs: min / max"),
    ("D10", "breakeven", "Prior-year survey (Q1 2025) new-well breakeven: Permian all 65, Eagle Ford 61.67, Other shale 63; shut-in Permian all 39, Eagle Ford 25.83, Other shale 40.81", "see item", "USD/bbl WTI", "Q1 2025", DAL, "Break-even page, history table", U_DALLAS_XLSX, "SOURCED", "Inputs: prior-year series"),
    ("D11", "price expectation", "Dallas Fed Q3 2026 survey: respondents expect WTI $88 at year-end 2026, $79 in two years, $82 in five years (range $70-126 for year-end 2026); WTI spot averaged $98.70 during collection (16-24 Sep 2026). No breakeven question this quarter (special questions: SPR, Persian Gulf exports, fuel spreads, FCF allocation, theft)", "88 / 79 / 82", "USD/bbl WTI", "Q3 2026", DAL, "Dallas Fed Energy Survey, Q3 2026", U_DALLAS_Q3, "SOURCED", "context only"),
    ("E1", "volume", "STEO crude oil production by region: Permian, Bakken, Eagle Ford, Rest of Lower 48 (monthly, Mb/d); tight oil by formation (TOPR*)", "see 'STEO raw'", "Mb/d", "STEO, run date", "US EIA", "Short-Term Energy Outlook, regional oil production series (COPRPM/BK/EF/R48, TOPR*)", U_STEO, "SOURCED", "Inputs: production"),
    ("E2", "decline", "STEO 'Existing Oil Production Change, One-year Trend': Permian COEOPPM, Bakken COEOPBK, Eagle Ford COEOPEF (kb/d, negative = base decline)", "see 'STEO raw'", "kb/d per year", "STEO, run date", "US EIA", "Short-Term Energy Outlook, regional series", U_STEO, "SOURCED", "Inputs: base decline = replacement barrels"),
    ("E3", "activity", "STEO new wells completed per month (NWCPM/NWCBK/NWCEF) and new-well oil production (CONW*)", "see 'STEO raw'", "wells / month; kb/d", "STEO, run date", "US EIA", "Short-Term Energy Outlook, regional series", U_STEO, "SOURCED", "Inputs: inventory years"),
    ("T1", "inventory", "Estimated remaining well locations with <$60/bbl breakeven economics: Permian ~65,000 (Midland + Delaware), Eagle Ford ~13,000, Williston ~5,000, Anadarko ~3,000, DJ ~3,000 (Montney ~34,000, Canada, not used); Enverus-based chart label", "65,000 / 13,000 / 5,000 / 3,000 / 3,000", "locations", "30 Jun 2026 (deck 2Q26)", "Texas Pacific Land (citing Enverus, EIA, company data)", "2Q 2026 TPL Investor Presentation, slide 9 (deck page 8 of text extract); also 8-K ex99.2 5 Aug 2026", U_TPL, "SOURCED (third-party chart label)", "Inputs: inventory locations"),
    ("T2", "volume", "Permian crude production 6.7 MMb/d (4Q 2025 average) per TPL / EIA", 6.7, "Mb/d", "4Q 2025", "Texas Pacific Land (citing EIA)", "2Q 2026 TPL Investor Presentation", U_TPL, "SOURCED", "Cross-check of STEO Permian"),
    ("O1", "well cost", "Permian Resources D&C cost per lateral foot, Q1 2026 (Q4 2025 ~$700, -14% vs 2024)", 685, "USD per lateral ft", "Q1 2026", "Permian Resources", "Q1 2026 earnings release (8-K ex99.1, 6 May 2026)", U_PR1, "SOURCED", "Cross-checks"),
    ("O2", "guidance", "Permian Resources Q2 2026 average daily oil production", 198071, "bbl/d", "Q2 2026", "Permian Resources", "Q2 2026 earnings release (8-K ex99.1, 5 Aug 2026)", U_PR2, "SOURCED", "Cross-checks"),
    ("O3", "well cost", "Matador full-year 2026 D&C cost estimate $785-805 per completed lateral foot ($795 midpoint); Q3 adjacent assets as low as $640 per foot; total 2026 capex midpoint $1.675 bn", "785-805", "USD per completed lateral ft", "Q2 2026", "Matador Resources", "Q2 2026 earnings release (8-K ex99.1, 5 Aug 2026)", U_MTDR, "SOURCED", "Cross-checks"),
    ("O4", "well cost", "Diamondback: lowest Barnett well cost to date under $400 per lateral foot (Barnett; not the Wolfcamp core); 2026 total capex raised to ~$3.9 bn (from ~$3.75 bn), operated D&C ~$3.31 bn, average lateral ~12,900 ft", "<400", "USD per lateral ft", "Q1 / Q2 2026", "Diamondback Energy", "Q1 2026 (8-K ex99.2, 4 May 2026) and Q2 2026 (8-K ex99.1, 3 Aug 2026) releases", U_FANG1, "SOURCED", "Cross-checks"),
    ("O5", "EUR / well cost", "Chord (Bakken, Western Extension case study, Feb 2026 capital assumptions): 2-/3-/4-mile well cost $7.1 / $8.5 / $10.7 MM ($710 / $568 / $532 per ft); oil EUR 450 / 675 / 855 Mbbl (45 / 45 / 43 bbl per ft); F&D $15.78 / $12.62 / $12.46 per bbl. Chord also says 10+ years of low-breakeven inventory (no number) and a >10% improvement in weighted-average inventory breakeven in 2025", "see item", "USD; Mbbl", "Aug 2026 deck", "Chord Energy", "2Q26 Earnings Presentation, 5 Aug 2026 (slides 13-15, 4-mile program)", U_CHRD, "SOURCED", "Cross-checks"),
    ("O6", "breakeven", "Northern Oil & Gas: ~500 gross Duvernay locations with average breakevens below $50 WTI (ALBERTA, Canada - NOT US shale, excluded from the curve)", "<50", "USD/bbl WTI", "26 May 2026", "Northern Oil and Gas", "8-K ex99.1, acquisition release", U_NOG, "SOURCED (excluded)", "none"),
    ("X1", "well cost", "EIA 'Trends in U.S. Oil and Natural Gas Upstream Costs' (IHS study): well costs by play 2006-2015, forecast to 2018", "n/a", "USD", "Mar 2016", "US EIA / IHS", "Upstream cost study", U_EIA_UP, "OPENED - too old", "not used (a decade old; costs and well design have changed)"),
]

NOT_REACHABLE = [
    ("Operator decks with a numeric breakeven (Diamondback, Devon, Occidental, EOG, Coterra, Civitas, Ovintiv, Chevron, ExxonMobil, APA, SM, Crescent, HighPeak, Vital)",
     "Opened their SEC 10-K / 10-Q / 2026 8-K exhibits (keyword windows in discovery_archive/results/oil/): no numeric new-well breakeven in the text. Most give breakevens only in investor-deck charts or footnotes: IR pages for Diamondback, Devon, Permian Resources, Coterra, Matador (hosts did not resolve or no PDF link found), Civitas, SM, Crescent, Magnolia, NOG (DNS failure from Actions), Ovintiv, Chevron, Exxon (404), Occidental and HighPeak (page opens, no deck link found). EOG's only deck link found is Feb 2024 (stale). Not read: numbers in chart images."),
    ("Permian Resources, Devon, EOG tier-1 inventory in years", "Press releases and 10-Ks opened; inventory life is stated qualitatively, not as locations plus years, in the text extracted."),
    ("Kansas City Fed Energy Survey", "Page opened (kansascityfed.org/surveys/energy-survey/); quarterly activity indices and price expectations only, no breakeven question found."),
    ("EIA AEO supply module notes (OGSM)", "Opened (eia.gov OGSM_Assumptions.pdf): methodology for well economics (IP, decline, discount) but no breakeven table by basin."),
    ("IEA Oil 2026, OPEC World Oil Outlook, Rystad, Enverus public pages", "IEA and OPEC pages did not return a document from Actions; Rystad news page and Enverus blog opened with navigation text only, no breakeven figures. Paid curves (Rystad, Enverus) are not public: paste your own into the override columns on 'Inputs'."),
    ("Dallas Fed Energy Survey by operator", "The survey is anonymous: basin averages, min and max only; no per-company breakevens."),
]

SNIPPETS = [
    ("SNIPPET - unopened", "Diamondback 2026 production guidance 500-510 MBO/d (web search result, Feb 2026 release not opened)", "Not used"),
    ("SNIPPET - unopened", "Permian Resources year-to-date acquisitions struck at a weighted-average front-month WTI of $72.50 per barrel (web search summary)", "Not used"),
    ("SNIPPET - unopened", "Devon Q2 2026: 1,359 MBoe/d total and 503 kb/d oil at the top of guidance (web search summary; press release PDF not opened)", "Not used"),
]

# Dallas Fed Energy Survey, annual Q1 special questions: 'Break-even page' workbook (update 25 Mar 2026) and Q1 2026 chart workbook.
# basin -> (shut-in Q1 2026, shut-in Q1 2025, new-well Q1 2026, new-well Q1 2025)
DALLAS = {
    "Permian Basin - Midland": (41.54, 35.44, 68.57, 60.56),
    "Permian Basin - Delaware": (33.82, 33.03, 63.21, 62.22),
    "Permian Basin - Other": (43.63, 45.15, 70.00, 69.75),
    "Permian (all responses)": (39, 39, 67, 65),
    "Eagle Ford": (40, 25.83, 63, 61.67),
    "Other U.S. (Shale)": (44.12, 40.81, 62.38, 63),
    "Other U.S. (Nonshale)": (47, 44.62, 67.82, 65.77),
    "All responses": (43, 41, 66, 65),
    "Large firms (>=10,000 b/d)": (32, 31, 59, 61),
    "Small firms (<10,000 b/d)": (46, 44, 68, 66),
}
DALLAS_RANGE = {  # new-well min, max of responses, Q1 2026 (Special Question 2)
    "Permian Basin - Midland": (50, 90), "Permian Basin - Delaware": (50, 85), "Permian Basin - Other": (50, 85), "Eagle Ford": (50, 70), "Other U.S. (Shale)": (35, 85)}
DALLAS_HIST_YEARS = ["2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026"]
DALLAS_HIST_NEW = {
    "Permian Basin - Midland": [46, 47.33, 47.5, 45.57, 45.8, 51.29, 58.33, 61.61, 60.56, 68.57],
    "Permian Basin - Delaware": [48, 48.92, 48.69, 51.56, 49.25, 49.74, 60.71, 63.81, 62.22, 63.21],
    "Eagle Ford": [47.5, None, 51.27, 46, 46.25, 48, 56.43, None, 61.67, 63],
    "Other U.S. (Shale)": [55.25, 53.57, 48.5, 51, 58.06, 68.7, 60.91, 59.41, 63, 62.38],
    "Permian (all responses)": [48, 50, 50, 49, 50, 52, 61, 65, 65, 67],
    "All responses": [50, 52, 50, 49, 52, 56, 62, 64, 65, 66]}

NPART = 3  # base, replacement, growth tiers per basin
# Basin tiers: (name, STEO production formula parts, decline id, Dallas group for shut-in/new-well, mapping status, inventory locations <$60 (TPL), completions id)
TIERS = [
    ("Permian", ["COPRPM"], "COEOPPM", "Permian (all responses)", "SOURCED", 65000, "NWCPM"),
    ("Bakken", ["COPRBK"], "COEOPBK", "Other U.S. (Shale)", "ASSUMPTION (mapping: Dallas publishes no Bakken line)", 5000, "NWCBK"),
    ("Eagle Ford", ["COPREF"], "COEOPEF", "Eagle Ford", "SOURCED (5 responses)", 13000, "NWCEF"),
    ("Other shale (Niobrara, Anadarko, Austin Chalk, other)", ["TOPRL48", "-TOPRPM", "-TOPRBK", "-TOPREF"], None, "Other U.S. (Shale)", "SOURCED (17 responses)", 6000, None),
]


# ------------------------------------------------------------------------------------------------------------------ data in
def fetch_steo(path_prev, tsv=None):
    rows = []
    if tsv:
        d = pd.read_csv(tsv, sep="\t", header=None, names=["id", "period", "value", "description", "unit"])
        return d[d["id"].isin(STEO_IDS)], f"local file {os.path.basename(tsv)}"
    key = os.environ.get("EIA_API_KEY", "")
    try:
        for i in STEO_IDS:
            r = requests.get("https://api.eia.gov/v2/steo/data", params={
                "api_key": key, "frequency": "monthly", "data[0]": "value", "facets[seriesId][]": i, "start": "2025-01",
                "sort[0][column]": "period", "sort[0][direction]": "asc", "length": 100}, timeout=60)
            r.raise_for_status()
            for x in r.json()["response"]["data"]:
                rows.append((i, x["period"], x["value"], x.get("seriesDescription", ""), x.get("unit", "")))
        d = pd.DataFrame(rows, columns=["id", "period", "value", "description", "unit"])
        d["value"] = pd.to_numeric(d["value"], errors="coerce")
        if d.empty:
            raise ValueError("empty STEO")
        return d, f"EIA API v2 steo, fetched {dt.date.today():%d %b %Y}"
    except Exception as e:  # noqa: BLE001
        print(f"STEO fetch failed ({type(e).__name__}: {e}); using the 'STEO raw' tab of the committed workbook")
        d = pd.read_excel(path_prev, "STEO raw")
        return d[["id", "period", "value", "description", "unit"]], "saved copy ('STEO raw' tab), fetch failed"


def ref_month(steo):
    need = ["COPRPM", "COPRBK", "COPREF", "TOPRL48", "TOPRPM", "TOPRBK", "TOPREF", "COEOPPM", "COEOPBK", "COEOPEF", "NWCPM", "NWCBK", "NWCEF"]
    sets = [set(steo[(steo["id"] == i) & steo["value"].notna()]["period"]) for i in need]
    return max(set.intersection(*sets))


def read_prices():
    w = pd.read_excel(WTI_XLSX, "Data")
    w["date"] = pd.to_datetime(w["date"])
    w = w.dropna(subset=["WTI_USD_per_bbl"]).set_index("date")
    last = w.index.max()
    s = w["WTI_USD_per_bbl"]
    sp = w["Brent_less_WTI"].dropna()
    spreads = {"Latest day": float(sp.iloc[-1]), "Last 12 months mean": float(sp[sp.index > last - pd.DateOffset(years=1)].mean()),
               "2021-2025 mean": float(sp["2021":"2025"].mean()), "2015-2025 mean": float(sp["2015":"2025"].mean()), "2015-2025 median": float(sp["2015":"2025"].median())}
    f = pd.read_excel(FWD_XLSX, "By M")
    f["trade_date"] = pd.to_datetime(f["trade_date"])
    f = f.set_index("trade_date").sort_index()
    fwd = {}
    for m in (1, 3, 6, 12, 18, 24):
        c = f[f"M{m}"].dropna()
        fwd[m] = (float(c.iloc[-1]), c.index[-1])
    return dict(last_date=last, spot=float(s.iloc[-1]), m12=float(s[s.index > last - pd.DateOffset(years=1)].mean()), min12=float(s[s.index > last - pd.DateOffset(years=1)].min()),
                max12=float(s[s.index > last - pd.DateOffset(years=1)].max()), spreads=spreads, fwd=fwd, fwd_latest=f.index.max())


def read_back(path):
    """Yellow cells of the committed workbook so owner edits survive reruns: assumptions by label, tier overrides by name."""
    prev = {"assume": {}, "tier": {}}
    if not os.path.exists(path):
        return prev
    try:
        ws = load_workbook(path)["Inputs"]
        for r in range(1, ws.max_row + 1):
            lab = ws.cell(r, 1).value
            if lab in [a[0] for a in ASSUME]:
                v = ws.cell(r, 2).value
                if isinstance(v, (int, float)) or (isinstance(v, str) and lab.startswith("STEO growth")):
                    prev["assume"][lab] = v
            elif lab in [t[0] for t in TIERS]:
                prev["tier"][lab] = {k: ws.cell(r, c).value for k, c in (("ov_shut", 11), ("ov_new", 12), ("shut", 5), ("new", 6), ("mn", 7), ("mx", 8), ("loc", 13))}
    except Exception as e:  # noqa: BLE001
        print(f"read-back skipped: {type(e).__name__}: {e}")
    return prev


ASSUME = [  # label, default, unit, status, note
    ("Brent-WTI spread used to convert forward Brent to WTI", None, "USD/bbl", "ASSUMPTION", "Default = 2021-2025 mean of Brent less WTI (brent_wti_daily.xlsx). The latest-day spread is far wider; see 'Price context' for all options."),
    ("Price shift for the +/- columns on 'Mapping'", 10, "USD/bbl", "OWNER", "Editable."),
    ("Decline rate for 'Other shale' (share of production per year)", None, "share", "PROXY", "No EIA decline series for these plays: mean of the three basins' decline / production. Replace if you have a better figure."),
    ("Months of legacy decline counted as the replacement need", 1, "months", "SOURCED at 1 (EIA's series is a monthly flow); >1 = ASSUMPTION", "EIA's existing-well change is the DPR-style monthly change in output of wells already producing (about -490 kb/d for the Permian against ~480 kb/d from new wells). 1 = next-month view. Raise it to see more volume exposed to the new-well breakeven (the legacy decline slows as young wells age, so n x monthly overstates)."),
    ("STEO growth horizon month (YYYY-MM)", "2027-12", "YYYY-MM", "DERIVED", "Growth tier = STEO crude production at this month less the reference month (STEO base case, EIA's own price path). Must exist in 'STEO raw'."),
    ("Price ladder start", 30, "USD/bbl", "OWNER", "'Supply ladder' tab."),
    ("Price ladder step", 5, "USD/bbl", "OWNER", "'Supply ladder' tab."),
]


# ------------------------------------------------------------------------------------------------------------------ build
def style_head(ws, row, ncol, fill=FILL_HEAD):
    for c in range(1, ncol + 1):
        cell = ws.cell(row, c)
        cell.font = BOLD
        cell.fill = fill
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def val(steo, i, p):
    s = steo[(steo["id"] == i) & (steo["period"] == p)]["value"]
    return float(s.iloc[0]) if len(s) else float("nan")


def build(out, tsv=None):
    prev = read_back(out)
    steo, steo_note = fetch_steo(out, tsv)
    rm = ref_month(steo)
    pr = read_prices()
    A = {}
    dflt = {ASSUME[0][0]: round(pr["spreads"]["2021-2025 mean"], 2), ASSUME[1][0]: 10, ASSUME[3][0]: 1, ASSUME[4][0]: "2027-12", ASSUME[5][0]: 30, ASSUME[6][0]: 5}
    # python mirror of the tier table
    tiers = []
    for (name, prod_parts, dec_id, group, status, loc, comp) in TIERS:
        prod = sum((-1 if p.startswith("-") else 1) * val(steo, p.lstrip("-"), rm) for p in prod_parts)
        dec = -val(steo, dec_id, rm) / 1000 if dec_id else None
        sh, pf, nw, nwp = DALLAS[group][0], DALLAS[group][1], DALLAS[group][2], DALLAS[group][3]
        mn, mx = DALLAS_RANGE[group] if group in DALLAS_RANGE else (min(DALLAS_RANGE[k][0] for k in DALLAS_RANGE if k.startswith("Permian")), max(DALLAS_RANGE[k][1] for k in DALLAS_RANGE if k.startswith("Permian")))
        t = dict(name=name, prod=prod, dec=dec, group=group, status=status, shut=sh, shut_prior=pf, new=nw, new_prior=nwp, mn=mn, mx=mx, loc=loc,
                 wells=(val(steo, comp, rm) * 12 if comp else None))
        pv = prev["tier"].get(name, {})
        for k in ("shut", "new", "mn", "mx", "loc"):
            if isinstance(pv.get(k), (int, float)):
                t[k] = pv[k]
        t["ov_shut"] = pv.get("ov_shut") if isinstance(pv.get("ov_shut"), (int, float)) else None
        t["ov_new"] = pv.get("ov_new") if isinstance(pv.get("ov_new"), (int, float)) else None
        tiers.append(t)
    rates = [t["dec"] / t["prod"] for t in tiers if t["dec"] is not None]
    dflt[ASSUME[2][0]] = round(float(np.mean(rates)), 4)
    for lab, d, *_ in ASSUME:
        A[lab] = prev["assume"].get(lab, dflt[lab])
    spread = A[ASSUME[0][0]]
    shift = A[ASSUME[1][0]]
    kmon = A[ASSUME[3][0]]
    gmonth = A[ASSUME[4][0]]
    for t in tiers:
        if t["dec"] is None:
            t["dec"] = A[ASSUME[2][0]] * t["prod"]
        t["grow"] = 0.0
        if len(TIERS[tiers.index(t)][1]) == 1:
            g = val(steo, TIERS[tiers.index(t)][1][0], gmonth) - t["prod"]
            t["grow"] = g if not np.isnan(g) else 0.0
        t["be_shut"] = t["ov_shut"] if t["ov_shut"] is not None else t["shut"]
        t["be_new"] = t["ov_new"] if t["ov_new"] is not None else t["new"]
    rows = []
    for t in tiers:
        b = t["name"].split(" (")[0]
        rows.append(dict(tier=f"{b} - base (existing wells after the decline)", kind="Base", vol=t["prod"] - kmon * t["dec"], be=t["be_shut"], prior=t["shut_prior"], basin=t["name"]))
        rows.append(dict(tier=f"{b} - replacement (new wells to offset the decline)", kind="Replacement", vol=kmon * t["dec"], be=t["be_new"], prior=t["new_prior"], basin=t["name"]))
        rows.append(dict(tier=f"{b} - growth (STEO base case to horizon)", kind="Growth", vol=t["grow"], be=t["be_new"], prior=t["new_prior"], basin=t["name"]))
    df = pd.DataFrame(rows)
    df["key"] = df["be"] + np.arange(len(df)) * 1e-6
    s = df.sort_values("key").reset_index(drop=True)
    s["cum"] = s["vol"].cumsum()
    prices = {"Spot WTI (" + f"{pr['last_date']:%d %b %Y}" + ")": pr["spot"], "WTI, 12-month average": pr["m12"],
              "Brent forward M12 less spread": pr["fwd"][12][0] - spread, "Brent forward M24 less spread": pr["fwd"][24][0] - spread}
    mp = []
    for lab, p in prices.items():
        for sh_, suffix in ((0, ""), (-shift, f" - {shift:g}"), (shift, f" + {shift:g}")):
            pp = p + sh_
            ok = s["be"] <= pp
            mp.append(dict(scenario=lab + suffix, price=pp, vol_in=float(s["vol"][ok].sum()), n_in=int(ok.sum()), share=float(s["vol"][ok].sum() / s["vol"].sum()),
                           marginal=float(s["be"][ok].max()) if ok.any() else None, cushion=pp - float(s["be"].max())))
    mv = pd.DataFrame(mp)
    ladder = pd.DataFrame({"WTI": [A[ASSUME[5][0]] + i * A[ASSUME[6][0]] for i in range(19)]})
    ladder["supply"] = [float(s["vol"][s["be"] <= p].sum()) for p in ladder["WTI"]]

    # ------- value tabs
    src = pd.DataFrame(SOURCES, columns=["ID", "Type", "Item", "Value", "Unit", "As-of", "Publisher", "Document", "URL", "Status", "Used where"]).set_index("ID")
    nr = pd.DataFrame(NOT_REACHABLE, columns=["Source", "What happened (Actions, 6 Oct 2026)"]).set_index("Source")
    snip = pd.DataFrame(SNIPPETS, columns=["Status", "Claim seen", "Why not used"]).set_index("Status")
    dal = pd.DataFrame(DALLAS, index=["Shut-in breakeven Q1 2026", "Shut-in breakeven Q1 2025", "New-well breakeven Q1 2026", "New-well breakeven Q1 2025"]).T
    dal.index.name = "Dallas Fed basin (USD/bbl WTI, average of responses)"
    dh = pd.DataFrame({k: v for k, v in DALLAS_HIST_NEW.items()}, index=[f"{y}:Q1" for y in DALLAS_HIST_YEARS]).T
    dh.index.name = "New-well breakeven by Q1 survey (USD/bbl WTI)"
    steo_raw = steo.sort_values(["id", "period"]).reset_index(drop=True)
    steo_raw.insert(0, "key", steo_raw["id"] + "|" + steo_raw["period"].astype(str))
    fw = pd.DataFrame([{"Contract": f"M{m}", "Brent forward USD/bbl": v, "As of trade date": d.strftime("%Y-%m-%d")} for m, (v, d) in pr["fwd"].items()]).set_index("Contract")
    ctx = pd.DataFrame({"WTI spot, last (USD/bbl)": [pr["spot"]], "Last date": [f"{pr['last_date']:%Y-%m-%d}"], "WTI last 12 months mean": [pr["m12"]], "WTI last 12 months min": [pr["min12"]],
                        "WTI last 12 months max": [pr["max12"]], "Brent-WTI spread used (Inputs)": [spread],
                        "Brent forward M12 as WTI": [pr["fwd"][12][0] - spread], "Brent forward M24 as WTI": [pr["fwd"][24][0] - spread]}).T.rename(columns={0: "Value"})
    sp = pd.DataFrame({"Brent less WTI (USD/bbl)": pr["spreads"]})
    cv = s[["tier", "kind", "basin", "vol", "be", "cum", "prior"]].rename(columns={"prior": "Prior-year breakeven (Dallas Q1 2025)", "tier": "Tier (sorted by breakeven)", "kind": "Kind", "basin": "Basin", "vol": "Volume (Mb/d)", "be": "Breakeven (USD/bbl WTI)", "cum": "Cumulative volume (Mb/d)"})
    cv.index = range(1, len(cv) + 1)
    cv.index.name = "Rank"
    notes = [
        "UNITS",
        "Breakeven: US$ per barrel of WTI. Volume: million barrels per day (Mb/d), crude oil production of US shale regions from EIA STEO. Inventory: locations. Well cost: US$ per lateral foot.",
        "",
        "WHAT THIS IS",
        LABEL_BANNER + ". A 12-month-ahead supply curve at maintenance-only drilling: each basin has a BASE tier (what is left after a year of natural decline, which stays on at the Dallas Fed shut-in breakeven) and a REPLACEMENT tier (the barrels new wells must add to hold output flat, at the Dallas Fed new-well breakeven).",
        "The breakevens are the Dallas Fed Energy Survey's basin averages (executive answers, Q1 2026): public and numeric, but a survey, not an engineering curve. Operator filings give D&C cost per foot, EUR and inventory for single assets only (see 'Sources' and 'Cross-checks'); no operator publishes a basin breakeven curve we could open.",
        "NOT on the curve: growth barrels above maintenance (no public source gives their volume at a price), Gulf of America offshore and conventional onshore oil, Canadian oil, rig or crew or takeaway limits, hedges, taxes. The three main basins use STEO regional CRUDE production (includes a small non-shale part); 'Other shale' = STEO tight oil minus Permian, Bakken and Eagle Ford tight oil.",
        "Dallas Fed basin lines: Bakken has none, so it borrows 'Other U.S. (Shale)' (flagged ASSUMPTION); Eagle Ford rests on 5 responses; the Permian uses 'all responses' (the Midland/Delaware split has no production split in STEO).",
        "",
        "HOW TO READ THE TABS",
        "Inputs: assumptions and the tier table (yellow = editable, with override columns for your own or a paid curve). Tiers: live volume and breakeven of the 12 tiers (base, replacement, growth per basin). Curve: tiers sorted by breakeven with cumulative volume. Mapping: which tiers are in the money at spot, the 12-month average, the Brent forward converted to WTI, and +/- the shift. Supply ladder: 12-month output at each WTI price. Cross-checks: operator cost per foot, Dallas Fed vs operator figures, inventory years. Price context: WTI and Brent forward values read from brent_wti_daily.xlsx and brent_forward_curve.xlsx (not edited). Dallas Fed: the survey tables. STEO raw: the EIA series used (volumes read from here by formula).",
        "",
        "STATUS",
        f"Dallas Fed breakevens SOURCED (7 of 8 tier-breakevens on 3 lines; Bakken line is a MAPPING ASSUMPTION), volumes and declines SOURCED (EIA STEO, {steo_note}; reference month {rm}), 'Other shale' decline PROXY, forward-to-WTI spread ASSUMPTION (default 2021-2025 mean).",
        "",
        "EDITING",
        "Yellow cells on 'Inputs' only. Leave override columns blank to use the Dallas Fed value. 'Curve values', 'Mapping values' and 'Ladder values' are the script's own copy of the arithmetic (used by the PNG charts); after editing in Excel the live tabs and native charts follow, the values tabs do not.",
    ]
    titles = {"UNITS", "WHAT THIS IS", "HOW TO READ THE TABS", "STATUS", "EDITING"}
    sheets = {"Sources": src, "Not reachable": nr, "Snippets (unused)": snip, "Dallas Fed": dal, "Dallas Fed history": dh, "STEO raw": steo_raw.set_index("key"),
              "Price context": ctx, "Price forwards": fw, "Spread options": sp, "Curve values": cv, "Mapping values": mv.set_index("scenario"), "Ladder values": ladder.set_index("WTI"),
              "Inputs": pd.DataFrame(), "Tiers": pd.DataFrame(), "Curve": pd.DataFrame(), "Mapping": pd.DataFrame(), "Supply ladder": pd.DataFrame(), "Cross-checks": pd.DataFrame(), "Chart data": pd.DataFrame()}
    os.makedirs(os.path.dirname(out), exist_ok=True)
    xlsx_notes.write_workbook(out, sheets, notes, titles)
    wb = load_workbook(out)
    names = ["Inputs", "Tiers", "Curve", "Mapping", "Supply ladder", "Cross-checks", "Chart data"]
    for n in names:
        del wb[n]
    ws = {n: wb.create_sheet(n, 1 + i) for i, n in enumerate(names)}
    # STEO raw layout: index col A = id, B period, C value  (to_excel with index)
    ref = write_inputs(ws["Inputs"], A, tiers, rm, steo_note, pr, dflt)
    write_tiers(ws["Tiers"], ref)
    write_curve(ws["Curve"], ref)
    write_mapping(ws["Mapping"], ref, pr)
    write_ladder(ws["Supply ladder"], ref)
    write_cross(ws["Cross-checks"], ref, tiers, s, rm, steo)
    write_chart_data(ws["Chart data"], ref, pr)
    for n in ("Sources", "Not reachable", "Snippets (unused)"):
        w = wb[n]
        for col, wd in zip("ABCDEFGHIJK", (7, 16, 70, 18, 16, 16, 28, 50, 60, 22, 36)):
            w.column_dimensions[col].width = wd
        for row in w.iter_rows(min_row=2):
            for c in row:
                c.alignment = WRAP
    wb.save(out + ".tmp")
    os.replace(out + ".tmp", out)
    print(f"{out}: oil cost curve built (STEO reference month {rm}; {steo_note}). Tiers sorted by breakeven:")
    print(s[["tier", "vol", "be", "cum"]].round(2).to_string())
    print(mv.round(2).to_string())
    return s, mv, ladder


def write_inputs(ws, A, tiers, rm, steo_note, pr, dflt):
    ws["A1"] = "Inputs - yellow cells are editable and survive reruns. " + LABEL_BANNER
    ws["A1"].font = BOLD
    ws["A2"] = "Block A: assumptions. Block B: basin tiers (volumes read live from 'STEO raw'; breakevens = Dallas Fed Q1 2026; override columns K-L take your own curve)."
    for j, h in enumerate(["Assumption", "Value", "Unit", "Status", "Note"], 1):
        ws.cell(4, j, h)
    style_head(ws, 4, 5)
    ws.cell(5, 1, "STEO reference month (latest month all series have)")
    ws.cell(5, 2, rm)
    ws.cell(5, 3, "YYYY-MM")
    ws.cell(5, 4, "DERIVED")
    ws.cell(5, 5, f"Volumes and declines are read for this month from 'STEO raw' ({steo_note}). Edit to another month present in that tab.")
    ws.cell(5, 2).fill = FILL_IN
    r = 6
    pos = {}
    for lab, d, unit, status, note in ASSUME:
        ws.cell(r, 1, lab)
        ws.cell(r, 2, A[lab])
        ws.cell(r, 2).fill = FILL_IN
        ws.cell(r, 3, unit)
        ws.cell(r, 4, status)
        ws.cell(r, 5, note)
        pos[lab] = r
        r += 1
    hdr = 13
    ws.cell(hdr - 1, 1, "Block B: basin tiers").font = BOLD
    heads = ["Basin tier", "Dallas Fed line used", "Production (Mb/d, STEO)", "Base decline (Mb/d per year, STEO existing wells)", "Shut-in breakeven (USD/bbl WTI)",
             "New-well breakeven, average (USD/bbl)", "New-well min of responses", "New-well max of responses", "Shut-in, Q1 2025 (prior year)", "New-well, Q1 2025 (prior year)",
             "OVERRIDE shut-in breakeven", "OVERRIDE new-well breakeven", "Remaining locations <$60/bbl (TPL / Enverus)", "Wells completed per year (STEO x12)", "Inventory at current pace (years)",
             "Breakeven status", "Source IDs", "STEO growth to horizon month (Mb/d)"]
    for j, h in enumerate(heads, 1):
        ws.cell(hdr, j, h)
    style_head(ws, hdr, len(heads))
    lk = lambda i: f"SUMIFS('STEO raw'!$D:$D,'STEO raw'!$A:$A,\"{i}|\"&$B$5)"
    for k, t in enumerate(tiers):
        rr = hdr + 1 + k
        name, prod_parts, dec_id, group, status, loc, comp = TIERS[k]
        ws.cell(rr, 1, name)
        ws.cell(rr, 2, group)
        ws.cell(rr, 3, "=" + "".join(("-" if p.startswith("-") else "+") + lk(p.lstrip("-")) for p in prod_parts).lstrip("+"))
        if dec_id:
            ws.cell(rr, 4, f"=-{lk(dec_id)}/1000")
        else:
            ws.cell(rr, 4, f"=$B${pos[ASSUME[2][0]]}*C{rr}")
        for c, key in ((5, "shut"), (6, "new"), (7, "mn"), (8, "mx"), (13, "loc")):
            ws.cell(rr, c, t[key])
            ws.cell(rr, c).fill = FILL_IN
        ws.cell(rr, 9, t["shut_prior"])
        ws.cell(rr, 10, t["new_prior"])
        for c, key in ((11, "ov_shut"), (12, "ov_new")):
            ws.cell(rr, c, t[key])
            ws.cell(rr, c).fill = FILL_IN
        ws.cell(rr, 14, f"={lk(comp)}*12" if comp else "n/a")
        ws.cell(rr, 15, f'=IF(ISNUMBER(N{rr}),M{rr}/N{rr},"n/a")')
        ws.cell(rr, 16, status)
        ws.cell(rr, 17, "D1-D10, E1-E3, T1")
        ws.cell(rr, 18, f"={lk(prod_parts[0])}-C{rr}" if False else (f'=SUMIFS(\'STEO raw\'!$E:$E,\'STEO raw\'!$A:$A,"{prod_parts[0]}|"&$B${pos[ASSUME[4][0]]})-C{rr}' if len(prod_parts) == 1 else 0))
        ws.cell(rr, 18).number_format = "0.000"
        ws.cell(rr, 3).number_format = "0.000"
        ws.cell(rr, 4).number_format = "0.000"
        ws.cell(rr, 15).number_format = "0.0"
    ws.cell(hdr + 6, 1, "Other shale (no STEO forecast of its own, so growth = 0 here): production = TOPRL48 - TOPRPM - TOPRBK - TOPREF (STEO tight oil); decline = proxy rate x production (PROXY); locations = Anadarko ~3,000 + DJ ~3,000 only (Niobrara etc. not itemised by TPL).")
    ws.cell(hdr + 7, 1, "Bakken uses the Dallas Fed 'Other U.S. (Shale)' line (ASSUMPTION: no Bakken line is published). Permian uses 'Permian (all responses)'. Inventory years = TPL/Enverus <$60 locations / STEO completions: below-$60 locations only, so a floor.")
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 26
    for j in range(3, 19):
        ws.column_dimensions[get_column_letter(j)].width = 18
    return dict(hdr=hdr, n=len(tiers), pos=pos, shift=f"Inputs!$B${pos[ASSUME[1][0]]}", spread=f"Inputs!$B${pos[ASSUME[0][0]]}",
                lad0=f"Inputs!$B${pos[ASSUME[5][0]]}", lads=f"Inputs!$B${pos[ASSUME[6][0]]}", kmon=f"Inputs!$B${pos[ASSUME[3][0]]}")


def write_tiers(ws, ref):
    ws["A1"] = "Tiers (live): three tiers per basin (base, replacement, growth). " + LABEL_BANNER
    ws["A1"].font = BOLD
    heads = ["Tier", "Kind", "Basin", "Volume (Mb/d)", "Breakeven used (USD/bbl WTI)", "Sort key", "Rank"]
    for j, h in enumerate(heads, 1):
        ws.cell(3, j, h)
    style_head(ws, 3, 7)
    h0 = ref["hdr"]
    n = ref["n"]
    K = NPART
    for k in range(n):
        ir = h0 + 1 + k
        nm = f"Inputs!$A${ir}"
        b = f'LEFT({nm},FIND(" (",{nm}&" (")-1)'
        for part in range(K):
            r = 4 + K * k + part
            if part == 0:
                ws.cell(r, 1, f'={b}&" - base (existing wells after the decline)"')
                ws.cell(r, 2, "Base")
                ws.cell(r, 4, f"=Inputs!$C${ir}-{ref['kmon']}*Inputs!$D${ir}")
                ws.cell(r, 5, f"=IF(ISNUMBER(Inputs!$K${ir}),Inputs!$K${ir},Inputs!$E${ir})")
            elif part == 1:
                ws.cell(r, 1, f'={b}&" - replacement (new wells to offset the decline)"')
                ws.cell(r, 2, "Replacement")
                ws.cell(r, 4, f"={ref['kmon']}*Inputs!$D${ir}")
                ws.cell(r, 5, f"=IF(ISNUMBER(Inputs!$L${ir}),Inputs!$L${ir},Inputs!$F${ir})")
            else:
                ws.cell(r, 1, f'={b}&" - growth (STEO base case to horizon)"')
                ws.cell(r, 2, "Growth")
                ws.cell(r, 4, f"=Inputs!$R${ir}")
                ws.cell(r, 5, f"=IF(ISNUMBER(Inputs!$L${ir}),Inputs!$L${ir},Inputs!$F${ir})")
            ws.cell(r, 3, f"={nm}")
            ws.cell(r, 6, f"=E{r}+ROW()/1000000")
            ws.cell(r, 7, f"=RANK(F{r},$F$4:$F${3 + K * n},1)")
            ws.cell(r, 4).number_format = "0.000"
    last = 3 + K * n
    ws.cell(last + 2, 1, "Total volume (Mb/d): production + STEO growth")
    ws.cell(last + 2, 4, f"=SUM(D4:D{last})")
    ws.cell(last + 2, 4).number_format = "0.000"
    for c, wd in zip("ABCDEFG", (62, 14, 40, 16, 22, 14, 8)):
        ws.column_dimensions[c].width = wd
    ref["t_first"], ref["t_last"] = 4, last


def write_curve(ws, ref):
    ws["A1"] = "Curve (live): tiers sorted by breakeven. " + LABEL_BANNER
    ws["A1"].font = BOLD
    heads = ["Rank", "Tier", "Kind", "Volume (Mb/d)", "Breakeven (USD/bbl WTI)", "Cumulative volume (Mb/d)", "Prior-year breakeven (Dallas Q1 2025, same tier)"]
    for j, h in enumerate(heads, 1):
        ws.cell(4, j, h)
    style_head(ws, 4, 7)
    a, b = ref["t_first"], ref["t_last"]
    h0 = ref["hdr"]
    n = NPART * ref["n"]
    for k in range(1, n + 1):
        r = 4 + k
        m = f"MATCH($A{r},Tiers!$G${a}:$G${b},0)"
        ws.cell(r, 1, k)
        ws.cell(r, 2, f"=INDEX(Tiers!$A${a}:$A${b},{m})")
        ws.cell(r, 3, f"=INDEX(Tiers!$B${a}:$B${b},{m})")
        ws.cell(r, 4, f"=INDEX(Tiers!$D${a}:$D${b},{m})")
        ws.cell(r, 5, f"=INDEX(Tiers!$E${a}:$E${b},{m})")
        ws.cell(r, 6, f"=D{r}" if k == 1 else f"=F{r - 1}+D{r}")
        # prior-year: shut-in col I for Base, new-well col J for Replacement, basin row = INT((match-1)/2)
        kk = f"INT(({m}-1)/{NPART})+1"
        ws.cell(r, 7, f'=IF(C{r}="Base",INDEX(Inputs!$I${h0 + 1}:$I${h0 + ref["n"]},{kk}),INDEX(Inputs!$J${h0 + 1}:$J${h0 + ref["n"]},{kk}))')
        for c in (4, 6):
            ws.cell(r, c).number_format = "0.000"
    ws.cell(5 + n + 1, 1, "Prior-year column follows the sorted order of the current curve (volumes identical); it is the Dallas Fed Q1 2025 average for the same tier.")
    for c, wd in zip("ABCDEFG", (7, 66, 14, 16, 18, 20, 24)):
        ws.column_dimensions[c].width = wd
    ref["c_first"], ref["c_last"] = 5, 4 + n


def write_mapping(ws, ref, pr):
    ws["A1"] = "Mapping (live): which tiers are in the money (WTI price >= tier breakeven). " + LABEL_BANNER
    ws["A1"].font = BOLD
    ws["A2"] = "Prices: spot and 12-month mean WTI and the Brent forward (converted to WTI with the spread on Inputs) from 'Price context' / 'Price forwards'; shift on Inputs."
    cf, cl = ref["c_first"], ref["c_last"]
    n = cl - cf + 1
    scen = [("WTI spot", "='Price context'!$B$2"), ("WTI, last 12 months mean", "='Price context'!$B$4"),
            ("Brent forward M12 less spread (WTI-equivalent)", f"='Price forwards'!$B$5-{ref['spread']}"), ("Brent forward M24 less spread (WTI-equivalent)", f"='Price forwards'!$B$7-{ref['spread']}")]
    heads = ["Scenario", "WTI price (USD/bbl)"] + [None] * n + ["Volume in the money (Mb/d)", "Share of total volume", "Tiers in the money", "Marginal in-the-money breakeven", "Cushion over the last tier (price - highest breakeven)"]
    ws.cell(4, 1, "Scenario")
    ws.cell(4, 2, "WTI price (USD/bbl)")
    for k in range(n):
        ws.cell(4, 3 + k, f"=Curve!$B${cf + k}")
        ws.cell(5, 3 + k, f"=Curve!$E${cf + k}")
    ws.cell(5, 1, "Tier breakeven (USD/bbl)")
    ws.cell(6, 1, "Tier volume (Mb/d)")
    for k in range(n):
        ws.cell(6, 3 + k, f"=Curve!$D${cf + k}")
    e = 3 + n
    for j, h in enumerate(["Volume in the money (Mb/d)", "Share of total volume", "Tiers in the money (of " + str(n) + ")", "Marginal in-the-money breakeven (USD/bbl)", "Cushion over the highest breakeven (USD/bbl)"], e):
        ws.cell(4, j, h)
    style_head(ws, 4, e + 4)
    r = 7
    first = get_column_letter(3)
    lastc = get_column_letter(2 + n)
    for lab, f in scen:
        for sign, suf in ((0, ""), (-1, " minus shift"), (1, " plus shift")):
            ws.cell(r, 1, lab + suf)
            ws.cell(r, 2, f + ("" if sign == 0 else f"{'-' if sign < 0 else '+'}{ref['shift']}"))
            ws.cell(r, 2).number_format = "0.0"
            for k in range(n):
                c = get_column_letter(3 + k)
                ws.cell(r, 3 + k, f'=IF($B{r}>={c}$5,"in","out")')
            ws.cell(r, e, f"=SUMPRODUCT(($B{r}>={first}$5:{lastc}$5)*{first}$6:{lastc}$6)")
            ws.cell(r, e + 1, f"={get_column_letter(e)}{r}/SUM({first}$6:{lastc}$6)")
            ws.cell(r, e + 2, f'=COUNTIF({first}{r}:{lastc}{r},"in")')
            ws.cell(r, e + 3, f'=IF({get_column_letter(e + 2)}{r}=0,"none",SUMPRODUCT(MAX(($B{r}>={first}$5:{lastc}$5)*{first}$5:{lastc}$5)))')
            ws.cell(r, e + 4, f"=$B{r}-MAX({first}$5:{lastc}$5)")
            ws.cell(r, e).number_format = "0.00"
            ws.cell(r, e + 1).number_format = "0%"
            ws.cell(r, e + 3).number_format = "0.0"
            ws.cell(r, e + 4).number_format = "0.0"
            r += 1
    ws.cell(r + 1, 1, "'in' = the tier's breakeven is at or below the price, so the volume is supported for the next 12 months at maintenance drilling. A price below the new-well breakevens but above shut-in keeps the base tiers but loses the replacement tiers (output falls by the decline).")
    ws.column_dimensions["A"].width = 54
    ws.column_dimensions["B"].width = 14
    for k in range(n + 5):
        ws.column_dimensions[get_column_letter(3 + k)].width = 20
    ws.row_dimensions[4].height = 60
    ref["map_rows"] = (7, r - 1)
    ref["map_cols"] = (e, e + 4)


def write_ladder(ws, ref):
    ws["A1"] = "Supply ladder (live): output 12 months ahead at maintenance-only drilling, by WTI price. " + LABEL_BANNER
    ws["A1"].font = BOLD
    for j, h in enumerate(["WTI (USD/bbl)", "Supported volume (Mb/d)", "Share of current production"], 1):
        ws.cell(3, j, h)
    style_head(ws, 3, 3)
    cf, cl = ref["c_first"], ref["c_last"]
    for i in range(19):
        r = 4 + i
        ws.cell(r, 1, f"={ref['lad0']}+{i}*{ref['lads']}")
        ws.cell(r, 2, f"=SUMPRODUCT((A{r}>=Curve!$E${cf}:$E${cl})*Curve!$D${cf}:$D${cl})")
        ws.cell(r, 3, f"=B{r}/SUM(Curve!$D${cf}:$D${cl})")
        ws.cell(r, 2).number_format = "0.00"
        ws.cell(r, 3).number_format = "0%"
    for c in "ABC":
        ws.column_dimensions[c].width = 26
    ref["lad_rows"] = (4, 22)


def write_cross(ws, ref, tiers, s, rm, steo):
    ws["A1"] = "Cross-checks: figures NOT used as curve inputs. " + LABEL_BANNER
    ws["A1"].font = BOLD
    ws["A3"] = "A. Dallas Fed new-well breakeven vs operator-disclosed economics"
    ws["A3"].font = BOLD
    rows = [
        ("Dallas Fed Q1 2026, large firms (>=10,000 b/d) / small firms new-well breakeven", "59 / 68", "USD/bbl WTI", "D7"),
        ("Dallas Fed Q1 2026, all responses new-well / shut-in breakeven", "66 / 43", "USD/bbl WTI", "D7, D8"),
        ("Permian Resources D&C cost per lateral foot, Q1 2026", 685, "USD/ft", "O1"),
        ("Matador full-year 2026 D&C per completed lateral foot (midpoint of 785-805)", 795, "USD/ft", "O3"),
        ("Diamondback Barnett (not core) lowest well cost to date", "<400", "USD/ft", "O4"),
        ("Chord Bakken 2-/3-/4-mile D&C per foot", "710 / 568 / 532", "USD/ft", "O5"),
        ("Chord Bakken 2-/3-/4-mile oil EUR", "450 / 675 / 855", "Mbbl", "O5"),
        ("TPL / Enverus remaining Permian locations <$60 breakeven", 65000, "locations", "T1"),
        ("NOG Duvernay (Canada, excluded) average breakeven", "<50", "USD/bbl WTI", "O6"),
    ]
    for j, h in enumerate(["Item", "Value", "Unit", "Source ID"], 1):
        ws.cell(4, j, h)
    style_head(ws, 4, 4)
    for i, rw in enumerate(rows):
        for j, v in enumerate(rw, 1):
            ws.cell(5 + i, j, v)
    r0 = 5 + len(rows) + 2
    ws.cell(r0, 1, "B. Volume cross-check: STEO Permian crude production (Mb/d) at the reference month vs TPL's 4Q 2025 Permian figure (6.7)").font = BOLD
    ws.cell(r0 + 1, 1, "STEO Permian crude production, reference month")
    ws.cell(r0 + 1, 2, f"=Inputs!$C${ref['hdr'] + 1}")
    ws.cell(r0 + 2, 1, "TPL (EIA) Permian crude, 4Q 2025 average")
    ws.cell(r0 + 2, 2, 6.7)
    ws.cell(r0 + 4, 1, "C. Inventory at the current pace (TPL <$60 locations / STEO completions)").font = BOLD
    for j, h in enumerate(["Basin", "Locations <$60", "Wells completed per year", "Years"], 1):
        ws.cell(r0 + 5, j, h)
    style_head(ws, r0 + 5, 4)
    for k in range(3):
        ir = ref["hdr"] + 1 + k
        ws.cell(r0 + 6 + k, 1, f"=Inputs!$A${ir}")
        ws.cell(r0 + 6 + k, 2, f"=Inputs!$M${ir}")
        ws.cell(r0 + 6 + k, 3, f"=Inputs!$N${ir}")
        ws.cell(r0 + 6 + k, 4, f"=Inputs!$O${ir}")
    ws.cell(r0 + 10, 1, "Caveat: the Dallas Fed averages (about $63-67 new-well) sit ABOVE TPL's $60 inventory cut-off, so most future locations need a price above the survey average only if they are weaker than the <$60 set; the two are different bases and are not combined.")
    ws.column_dimensions["A"].width = 80
    for c in "BCD":
        ws.column_dimensions[c].width = 22


def write_chart_data(ws, ref, pr):
    ws["A1"] = "Chart data (live): cost curve steps, price lines and the supply ladder. " + LABEL_BANNER
    ws["A1"].font = BOLD
    cf, cl = ref["c_first"], ref["c_last"]
    n = cl - cf + 1
    ws["A3"] = "WTI spot"
    ws["B3"] = "='Price context'!$B$2"
    ws["A4"] = "WTI 12-month mean"
    ws["B4"] = "='Price context'!$B$4"
    ws["A5"] = "Brent forward M12 as WTI"
    ws["B5"] = f"='Price forwards'!$B$5-{ref['spread']}"
    for j, h in enumerate(["x: cumulative volume (Mb/d)", "y: breakeven (Dallas Fed Q1 2026)", "x prior", "y: same tiers, Dallas Fed Q1 2025", "x line", "y WTI spot", "y WTI 12-month mean", "y Brent M12 as WTI"], 4):
        ws.cell(7, j, h)
    style_head(ws, 7, 11)
    ws["E6"] = "Cost curve (Dallas Fed Q1 2026 breakevens)"
    ws["G6"] = "Same tiers, Dallas Fed Q1 2025 (prior year)"
    ws["H6"] = "WTI spot"
    ws["I6"] = "WTI, 12-month mean"
    ws["J6"] = "Brent forward M12 less spread"
    for k in range(n):
        for part in (0, 1):
            r = 8 + 2 * k + part
            cr = cf + k
            x = "0" if (k == 0 and part == 0) else (f"Curve!$F${cr - 1}" if part == 0 else f"Curve!$F${cr}")
            ws.cell(r, 4, f"={x}")
            ws.cell(r, 5, f"=Curve!$E${cr}")
            ws.cell(r, 6, f"={x}")
            ws.cell(r, 7, f"=Curve!$G${cr}")
    last = 7 + 2 * n
    ws.cell(8, 8, 0)
    ws.cell(9, 8, f"=D{last}")
    for rr in (8, 9):
        ws.cell(rr, 9, "=$B$3")
        ws.cell(rr, 10, "=$B$4")
        ws.cell(rr, 11, "=$B$5")
    # NB: x of line series is column H; y are columns I, J, K
    ws.cell(7, 8, "x line")
    ws.cell(7, 9, "y WTI spot")
    ws.cell(7, 10, "y WTI 12-month mean")
    ws.cell(7, 11, "y Brent M12 as WTI")
    ws["H6"] = ""
    ws["I6"] = "WTI spot"
    ws["J6"] = "WTI, 12-month mean"
    ws["K6"] = "Brent forward M12 less spread"
    a, b = ref["lad_rows"]
    ws.cell(last + 3, 1, "Supply ladder").font = BOLD
    for j, h in enumerate(["WTI (USD/bbl)", "Supported volume (Mb/d)"], 1):
        ws.cell(last + 4, j, h)
    style_head(ws, last + 4, 2)
    for i in range(a, b + 1):
        ws.cell(last + 4 + i - a + 1, 1, f"='Supply ladder'!A{i}")
        ws.cell(last + 4 + i - a + 1, 2, f"='Supply ladder'!B{i}")
    ref["lad_chart"] = (last + 4, last + 4 + (b - a + 1))
    ref["step"] = (8, last)
    ws.column_dimensions["A"].width = 30
    for j in range(2, 12):
        ws.column_dimensions[get_column_letter(j)].width = 22


def chart_layout():
    """Cell positions on 'Chart data' used by add_charts.py (native charts)."""
    n = NPART * len(TIERS)
    last = 7 + 2 * n
    return {"step_first": 8, "step_last": last, "lad_hdr": last + 4, "lad_last": last + 4 + 19, "banner": LABEL_BANNER}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--steo-tsv", default=None)
    args = ap.parse_args()
    build(args.out, args.steo_tsv)


if __name__ == "__main__":
    main()
