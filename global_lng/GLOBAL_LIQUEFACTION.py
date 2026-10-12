"""
Global LNG liquefaction capacity (existing nameplate, projects under construction / FID / planned, closures) - the
'Global supply' section of the master FUTURE capacity table (shared row schema, see the 'Projects' sheet).

    python3 global_lng/GLOBAL_LIQUEFACTION.py            # incremental (default)
    python3 global_lng/GLOBAL_LIQUEFACTION.py --full     # re-read every whole file

Output: output/Data and Chart Outputs/global_lng_liquefaction_capacity.xlsx

What is automatic and what is curated
  * United States: the EIA 'U.S. Liquefaction Capacity' workbook (newest quarterly file found) is parsed train by train
    (nameplate = baseload mtpa as published by EIA, peak in the notes, status, in-service date). The file is downloaded
    with a conditional request and re-parsed only when it changed (ETag / Last-Modified / SHA-256 on 'Releases'); its
    parsed rows are kept on the 'EIA rows' sheet, which is the history store between runs.
  * Everything outside the United States is a curated table in this script (CURATED / EXISTING_OTHER): each number was
    read in an opened document (company release, QatarEnergy deck, DISR Resources and Energy Quarterly, GIIGNL
    Annual Report 2026, press copy of a CEO statement), and each row names that document. The documents are watched:
    'Releases' keeps their Last-Modified / ETag / size and says CHANGED when a source file was replaced, i.e. the
    curated rows need a fresh read. A newer Resources and Energy Quarterly edition is detected from DISR's listing.
  * History (committed workbook): first_seen / last_seen, and every change of status / expected date / capacity of a
    project goes to 'History' (old -> new, date seen). A project that leaves the sources keeps its row (last_seen stops).

Units: capacity_value in mtpa as published; capacity_bcfd = mtpa x 1.36 bcm/y per mtpa (DISR's rule, REQ note c) x
35.3147 / 365.25 = 0.1315 Bcf/d per mtpa (DERIVED, not a source figure).

Dates: where a source gives only a year, half-year or quarter, the monthly series places the start at a stated point
of that period (year -> July, H1 -> April, H2 -> October, Qn -> middle month): the 'Placement' sheet lists every
project's placement and rule. A project with no date in any opened source is NOT in the dated monthly series; it is
carried in the 'undated' columns. No ramp-up is applied in the sourced series; the editable yellow 'ramp months' input
on 'Capacity by month' (ASSUMPTION) feeds only the separate 'with ramp' columns.
"""
import argparse
import datetime as dt
import hashlib
import io
import os
import re
import sys

import numpy as np
import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.chart import AreaChart, BarChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_charts  # noqa: E402
import xlsx_notes  # noqa: E402

OUT_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
PNG_DIR = os.path.join(ROOT, "output", "PNG Charts")
OUT = os.path.join(OUT_DIR, "global_lng_liquefaction_capacity.xlsx")
BASENAME = "global_lng_liquefaction_capacity"

TODAY = dt.date.today()
NOW_M = pd.Timestamp(TODAY.year, TODAY.month, 1)
FIRST_M, LAST_M = pd.Timestamp("2021-01-01"), pd.Timestamp("2030-12-01")
MTPA_TO_BCFD = 0.1315            # 1 mtpa = 1.36 bcm/y x 35.3147 / 365.25 Bcf/d  (DERIVED)
RAMP_DEFAULT = 12                # months, ASSUMPTION (editable on 'Capacity by month')

SCHEMA = ["region", "country", "asset_type", "project", "operator", "fuel_or_type", "capacity_value", "capacity_unit",
          "capacity_mw", "capacity_bcfd", "status", "expected_online", "online_date_basis", "first_seen", "last_seen",
          "source_publisher", "source_document", "source_url", "as_of_date", "evidence", "notes"]

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36",
      "Accept": "*/*", "Accept-Language": "en"}

EIA_BASE = "https://www.eia.gov/naturalgas/importsexports/liquefactioncapacity/U.S.liquefactioncapacity_{y}_Q{q}.xlsx"
REQ_LIST = "https://www.industry.gov.au/publications/resources-and-energy-quarterly"
REQ_EDITION = "september-2026"

# ---------------------------------------------------------------------------------------------------------------------
# Documents opened for the curated rows (publisher, document, url, as-of date). 'watch' = url fetched for change detection.
# ---------------------------------------------------------------------------------------------------------------------
SRC = {
    "REQ": ("Department of Industry, Science and Resources (Australia)",
            "Resources and Energy Quarterly, September 2026 (section 6 LNG; map 'Nameplate capacity, million tonnes per annum')",
            "https://www.industry.gov.au/sites/default/files/2026-10/resources-and-energy-quarterly-september-2026.pdf", "2026-10-02"),
    "QE_DECK": ("QatarEnergy", "Update to Investors, November 2025 (North Field expansion programme, p.9)",
                "https://www.qatarenergy.qa/en/MediaCenter/Publications/QatarEnergy_Investors_Presentation_November_2025.pdf", "2025-11-01"),
    "QE_AR24": ("QatarEnergy", "Annual Review 2024",
                "https://www.qatarenergy.qa/en/MediaCenter/Publications/%E2%80%8BQatarEnergy%20Annual%20Review%202024.pdf", "2025-06-01"),
    "CHEN_Q2": ("Cheniere Energy", "Cheniere Reports Second Quarter 2026 Results (6 Aug 2026)",
                "https://lngir.cheniere.com/_assets/_75dc2c9a1fb706f14e67cee93b0e9e5a/cheniere/news/2026-08-06_Cheniere_Reports_Second_Quarter_2026_Results_and_343.pdf", "2026-08-06"),
    "GOLAR_Q2": ("Golar LNG", "Interim results for the period ended June 30, 2026",
                 "https://mfn.se/one/a/golar-lng/golar-lng-limited-interim-results-for-the-period-ended-june-30-2026-b168e773", "2026-08-13"),
    "ENI_FID": ("Eni", "Eni announces Final Investment Decision for Mozambique's Coral North project (2 Oct 2025)",
                "https://www.eni.com/en-IT/media/press-release/2025/10/pr-eni-announces-final-investment-decision-mozambique-coral-north-project.html", "2025-10-02"),
    "NLNG_FF": ("Nigeria LNG Limited", "Facts and Figures 2024 (six trains 22 MTPA; Train 7 lifts capacity to 30 MTPA; FID Dec 2019)",
                "https://www.nlng.com/documents/2024_NLNG_FACTS_FIGURESB_compressed.pdf", "2025-01-01"),
    "NOVATEK_FID": ("Novatek", "FID release, Arctic LNG 2 (2019: three trains of 6.6 mtpa; schedule since changed)",
                    "https://www.novatek.ru/common/upload/press/FID_Arctic_LNG_2_Eng.pdf", "2019-09-01"),
    "SHELL_LNGC": ("Shell Canada Energy / Shell plc (release text as republished by Hellenic Shipping News)",
                   "Shell takes final investment decision to double LNG Canada capacity (29 Sep 2026)",
                   "https://www.hellenicshippingnews.com/shell-takes-final-investment-decision-to-double-lng-canada-capacity", "2026-09-29"),
    "MOZ_SITE": ("Mozambique LNG (TotalEnergies-operated project site)", "Project page: development (two liquefaction units, 13.12 MTPA)",
                 "https://mozambiquelng.co.mz/about-the-mozambique-liquefied-natural-gas-project/development/", "2026-10-12"),
    "WOODFIBRE": ("Woodfibre LNG", "Construction page (2.1 million tonnes of LNG per year; substantial completion expected by 2027)",
                  "https://woodfibrelng.ca/construction/", "2026-10-12"),
    "CHEN_SP": ("Cheniere Energy", "Cheniere Reports Second Quarter 2026 Results (Sabine Pass Stage 5 expansion)",
                "https://lngir.cheniere.com/_assets/_75dc2c9a1fb706f14e67cee93b0e9e5a/cheniere/news/2026-08-06_Cheniere_Reports_Second_Quarter_2026_Results_and_343.pdf", "2026-08-06"),
}

# Documents opened but NOT sources of any Projects number: cross-check (GIIGNL) or background / dates for Tried items.
XSRC = {
    "GIIGNL": ("GIIGNL (used only where no official figure was opened; named as publisher)",
               "GIIGNL Annual Report 2026 edition (2025 data): liquefaction plants table and 'Liquefaction' regional sections",
               "https://giignl-documents.s3.fr-par.scw.cloud/public/giignl-2026-annual-report.pdf", "2026-05-20"),
    "WDS_Q2": ("Woodside Energy", "Second Quarter Report 2026 (period ended 30 June 2026)",
               "https://www.woodside.com/docs/default-source/investor-documents/quarterly-and-half-yearly-pdfs-and-data-tables/2026/second-quarter-2026-report.pdf", "2026-07-29"),
    "WDS_H1": ("Woodside Energy", "Half-Year Report 2026 (period ended 30 June 2026)",
               "https://www.woodside.com/docs/default-source/investor-documents/quarterly-and-half-yearly-pdfs-and-data-tables/2026/half-year-2026-report.pdf?sfvrsn=df6d6ddc_6", "2026-08-26"),
    "TTE_Q2": ("TotalEnergies", "2Q26 results press release (ECA LNG start-up early July 2026)",
               "https://totalenergies.com/system/files/documents/totalenergies_pr-results-2q26_2026_en.pdf", "2026-07-22"),
    "PEMBINA_Q2": ("Pembina Pipeline Corporation", "Q2 2026 results news release (Cedar LNG first exports late 2028)",
                   "https://www.pembina.com/media-centre/news/details/fa13b9e4-bbc0-44d0-9ab6-88c72cbd54e5", "2026-07-30"),
    "REUTERS_QE": ("Reuters (copy on BOE Report)", "QatarEnergy's LNG expansion project poised to start operations in November, sources say (7 Oct 2026)",
                   "https://boereport.com/2026/10/07/qatarenergys-lng-expansion-project-poised-to-start-operations-in-november-sources-say/amp/", "2026-10-07"),
}

# ---------------------------------------------------------------------------------------------------------------------
# Curated rows outside the United States EIA table.
# row = dict(country, project, operator, cap (mtpa, None = not stated in an opened document), status, when, src, notes,
#            evidence ('SOURCED'/'DERIVED'), start ('YYYY' or 'YYYY-MM' for existing plants), end, offline_from, kind)
#   when: 'YYYY-MM' | 'YYYY' | 'YYYY-H1' | 'YYYY-H2' | 'YYYY-Q1'..'Q4' | None
#   kind: 'existing' (operating nameplate), 'project'
# ---------------------------------------------------------------------------------------------------------------------
def ex(country, project, operator, cap, start, src, notes="", evidence="SOURCED", status="OPERATING", end=None,
       offline_from=None, start_known=True):
    return dict(country=country, project=project, operator=operator, cap=cap, status=status,
                when=str(start) if start_known else None, src=src, notes=notes, evidence=evidence, end=end,
                offline_from=offline_from, kind="existing", start_known=start_known)


def pj(country, project, operator, cap, status, when, src, notes="", evidence="SOURCED"):
    return dict(country=country, project=project, operator=operator, cap=cap, status=status, when=when, src=src,
                notes=notes, evidence=evidence, end=None, offline_from=None, kind="project", start_known=True)


HORMUZ = "Export routes depend on the Strait of Hormuz, which was disrupted in 2026 (DISR REQ Sep 2026): nameplate unchanged, exports constrained."

CURATED = [
    # ---- Qatar (QatarEnergy deck, DISR REQ; damaged trains and NFE timing from the labelled publishers) -------------
    ex("Qatar", "Existing LNG trains at Ras Laffan (14 trains) excl. the two damaged trains", "QatarEnergy LNG", 64.2, "1996", "QE_DECK",
       "77 mtpa current (QatarEnergy Nov 2025 deck: 14 trains) less 12.8 mtpa damaged (next row, GIIGNL figure). DISR REQ Sep 2026: capacity 'largely intact', ~80% of capacity within a month of the "
       "Strait reopening. " + HORMUZ, "DERIVED"),
    ex("Qatar", "Ras Laffan LNG Trains 4 and 6 (damaged by missile strikes 18 Mar 2026)", "QatarEnergy LNG", 12.8, "1996", "GIIGNL",
       "NUMBER FROM GIIGNL (publisher; no QatarEnergy document with it was reachable): strikes damaged LNG Trains 4 and 6, removing 12.8 mtpa, repairs three to five years per QatarEnergy's March 2026 statement as quoted by GIIGNL; "
       "Reuters 7 Oct 2026: repairs could take three years. DISR REQ Sep 2026 confirms missile damage mid-March 2026 (no mtpa). No return date is stated.",
       status="SUSPENDED", offline_from="2026-03"),
    pj("Qatar", "North Field East expansion, Train 1 of 4 (8 mtpa)", "QatarEnergy LNG", 8.0, "UNDER_CONSTRUCTION", "2027-H1", "REUTERS_QE",
       "Date: QatarEnergy CEO, production expected in H1 2027 (Reuters copy, 7 Oct 2026; same report: train ready to start in November 2026, LNG from it in Q1 2027, per two unnamed sources). "
       "Capacity: NFE = 32 mtpa in 4 trains of 8 (QatarEnergy deck Nov 2025). GIIGNL 2026: first two NFE trains had been expected late 2026, now likely postponed. " + HORMUZ),
    pj("Qatar", "North Field East expansion, Trains 2-4 (3 x 8 mtpa)", "QatarEnergy LNG", 24.0, "UNDER_CONSTRUCTION", None, "QE_DECK",
       "4 trains x 8 mtpa = 32 mtpa (deck p.9); no per-train date in an opened QatarEnergy document. Programme target 142 mtpa by 2030 (deck). "
       "DISR REQ Sep 2026 (NexantECA): Qatari capacity expected to stay below 90 Mt until 2029 and reach 133 Mt by 2031 (was 142 Mt by 2031)."),
    pj("Qatar", "North Field South expansion (2 trains x 8 mtpa)", "QatarEnergy LNG", 16.0, "UNDER_CONSTRUCTION", None, "QE_DECK",
       "16 mtpa, 2 trains of 8 mtpa (deck p.9). No start date in an opened QatarEnergy document; part of the 142 mtpa-by-2030 programme."),
    pj("Qatar", "North Field West expansion (2 trains x 8 mtpa)", "QatarEnergy LNG", 16.0, "PLANNED", None, "QE_DECK",
       "16 mtpa, 2 trains (deck p.9; announced Feb 2024). Programme target 142 mtpa by 2030. No FID or per-train date in an opened document."),
    # ---- Canada ---------------------------------------------------------------------------------------------------
    ex("Canada", "LNG Canada Phase 1 (2 trains)", "LNG Canada (Shell 40%)", 14.0, "2025", "SHELL_LNGC",
       "Capacity: Shell release 29 Sep 2026 (Phase 2 raises capacity 'from 14 mtpa to 28 mtpa'). Start year 2025 from GIIGNL 2026 (Trains 1 and 2, 7.0 each, online in 2025; month not stated).",
       status="COMMISSIONED_RECENT"),
    ex("Canada", "Tilbury LNG", "FortisBC", 0.3, "1971", "GIIGNL", "Small-scale plant, Delta BC (GIIGNL number)."),
    pj("Canada", "LNG Canada Phase 2 (2 trains, +14 mtpa)", "LNG Canada (Shell 40%)", 14.0, "FID", None, "SHELL_LNGC",
       "FID 29 Sep 2026: doubles capacity from 14 to 28 mtpa; 'commercial operations expected to begin in the early 2030s' (no month; not in the dated series)."),
    pj("Canada", "Woodfibre LNG", "Woodfibre LNG", 2.1, "UNDER_CONSTRUCTION", "2027", "WOODFIBRE",
       "Company: designed to produce 2.1 million tonnes of LNG per year; substantial completion expected by 2027 (GIIGNL: 65% complete March 2026)."),
    pj("Canada", "Cedar LNG (floating, Kitimat)", "Cedar LNG (Haisla Nation / Pembina)", None, "UNDER_CONSTRUCTION", "2028-Q4", "PEMBINA_Q2",
       "Pembina Q2 2026: first exports expected late 2028; vessel more than 70% complete. NO capacity figure in any opened document (Pembina, Cedar site, GIIGNL): left blank, not in the totals ('Tried')."),
    pj("Canada", "Ksi Lisims LNG (floating, 12 mtpa)", "Ksi Lisims LNG (Nisga'a Nation / Rockies LNG / Western LNG)", 12.0, "PERMITTED", None, "GIIGNL",
       "NUMBER FROM GIIGNL (no official figure opened): 12 MTPA FLNG, federal approvals, fast-tracked; no FID reported there."),
    # ---- Mexico --------------------------------------------------------------------------------------------------
    pj("Mexico", "Energia Costa Azul LNG Phase 1 (1 train, 3.25 mtpa)", "Sempra Infrastructure", 3.25, "COMMISSIONED_RECENT", "2026-07", "GIIGNL",
       "Capacity 3.25 mtpa from GIIGNL 2026 (no official figure opened). Start: TotalEnergies 2Q26 release says ECA LNG 'started-up early July' 2026."),
    ex("Mexico", "Altamira FLNG 1 (floating)", "New Fortress Energy", 1.4, "2024", "GIIGNL", "GIIGNL plants table: operating since 2024."),
    # ---- Australia (DISR nameplate map; NWS, Pluto Train 2) -----------------------------------------------------------
    ex("Australia", "Gorgon LNG", "Chevron", 15.6, "2016", "REQ", "DISR REQ Sep 2026 nameplate map (labels read from the PDF text; plant-to-label match checked against the 88 Mtpa total)."),
    ex("Australia", "Wheatstone LNG", "Chevron", 8.9, "2017", "REQ", "DISR REQ Sep 2026 nameplate map."),
    ex("Australia", "Ichthys LNG", "INPEX", 8.9, "2018", "REQ", "DISR REQ Sep 2026 nameplate map."),
    ex("Australia", "Pluto LNG Train 1", "Woodside", 4.7, "2012", "REQ", "DISR REQ Sep 2026 nameplate map (GIIGNL: 4.9)."),
    ex("Australia", "Prelude FLNG", "Shell", 3.6, "2019", "REQ", "DISR REQ Sep 2026 nameplate map."),
    ex("Australia", "Darwin LNG", "Santos", 3.7, "2006", "REQ", "DISR REQ Sep 2026 nameplate map. Barossa gas replaced Bayu-Undan; first Barossa cargo Jan 2026 (GIIGNL)."),
    ex("Australia", "Australia Pacific LNG (APLNG)", "Origin / ConocoPhillips", 9.0, "2016", "REQ", "DISR REQ Sep 2026 nameplate map."),
    ex("Australia", "Gladstone LNG (GLNG)", "Santos", 7.8, "2015", "REQ", "DISR REQ Sep 2026 nameplate map."),
    ex("Australia", "Queensland Curtis LNG (QCLNG)", "Shell", 8.5, "2015", "REQ", "DISR REQ Sep 2026 nameplate map."),
    ex("Australia", "North West Shelf LNG, four remaining trains", "Woodside", 14.5, "1989", "GIIGNL",
       "DISR REQ lists NWS at 16.9 Mtpa (all five trains). Woodside H1 2026: Karratha Gas Plant has four onshore LNG trains; GIIGNL 2026 (number used): Train 2 permanently ceased, total 17 -> 14.5 mtpa."),
    ex("Australia", "North West Shelf LNG Train 2 (permanently ceased)", "Woodside", 2.6, "1989", "GIIGNL",
       "GIIGNL 2026: Woodside definitively ceased operations at Train 2 (2.6 mtpa) as feedgas declined; month not stated (placed mid-2025).", status="CLOSED", end="2025-07"),
    pj("Australia", "Pluto Train 2 (Scarborough Energy Project)", "Woodside", None, "UNDER_CONSTRUCTION", "2026-Q4", "WDS_H1",
       "Woodside H1 2026: 98% complete, first LNG cargo on track for Q4 2026. NO nameplate capacity in any opened document (Woodside, DISR REQ, GIIGNL): left blank, not in the totals ('Tried')."),
    # ---- Russia (sanctioned) ---------------------------------------------------------------------------------------
    ex("Russia", "Yamal LNG, Trains 1-4", "Novatek (Yamal LNG)", 17.4, "2017", "GIIGNL", "T1-T3 5.5 each + T4 0.9 (GIIGNL plants table)."),
    ex("Russia", "Sakhalin-2 (2 trains)", "Sakhalin Energy", 10.8, "2009", "GIIGNL", "2 x 5.4 (GIIGNL)."),
    ex("Russia", "Portovaya LNG", "Gazprom / Portovaya", 1.5, "2022", "GIIGNL",
       "US sanctions Jan 2025; GIIGNL: exports from both Baltic terminals ceased after the 27 Feb 2025 grace period; DISR REQ Sep 2026: Portovaya and Arctic LNG 2 resumed regular loadings in mid-2026, to China. Sanctioned by European buyers."),
    ex("Russia", "Vysotsk LNG", "Novatek", 0.9, "2019", "GIIGNL", "US-sanctioned Jan 2025; exports ceased after Feb 2025 (GIIGNL)."),
    ex("Russia", "Arctic LNG 2, Train 1", "Novatek", 6.6, "2024", "NOVATEK_FID",
       "6.6 mtpa per train (Novatek FID release); start 2024 from GIIGNL. Under US, EU and UK sanctions (2024); DISR REQ Sep 2026: regular cargo loadings to China since mid-2026."),
    ex("Russia", "Arctic LNG 2, Train 2", "Novatek", 6.6, "2025", "NOVATEK_FID",
       "6.6 mtpa per train (Novatek FID release); GIIGNL 2026: Train 2 started production in May 2025, exports limited by sanctions and ice-class carriers. Sanctioned.", status="COMMISSIONED_RECENT"),
    pj("Russia", "Arctic LNG 2, Train 3", "Novatek", 6.6, "PLANNED", None, "NOVATEK_FID",
       "Original schedule (Novatek 2019 FID release) was 2026; GIIGNL 2026: Novatek's other planned projects remain delayed. No current Train 3 date in an opened document. Sanctioned."),
    # ---- Africa ---------------------------------------------------------------------------------------------------
    ex("Nigeria", "Nigeria LNG, Trains 1-6", "NLNG (NNPC 49%, Shell, TotalEnergies, Eni)", 22.0, "1999", "NLNG_FF",
       "Operator: six trains, 22 MTPA. DISR REQ Sep 2026: output recovering from the 2022 force majeure."),
    pj("Nigeria", "Nigeria LNG Train 7", "NLNG", 8.0, "UNDER_CONSTRUCTION", None, "NLNG_FF",
       "Operator: Train 7 lifts capacity from 22 to 30 mtpa (+8 DERIVED); FID Dec 2019. GIIGNL 2026: over 80% complete. No start date in an opened document.", "DERIVED"),
    ex("Algeria", "Arzew GL1Z, GL2Z, GL3Z and Skikda GL1K", "Sonatrach", 25.3, "1978", "GIIGNL", "4.7 + 7.9 + 8.2 + 4.5 (GIIGNL plants table; GIIGNL Algeria total 25.3)."),
    ex("Egypt", "Damietta LNG and Idku (Egyptian LNG) T1-T2", "SEGAS (Eni/EGAS) and ELNG (Shell/Petronas)", 12.2, "2005", "GIIGNL", "5.0 + 3.6 + 3.6 (GIIGNL)."),
    ex("Angola", "Angola LNG", "Angola LNG (Chevron, Sonangol, bp, Eni, TotalEnergies)", 5.2, "2013", "GIIGNL", "GIIGNL plants table."),
    ex("Equatorial Guinea", "EG LNG", "EG LNG (Marathon / Sonagas)", 3.7, "2007", "GIIGNL", "GIIGNL plants table."),
    ex("Mozambique", "Coral South FLNG", "Eni (Area 4)", 3.4, "2022", "GIIGNL", "GIIGNL plants table; operating since 2022."),
    ex("Cameroon", "Kribi FLNG Hilli Episeyo (contract ended, relocating to Argentina)", "Golar LNG", 2.4, "2018", "GOLAR_Q2",
       "Golar Q2 2026: Hilli ended its eight-year Cameroon contract (156 cargoes), repositioning to Singapore for modifications before a 20-year contract in Argentina. Capacity 2.4 from GIIGNL. "
       "Closure month not stated (placed mid-2026).", status="CLOSED", end="2026-07"),
    ex("Congo", "Tango FLNG", "Eni (Congo LNG)", 0.6, "2024", "GIIGNL", "GIIGNL: Phase 1 FLNG."),
    ex("Congo", "Nguya FLNG (Congo LNG Phase 2)", "Eni", 2.4, "2026-02", "REQ",
       "Capacity: DISR REQ Sep 2026 (phase two, +2.4 Mtpa, 'nearing production'). Start Feb 2026 from GIIGNL 2026 (first cargo shipped); the two sources differ.", status="COMMISSIONED_RECENT"),
    ex("Mauritania/Senegal", "Greater Tortue Ahmeyim (GTA) Phase 1, FLNG Gimi", "bp (Kosmos, PETROSEN, SMH)", 2.3, "2025", "REQ",
       "DISR REQ Sep 2026: 2.3 Mtpa capacity, ramping up by H2 2026. Start year 2025 from GIIGNL 2026 (GIIGNL plants table: 2.4).", status="COMMISSIONED_RECENT"),
    pj("Mozambique", "Mozambique LNG (Afungi, 2 units)", "TotalEnergies", 13.12, "UNDER_CONSTRUCTION", "2029", "MOZ_SITE",
       "Project site: two liquefaction units, 13.12 MTPA. DISR REQ Sep 2026: '13.2 Mtpa from 2029' (year used). GIIGNL quotes 12.8."),
    pj("Mozambique", "Coral North FLNG (Eni, Area 4)", "Eni", 3.6, "FID", "2028", "ENI_FID",
       "Eni FID release 2 Oct 2025: production liquefaction capacity 3.6 MTPA, aiming to deliver within schedule in 2028."),
    # ---- Other Atlantic / Middle East / Asia-Pacific existing and projects (GIIGNL where no official figure was opened) -----
    ex("Norway", "Hammerfest LNG (Snohvit)", "Equinor", 4.2, "2007", "GIIGNL", "GIIGNL plants table."),
    ex("Peru", "Peru LNG", "Hunt Oil / Shell", 4.5, "2010", "GIIGNL", "GIIGNL plants table."),
    ex("Trinidad and Tobago", "Atlantic LNG Trains 2-4", "Atlantic LNG (Shell, bp, NGC)", 11.8, "2002", "GIIGNL",
       "3.3 + 3.3 + 5.2 (GIIGNL); GIIGNL: operating below nameplate on feedgas shortage; Shell Aphrodite gas from 2027."),
    ex("Trinidad and Tobago", "Atlantic LNG Train 1 (mothballed)", "Atlantic LNG", 3.0, "1999", "GIIGNL",
       "GIIGNL 2026: Train 1 shut down for insufficient feedgas ('mothballed'). Offline for the whole window.", status="SUSPENDED", offline_from="2000-01"),
    ex("Yemen", "Balhaf LNG (2 trains)", "Yemen LNG", 7.2, "2009", "GIIGNL",
       "Shut since 2015 under force majeure, no restart announced (GIIGNL plants table 2 x 3.6; its text says 6.7 mtpa).", status="SUSPENDED", offline_from="2015-01"),
    ex("Oman", "Oman LNG (Qalhat), Trains 1-2 and Qalhat LNG", "Oman LNG", 11.5, "2000", "GIIGNL", "3.9 + 3.9 + 3.7 (GIIGNL). Oman's export port is outside the Strait of Hormuz."),
    pj("Oman", "Marsa LNG (1 mtpa)", "TotalEnergies (80%) / OQ", 1.0, "UNDER_CONSTRUCTION", "2028", "GIIGNL", "GIIGNL 2026: FID April 2024, production expected 2028."),
    ex("United Arab Emirates", "Das Island LNG (3 trains)", "ADNOC LNG", 5.8, "1977", "GIIGNL", "1.3 + 1.3 + 3.2 (GIIGNL). " + HORMUZ),
    pj("United Arab Emirates", "Ruwais LNG (2 trains)", "ADNOC Gas (ADNOC 60%)", 9.6, "UNDER_CONSTRUCTION", "2028", "GIIGNL",
       "GIIGNL 2026: FID June 2024, 9.6 MTPA, start-up expected 2028. " + HORMUZ),
    ex("Malaysia", "PETRONAS LNG complex (Bintulu, MLNG T1-T9) and PFLNG Satu / Dua", "PETRONAS", 32.1, "1983", "GIIGNL",
       "MLNG 29.4 (2.8x3 + 3.2x3 + 3.9x2 + 3.6) + PFLNG Satu 1.2 + Dua 1.5 (GIIGNL; its Malaysia total 32.0)."),
    ex("Brunei", "Brunei LNG (5 trains)", "Brunei LNG", 7.0, "1973", "GIIGNL", "5 x 1.4 (GIIGNL)."),
    ex("Indonesia", "Tangguh LNG, Trains 1-3", "bp (Berau)", 11.4, "2009", "GIIGNL", "3 x 3.8; Train 3 started 2023 (GIIGNL)."),
    ex("Indonesia", "Donggi-Senoro LNG", "Donggi-Senoro LNG", 2.0, "2015", "GIIGNL", "GIIGNL plants table."),
    ex("Indonesia", "Bontang LNG (Train E on standby, F-H operating)", "Pertamina", 11.5, "1977", "GIIGNL",
       "DERIVED: GIIGNL Indonesia total 24.9 less Tangguh 11.4 less Donggi-Senoro 2.0. Only Trains G and H operating on limited feedgas (GIIGNL).", "DERIVED"),
    pj("Indonesia", "Abadi LNG (Masela)", "INPEX / Pertamina / PETRONAS", 9.5, "PLANNED", None, "GIIGNL",
       "GIIGNL 2026: environmental approval granted, FEED completion expected end-2026; 9.5 MTPA design. No FID or start date."),
    ex("Papua New Guinea", "PNG LNG (2 trains)", "ExxonMobil / Santos", 8.3, "2014", "GIIGNL", "GIIGNL total 8.3 (trains 4.1 + 4.1 in its table)."),
    # ---- Argentina / Golar -----------------------------------------------------------------------------------------
    pj("Argentina", "Southern Energy (SESA) FLNG Hilli Episeyo (relocated from Cameroon)", "Southern Energy (PAE, YPF, Pampa, Harbour, Golar)", 2.45, "FID", "2027-H2", "GOLAR_Q2",
       "Golar Q2 2026: 20-year contract commencing H2 2027; capacity 2.45 mtpa from GIIGNL 2026."),
    pj("Argentina", "Southern Energy (SESA) FLNG Esperanza (MKII)", "Southern Energy", 3.5, "UNDER_CONSTRUCTION", "2028-H2", "GOLAR_Q2",
       "Golar Q2 2026: 3.5 MTPA MKII unit, conversion completes Q4 2027, contract start-up scheduled H2 2028."),
    pj("Unallocated (FLNG)", "Golar FLNG Four (MKII, no charter yet)", "Golar LNG", 3.5, "FID", "2029", "GOLAR_Q2",
       "Golar Q2 2026: EPC signed 12 Aug 2026 with CIMC Raffles, 3.5 MTPA, delivery expected by year-end 2029; no charter or site yet."),
    # ---- United States, beyond the EIA table -----------------------------------------------------------------------
    pj("United States", "Sabine Pass Stage 5 expansion (Train 7 + BOG re-liquefaction)", "Cheniere", 6.0, "PLANNED", None, "CHEN_SP",
       "Cheniere Q2 2026: expected total production capacity of 'over 6 mtpa' (6.0 = the stated floor); FERC permit expected late 2026, FID early 2027; limited notice to proceed to Bechtel."),
]

# ---------------------------------------------------------------------------------------------------------------------
# Cross-check figures (GIIGNL Annual Report 2026, press copies). NEVER in Projects: shown on the 'Cross-check' sheet next to the raw total.
# (country, item, publisher, value mtpa or None, text, compare: country name in Projects whose OPERATING total is compared or None)
# ---------------------------------------------------------------------------------------------------------------------
G26 = "GIIGNL Annual Report 2026"
CROSS = [
    ("WORLD", "Global liquefaction nominal capacity end-2025", G26, 524.0, "524 MTPA incl. 16.5 MTPA FLNG (executive summary)", "WORLD"),
    ("United States", "US liquefaction capacity 2025", G26, 110.0, "'nearly 110 MTPA' (nominal basis; EIA baseload sum is on a different basis and includes units commissioned in 2026)", "United States"),
    ("Qatar", "Existing trains", G26, 77.0, "Plants table / narrative: 77 MTPA", "Qatar"),
    ("Qatar", "Ras Laffan Trains 4 and 6 damaged 18 Mar 2026", G26, 12.8, "Per QatarEnergy's March 2026 statement as quoted by GIIGNL: repairs 3-5 years; possible force majeure up to five years. Reuters (7 Oct 2026, via BOE Report): two trains damaged, repairs ~3 years, 17% of capacity.", None),
    ("Qatar", "North Field East first train", "Reuters via BOE Report (press copy)", None, "CEO: production expected H1 2027; first train ready to start in Nov 2026, LNG from it in Q1 2027 (two unnamed sources). GIIGNL: first two NFE trains had been expected late 2026, likely postponed.", None),
    ("Australia", "Plants table total", G26, 85.8, "GIIGNL Australia total (DISR REQ nameplate map sums to 88)", "Australia"),
    ("Australia", "North West Shelf after Train 2 closure", G26, 14.5, "Train 2 (2.6 MTPA) permanently ceased; total from 17 to 14.5 MTPA", None),
    ("Australia", "Pluto Train 2 (Scarborough)", "Woodside (dates) / none (capacity)", None, "First LNG cargo Q4 2026 (Woodside H1 2026); no Mtpa figure in an opened raw source", None),
    ("Canada", "LNG Canada Trains 1-2 + Tilbury", G26, 14.3, "7.0 + 7.0 online in 2025, Tilbury 0.3", "Canada"),
    ("Canada", "Cedar LNG (floating)", G26, None, "Start-up targeted 2028 (Pembina: late 2028); no mtpa in the opened documents", None),
    ("Canada", "Ksi Lisims LNG (floating)", G26, 12.0, "12 MTPA FLNG, federal approvals, fast-tracked; no FID", None),
    ("Mexico", "Energia Costa Azul Phase 1 + Altamira FLNG 1", G26, 4.65, "3.25 + 1.4. TotalEnergies 2Q26 release: ECA started up early July 2026 (no capacity)", None),
    ("Russia", "Russia plants (Yamal 17.4, Sakhalin-2 10.8, Portovaya 1.5, Vysotsk 0.9, Arctic LNG 2 T1 6.6)", G26, 37.2, "GIIGNL Russia total 26.4 excl. Sakhalin-2; Arctic LNG 2 Train 2 started May 2025", "Russia"),
    ("Nigeria", "NLNG six trains", G26, 22.2, "3.3x3 + 4.1x3 (operator says 22)", "Nigeria"),
    ("Algeria", "Arzew GL1Z/GL2Z/GL3Z, Skikda", G26, 25.3, "Plants table", None),
    ("Egypt", "Damietta + Idku", G26, 12.2, "5.0 + 3.6 + 3.6", None),
    ("Angola", "Angola LNG", G26, 5.2, "", None),
    ("Equatorial Guinea", "EG LNG", G26, 3.7, "", None),
    ("Mozambique", "Coral South FLNG", G26, 3.4, "Operating since 2022", None),
    ("Mozambique", "Mozambique LNG", G26, 12.8, "Narrative figure; operator site says 13.12", None),
    ("Cameroon", "Hilli Episeyo FLNG (Kribi)", G26, 2.4, "Golar Q2 2026: contract ended, relocating to Argentina (Hilli + Gimi together 5.1 MTPA)", None),
    ("Argentina", "Hilli Episeyo (SESA)", G26, 2.45, "Golar Q2 2026: 20-year contract commencing H2 2027", None),
    ("Congo", "Tango FLNG 0.6 + Nguya 2.4", G26, 3.0, "Nguya first cargo Feb 2026 per GIIGNL", None),
    ("Mauritania/Senegal", "GTA Phase 1 (Gimi)", G26, 2.4, "plants table; narrative 'around 2.3'", "Mauritania/Senegal"),
    ("Norway", "Hammerfest LNG", G26, 4.2, "", None),
    ("Peru", "Peru LNG", G26, 4.5, "", None),
    ("Trinidad and Tobago", "Atlantic LNG T1-T4 (T1 mothballed)", G26, 14.8, "", None),
    ("Yemen", "Balhaf LNG (shut since 2015)", G26, 7.2, "table 2 x 3.6; narrative 6.7", None),
    ("Oman", "Oman LNG + Qalhat", G26, 11.5, "", None),
    ("Oman", "Marsa LNG", G26, 1.0, "FID Apr 2024, 2028", None),
    ("United Arab Emirates", "Das Island", G26, 5.8, "", None),
    ("United Arab Emirates", "Ruwais LNG", G26, 9.6, "FID June 2024, start-up 2028", None),
    ("Malaysia", "PETRONAS LNG complex + PFLNG", G26, 32.1, "", None),
    ("Brunei", "Brunei LNG", G26, 7.0, "", None),
    ("Indonesia", "Bontang, Tangguh, Donggi-Senoro", G26, 24.9, "Abadi LNG 9.5 MTPA planned (FEED end-2026)", None),
    ("Papua New Guinea", "PNG LNG", G26, 8.3, "", None),
]

# ---------------------------------------------------------------------------------------------------------------------
# Tried and not usable
# ---------------------------------------------------------------------------------------------------------------------
TRIED = [
    # ---- sources tried -----------------------------------------------------------------------------------------
    ("SOURCE: FERC 'North American LNG Export Terminals - Existing, Approved not yet built, Proposed'", "https://www.ferc.gov/media/north-american-lng-export-terminals-existing-approved-not-yet-built-and-proposed",
     "Page reachable (12 Oct 2026) but the list is a linked document the HTML does not expose; FERC's earlier LNG pages list no start dates (TEXAS_STEO_FERC_PROBE). EIA's liquefaction workbook (cites the FERC and DOE dockets) is the US source."),
    ("SOURCE: Canada Energy Regulator LNG export licences", "https://www.cer-rec.gc.ca/en/data-analysis/energy-commodities/natural-gas/report/lng-export-licences/index.html",
     "404 after the CER site relaunch; replacement page not found. Licence volumes are not plant capacity in any case; operator releases are used for Canada."),
    ("SOURCE: LNG Canada site (Phase 1 / Phase 2 pages)", "https://www.lngcanada.ca/", "HTTP 202 (no content) from Actions on every inner page; Shell's release text (as republished) used."),
    ("SOURCE: Shell release on shell.com", "https://www.shell.com/news-and-insights/newsroom/news-and-media-releases/2026/shell-takes-final-investment-decision-to-double-lng-canada-capacity.html",
     "Script-rendered page; the 14 -> 28 mtpa text is not in the HTML. The same release text was read on Hellenic Shipping News and is labelled so."),
    ("SOURCE: DISR REQ Excel data (forecast / historical)", "https://www.industry.gov.au/sites/default/files/2026-10/resources-and-energy-quarterly-september-2026-forecast-data.xlsx",
     "Opened: Australian export volumes and prices only, no project capacity table. The PDF nameplate map is used."),
    ("SOURCE: QatarEnergy statements on the Ras Laffan strikes (March 2026) and on per-train start dates", "https://www.qatarenergy.qa/en/MediaCenter/Pages/News.aspx",
     "News list is script-rendered (link templates only); the March statement, a damaged-train figure and a dated NFE/NFS/NFW schedule are NOT in any QatarEnergy document opened (deck Nov 2025, Annual Review 2024). The 12.8 mtpa / three-to-five-year figures are GIIGNL / Reuters: 'Cross-check' only."),
    ("SOURCE: IGU World LNG Report 2026", "https://www.igu.org/resources/world-lng-report-2025/", "404 on the guessed address; no direct PDF link found."),
    ("SOURCE: TotalEnergies / Mozambique LNG restart release (businesswire)", "https://www.businesswire.com/news/home/20260129402488/en/", "403 from Actions; project site and DISR REQ used."),
    ("SOURCE: Sempra Q2 2026 release; Kosmos GTA release; SEC exhibits", "https://www.sempra.com/newsroom/press-releases/sempra-reports-strong-second-quarter-2026-results", "403 from Actions."),
    ("SOURCE: GIIGNL Excel tables / LNG Trade Matrix / World LNG Map", "https://giignl.org/resources/", "Members only; the free annual report PDF is used on 'Cross-check' only."),
    # ---- not sourced: left blank / absent, not filled ----------------------------------------------------------------
    ("NOT SOURCED: Qatar damaged-train size and return date from QatarEnergy itself", "https://www.qatarenergy.qa/en/MediaCenter/Pages/News.aspx", "March 2026 statement not reachable; 12.8 mtpa / 3-5 years are GIIGNL (labelled in Projects)."),
    ("NOT SOURCED: Qatar North Field East Trains 2-4, South and West start dates", "https://www.qatarenergy.qa/en/MediaCenter/Publications/QatarEnergy_Investors_Presentation_November_2025.pdf", "Only 'by 2030' for the programme; trains are undated (NFE Train 1: CEO H1 2027 via Reuters copy)."),
    ("NOT SOURCED: Pluto Train 2 (Scarborough) capacity", "https://www.woodside.com/investors/reports-investor-briefings", "Woodside Q2/H1 2026 and FY2025 briefing opened: first LNG cargo Q4 2026, no Mtpa for Train 2; DISR REQ and GIIGNL have none. Row kept with blank capacity (not in totals)."),
    ("NOT SOURCED: Cedar LNG capacity", "https://www.pembina.com/media-centre/news/details/fa13b9e4-bbc0-44d0-9ab6-88c72cbd54e5", "Pembina Q2 2026 gives late 2028, no mtpa; Cedar site and GIIGNL none. Row kept with blank capacity."),
    ("NOT SOURCED: Hilli-only capacity from Golar", "https://mfn.se/one/a/golar-lng/golar-lng-limited-interim-results-for-the-period-ended-june-30-2026-b168e773", "Golar gives 'Hilli + Gimi = 5.1 MTPA'; Hilli's own 2.4 / 2.45 mtpa is GIIGNL (labelled)."),
    ("NOT SOURCED: North West Shelf Train 2 closure month; LNG Canada / GTA / Arctic LNG 2 start months", "https://www.industry.gov.au/publications/resources-and-energy-quarterly-september-2026", "Only years (GIIGNL) or none; placed as stated on 'Placement'."),
    ("NOT SOURCED: Papua LNG, Rovuma LNG, Tanzania LNG, GTA Phase 2, Yakaar-Teranga, Mexico Pacific Saguaro, Argentina LNG (YPF/Eni), Murmansk / Ob LNG, other US proposals", "(none)",
     "No opened document with a capacity and status (Santos Q2 2026 has no Papua LNG). Search snippets are never used. Add when a release is opened."),
    ("NOT SOURCED: Nigeria LNG Train 7 start date", "https://www.nlng.com/operations/train-7", "Official page text empty (script-rendered); Facts and Figures 2024 gives capacity only."),
]

# country -> chart group
GROUPS = {"United States": "United States", "Qatar": "Qatar", "Canada": "Canada", "Mexico": "Mexico", "Australia": "Australia",
          "Russia": "Russia", "Nigeria": "Africa", "Algeria": "Africa", "Egypt": "Africa", "Angola": "Africa",
          "Equatorial Guinea": "Africa", "Mozambique": "Africa", "Cameroon": "Africa", "Congo": "Africa",
          "Mauritania/Senegal": "Africa"}
GROUP_ORDER = ["United States", "Qatar", "Australia", "Canada", "Mexico", "Russia", "Africa", "Other"]
GROUP_COLOUR = {"United States": "2A78D6", "Qatar": "8E2C57", "Australia": "EDA100", "Canada": "1BAF7A", "Mexico": "008300",
                "Russia": "6E4B2A", "Africa": "EB6834", "Other": "9A9A9A"}


def group_of(country):
    return GROUPS.get(country, "Other")


# ---------------------------------------------------------------------------------------------------------------------
def num(x):
    if x is None:
        return np.nan
    if isinstance(x, (int, float, np.number)):
        return float(x)
    m = re.search(r"[-+]?\d*\.?\d+", str(x).replace(",", ""))
    return float(m.group()) if m else np.nan


def sess():
    s = requests.Session()
    s.headers.update(UA)
    return s


def when_to_exp(when):
    """(expected_online text, basis, placement Timestamp or None)."""
    if when is None or str(when) in ("None", "nan", ""):
        return "", "NOT_STATED", None
    w = str(when)
    m = re.fullmatch(r"(\d{4})-(\d{2})", w)
    if m:
        return w, "SOURCED_MONTH", pd.Timestamp(int(m[1]), int(m[2]), 1)
    m = re.fullmatch(r"(\d{4})-Q([1-4])", w)
    if m:
        return m[1], "SOURCED_QUARTER", pd.Timestamp(int(m[1]), 3 * int(m[2]) - 1, 1)
    m = re.fullmatch(r"(\d{4})-H([12])", w)
    if m:
        return m[1], "SOURCED_YEAR", pd.Timestamp(int(m[1]), 4 if m[2] == "1" else 10, 1)
    m = re.fullmatch(r"(\d{4})", w)
    if m:
        return w, "SOURCED_YEAR", pd.Timestamp(int(w), 7, 1)
    return "", "NOT_STATED", None


def placement_rule(when):
    w = str(when)
    if re.fullmatch(r"\d{4}-\d{2}", w):
        return "month as published"
    if "-Q" in w:
        return "middle month of the quarter"
    if "-H1" in w:
        return "April (middle of H1)"
    if "-H2" in w:
        return "October (middle of H2)"
    if re.fullmatch(r"\d{4}", w):
        return "July (middle of the year)"
    return "no date in an opened source - not in the dated series"


# ---------------------------------------------------------------------------------------------------------------------
# EIA workbook
# ---------------------------------------------------------------------------------------------------------------------
def eia_candidates():
    y, q = TODAY.year, (TODAY.month - 1) // 3 + 1
    out = []
    for _ in range(6):
        out.append((y, q, EIA_BASE.format(y=y, q=q)))
        q -= 1
        if q == 0:
            y, q = y - 1, 4
    return out


def parse_eia(content):
    wb = load_workbook(io.BytesIO(content), data_only=True)
    release = None
    ws0 = wb.worksheets[0]
    for row in ws0.iter_rows(values_only=True):
        if row and any(str(c).startswith("Release Date") for c in row if c):
            release = next((c for c in row if isinstance(c, (dt.datetime, dt.date))), None)
    rows = []
    for ws in wb.worksheets:
        t = ws.title.lower()
        if t.startswith("existing"):
            kind = "UC"
        elif t.startswith("approved"):
            kind = "APPROVED"
        else:
            continue
        start = None
        for r in ws.iter_rows(values_only=True):
            r = list(r)
            if r and r[0] and str(r[0]).strip().startswith("Project name"):
                start = True
                continue
            if not start or not r or not r[0]:
                continue
            name = str(r[0]).strip()
            if name.startswith(("Notes", "^", "+", "Note ", "1 ", "This list", "Project status")) or len(name) < 3:
                continue
            if kind == "UC":
                if r[1] is None:
                    continue
                rest = [c for c in r[6:] if c not in (None, "")]
                status_raw = str(rest[0]) if rest else ""
                insvc = rest[1] if len(rest) > 1 else None
                rows.append(dict(sheet="UC", project=name, train=str(r[1]).strip(), base_bcfd=num(r[2]), base_mtpa=num(r[3]),
                                 peak_bcfd=num(r[4]), peak_mtpa=num(r[5]), status_raw=status_raw,
                                 in_service_raw=insvc.strftime("%Y-%m-%d") if isinstance(insvc, (dt.datetime, dt.date)) else ("" if insvc is None else str(insvc)),
                                 operator=str(r[-1]) if r[-1] else (str(r[-2]) if len(r) > 1 and r[-2] else ""), detail=""))
            else:
                if len(r) < 8 or r[4] in (None, "") and r[6] in (None, ""):
                    continue
                rows.append(dict(sheet="APPROVED", project=name, train=str(r[4]), base_bcfd=num(r[2]), base_mtpa=num(r[3]),
                                 peak_bcfd=num(r[5]), peak_mtpa=num(r[6]), status_raw="Approved", in_service_raw="",
                                 operator=str(r[1]) if r[1] else "", detail=str(r[7]) if r[7] else ""))
    df = pd.DataFrame(rows)
    # the operator column of the train table is the last non-empty cell
    return df, release


def fetch_eia(prev_rel, prev_eia, full):
    """Return (eia_rows DataFrame, release_row dict). Conditional download; reuse the stored parse when unchanged."""
    s = sess()
    hit = None
    for y, q, url in eia_candidates():
        hdr = {}
        if not full and prev_rel is not None and url in set(prev_rel["url"]):
            p = prev_rel[prev_rel["url"] == url].iloc[0]
            if isinstance(p.get("etag"), str) and p["etag"]:
                hdr["If-None-Match"] = p["etag"]
            elif isinstance(p.get("last_modified"), str) and p["last_modified"]:
                hdr["If-Modified-Since"] = p["last_modified"]
        try:
            r = s.get(url, headers=hdr, timeout=90)
        except Exception as e:  # noqa: BLE001
            print("  EIA fetch failed", url, type(e).__name__)
            continue
        print("  EIA", url.split("/")[-1], r.status_code)
        if r.status_code == 304 and prev_eia is not None:
            p = prev_rel[prev_rel["url"] == url].iloc[0].to_dict()
            p["checked"] = str(TODAY)
            p["result"] = "UNCHANGED (304) - stored parse reused"
            return prev_eia, p
        if r.status_code == 200 and r.content[:2] == b"PK":
            hit = (url, r)
            break
    if hit is None:
        if prev_eia is not None:
            row = prev_rel[prev_rel["document"].str.startswith("EIA")].iloc[0].to_dict() if prev_rel is not None else {}
            row["checked"] = str(TODAY)
            row["result"] = "UNREACHABLE - stored parse reused"
            return prev_eia, row
        raise RuntimeError("EIA liquefaction workbook not reachable and no stored parse")
    url, r = hit
    sha = hashlib.sha256(r.content).hexdigest()
    row = dict(document="EIA U.S. liquefaction capacity workbook", url=url, etag=r.headers.get("ETag", ""),
               last_modified=r.headers.get("Last-Modified", ""), size=len(r.content), sha256=sha, checked=str(TODAY), result="")
    if (not full and prev_rel is not None and prev_eia is not None and
            (prev_rel[prev_rel["url"] == url]["sha256"] == sha).any()):
        row["result"] = "UNCHANGED (same SHA-256) - stored parse reused"
        return prev_eia, row
    df, release = parse_eia(r.content)
    df["release_date"] = release.strftime("%Y-%m-%d") if release else ""
    df["file"] = url.split("/")[-1]
    row["result"] = "NEW/CHANGED - parsed"
    row["release_date"] = df["release_date"].iloc[0] if len(df) else ""
    return df, row


def eia_status(raw):
    s = raw.lower()
    if s.startswith("commercial"):
        return "OPERATING"
    if s.startswith("commissioning"):
        return "COMMISSIONED"
    if s.startswith("under construction"):
        return "UNDER_CONSTRUCTION"
    return "PLANNED"


def eia_when(raw):
    """EIA in-service text -> when string."""
    raw = str(raw).strip()
    m = re.fullmatch(r"(\d{4})-(\d{2})-\d{2}", raw)
    if m:
        return f"{m[1]}-{m[2]}"
    m = re.fullmatch(r"([12])H(\d{4})", raw)
    if m:
        return f"{m[2]}-H{m[1]}"
    m = re.fullmatch(r"(\d{4})", raw)
    if m:
        return raw
    return None


CC3_TRAINS = [  # CCL Stage 3 midscale trains (Cheniere Q2 2026 release): (label, status, when, note)
    ("Midscale Train 1", "COMMISSIONED_RECENT", "2025", "Cheniere Q2 2026: substantial completion of Trains 1-4 in 2025 (month per train not stated); EIA: first LNG Dec 2024 / in service Feb 2025"),
    ("Midscale Train 2", "COMMISSIONED_RECENT", "2025", "as Train 1"),
    ("Midscale Train 3", "COMMISSIONED_RECENT", "2025", "as Train 1"),
    ("Midscale Train 4", "COMMISSIONED_RECENT", "2025", "as Train 1"),
    ("Midscale Train 5", "COMMISSIONED_RECENT", "2026-03", "Cheniere Q2 2026: substantial completion March 2026"),
    ("Midscale Train 6", "COMMISSIONED_RECENT", "2026-06", "Cheniere Q2 2026: substantial completion June 2026"),
    ("Midscale Train 7", "UNDER_CONSTRUCTION", "2026-Q4", "Cheniere Q2 2026 (6 Aug): first LNG expected imminently, substantial completion expected in the fall of 2026"),
]


def eia_to_rows(eia):
    rel = eia["release_date"].iloc[0] if len(eia) else ""
    url = EIA_BASE.format(y=2026, q=3)
    if len(eia) and "file" in eia:
        url = "https://www.eia.gov/naturalgas/importsexports/liquefactioncapacity/" + str(eia["file"].iloc[0])
    pub = "U.S. Energy Information Administration"
    doc = f"U.S. Liquefaction Capacity workbook ({eia['file'].iloc[0] if len(eia) else ''}; release {rel})"
    out = []
    for _, r in eia[eia["sheet"] == "UC"].iterrows():
        st = eia_status(r["status_raw"])
        when = eia_when(r["in_service_raw"])
        base = r["base_mtpa"]
        peak = r["peak_mtpa"]
        note = f"EIA status '{r['status_raw']}'; in-service '{r['in_service_raw']}'. Baseload nameplate {base:.2f} mtpa; EIA peak nameplate {peak:.2f} mtpa (EIA's own Mtpa where marked ^ else its Bcf/d conversion)."
        if r["project"].startswith("Corpus Christi Liquefaction Stage 3"):
            per = round(base / 7, 3)
            for lab, st3, w3, n3 in CC3_TRAINS:
                out.append(dict(country="United States", project=f"Corpus Christi Stage 3 {lab}", operator="Cheniere Energy", cap=per,
                                status=st3, when=w3, src="EIA", extra_src="CHEN_Q2", start_known=True, notes=f"{n3}. Stage 3 = 7 midscale trains, {base:.2f} mtpa baseload in EIA; per train {per} mtpa DERIVED (EIA note: 1.43 mtpa per train). "
                                f"Cheniere Q2 2026: Sabine Pass + Corpus Christi ~55 mtpa in operation. EIA peak for the stage {peak:.2f} mtpa.",
                                evidence="DERIVED", end=None, offline_from=None, kind="existing" if st3 != "UNDER_CONSTRUCTION" else "project",
                                eia=(doc, url, rel)))
            continue
        if st in ("OPERATING", "COMMISSIONED"):
            ym = when_to_exp(when)[2]
            st_f = "COMMISSIONED_RECENT" if (ym is not None and ym >= pd.Timestamp("2025-01-01")) else "OPERATING"
            out.append(dict(country="United States", project=f"{r['project']} {r['train']}", operator=r["operator"], cap=base, status=st_f,
                            when=when, src="EIA", notes=note, evidence="SOURCED", end=None, offline_from=None, kind="existing", start_known=True, eia=(doc, url, rel)))
        else:
            out.append(dict(country="United States", project=f"{r['project']} {r['train']}", operator=r["operator"], cap=base,
                            status="UNDER_CONSTRUCTION" if st == "UNDER_CONSTRUCTION" else "PLANNED", when=when, src="EIA",
                            notes=note + (" 2H/1H dates are half-years." if when and "-H" in when else ""), evidence="SOURCED", end=None,
                            offline_from=None, kind="project", start_known=True, eia=(doc, url, rel)))
    for _, r in eia[eia["sheet"] == "APPROVED"].iterrows():
        name = r["project"]
        if "additional DOE application" in name:
            continue     # EIA Note A: the two Lake Charles applications are not additive
        cap = r["peak_mtpa"] if not np.isnan(r["peak_mtpa"]) else r["base_mtpa"]
        detail = str(r["detail"])
        status = "PERMITTED"
        n = f"EIA 'Approved' tab (DOE and FERC/MARAD approved, no FID): {detail[:240]}. Design capacity {cap:.2f} mtpa ({int(num(r['train'])) if not np.isnan(num(r['train'])) else '?'} train(s))."
        ev = "SOURCED"
        if name.startswith("Lake Charles"):
            status = "SUSPENDED"
            n += " Energy Transfer suspended development on 18 Dec 2025 (EIA note D)."
        if name.startswith("Delfin"):
            cap = round(cap - 4.4, 3)
            n += " Total approved 8.8 mtpa less the first vessel (4.4 mtpa, under construction in the other tab): remainder DERIVED."
            ev = "DERIVED"
            name = "Delfin FLNG additional vessels (beyond Train 1)"
        if name.startswith("Alaska"):
            status = "PLANNED"
        out.append(dict(country="United States", project=name, operator=r["operator"], cap=cap, status=status, when=None, src="EIA", notes=n,
                        evidence=ev, end=None, offline_from=None, kind="project", start_known=True, eia=(doc, url, rel)))
    return out


# ---------------------------------------------------------------------------------------------------------------------
def build_projects(eia, prev_projects):
    rows = []
    for r in eia_to_rows(eia) + CURATED:
        exp, basis, plc = when_to_exp(r["when"])
        status = r["status"]
        startunk = (r["kind"] == "existing" and not r.get("start_known", True))
        if startunk:
            plc = NOW_M          # start not in a raw source: operating at the run date, earlier history not sourced
        # an existing plant: expected_online = start (year/month) as published
        elif r["kind"] == "existing" and r["status"] in ("OPERATING", "CLOSED", "SUSPENDED") and basis == "SOURCED_YEAR" and r["when"] and int(str(r["when"])[:4]) < 2021:
            plc = pd.Timestamp(int(str(r["when"])[:4]), 7, 1)
        if "eia" in r:
            pub, doc, url, asof = "U.S. Energy Information Administration", r["eia"][0], r["eia"][1], r["eia"][2]
            if r.get("extra_src"):
                x = SRC[r["extra_src"]]
                pub, doc, url = pub + "; " + x[0], doc + "; " + x[1], url + " ; " + x[2]
        else:
            pub, doc, url, asof = (SRC.get(r["src"]) or XSRC[r["src"]])
        cap = r["cap"]
        rows.append(dict(
            region="Global supply", country=r["country"], asset_type="liquefaction_train", project=r["project"], operator=r["operator"],
            fuel_or_type="liquefaction", capacity_value=cap if cap is not None else np.nan, capacity_unit="mtpa", capacity_mw=np.nan,
            capacity_bcfd=round(cap * MTPA_TO_BCFD, 4) if cap is not None else np.nan, status=status, expected_online=exp,
            online_date_basis=basis, source_publisher=pub, source_document=doc, source_url=url, as_of_date=asof,
            evidence=r["evidence"], notes=(r["notes"] + (" [capacity_bcfd DERIVED: mtpa x 1.36 bcm/y x 35.3147 / 365.25 = 0.1315 Bcf/d per mtpa]")),
            _when=r["when"], _startunk=startunk, _kind=r["kind"], _end=r["end"], _offline_from=r["offline_from"], _plc=plc, _group=group_of(r["country"])))
    df = pd.DataFrame(rows)
    # first_seen / last_seen and history from the committed workbook
    today = str(TODAY)
    key = df["region"] + "|" + df["country"] + "|" + df["project"]
    df["_key"] = key
    hist = []
    first, last = {}, {}
    if prev_projects is not None and len(prev_projects):
        pk = prev_projects["region"] + "|" + prev_projects["country"] + "|" + prev_projects["project"]
        prev = prev_projects.assign(_key=pk).set_index("_key")
        for k in key:
            if k in prev.index:
                p = prev.loc[k]
                first[k] = str(p["first_seen"])
                last[k] = today
                for fld in ("status", "expected_online", "capacity_value"):
                    new = df.loc[df["_key"] == k, fld].iloc[0]
                    old = p[fld]
                    if (pd.isna(old) and pd.isna(new)) or str(old) == str(new) or (not pd.isna(old) and not pd.isna(new) and fld == "capacity_value" and abs(float(old) - float(new)) < 1e-9):
                        continue
                    hist.append(dict(key=k, field=fld, old=old, new=new, date_seen=today))
        gone = [k for k in prev.index if k not in set(key)]
        for k in gone:
            hist.append(dict(key=k, field="dropped from sources", old=prev.loc[k, "status"], new="(not in this run)", date_seen=today))
    df["first_seen"] = [first.get(k, today) for k in key]
    df["last_seen"] = [last.get(k, today) for k in key]
    # rows dropped from the sources stay visible
    if prev_projects is not None and len(prev_projects):
        pk = prev_projects["region"] + "|" + prev_projects["country"] + "|" + prev_projects["project"]
        keep = prev_projects[~pk.isin(set(key))].copy()
        if len(keep):
            keep["notes"] = keep["notes"].astype(str) + " [NOT IN LATEST RUN: kept from earlier run, see History]"
            for c in ("_when", "_startunk", "_kind", "_end", "_offline_from", "_plc", "_group", "_key"):
                keep[c] = None
            df = pd.concat([df, keep.reindex(columns=df.columns)], ignore_index=True)
    return df, pd.DataFrame(hist, columns=["key", "field", "old", "new", "date_seen"])


# ---------------------------------------------------------------------------------------------------------------------
def sv(v):
    return v if isinstance(v, str) and v else None


def months():
    return pd.date_range(FIRST_M, LAST_M, freq="MS")


def layer_of(p):
    """Layer for the dated/undated monthly logic."""
    st = p["status"]
    if st in ("OPERATING", "COMMISSIONED_RECENT", "CLOSED"):
        return "OP"
    if st == "SUSPENDED" and p["_kind"] == "existing":
        return "OFF"
    if st == "UNDER_CONSTRUCTION":
        return "UC"
    if st == "FID":
        return "FID"
    if st in ("PERMITTED", "PLANNED", "QUEUE", "TARGET"):
        return "PL"
    return "MEMO"      # suspended / cancelled projects


def effective_start(p):
    """Start month used in the series; future projects whose source date is already past start in the run month."""
    plc = p["_plc"]
    if plc is None or pd.isna(plc):
        return None
    if p["_kind"] == "project" and plc < NOW_M:
        return NOW_M
    return plc


def capacity_series(df):
    idx = months()
    cols = {k: np.zeros(len(idx)) for k in ("OP", "OFF", "UC", "FID", "PL")}
    by_group = {(g, lay): np.zeros(len(idx)) for g in GROUP_ORDER for lay in ("OP", "OFF", "UC", "FID", "PL")}
    for _, p in df.iterrows():
        cap = p["capacity_value"]
        if pd.isna(cap) or p["_kind"] is None:
            continue
        lay = layer_of(p)
        if lay == "MEMO":
            continue
        g = p["_group"]
        if lay in ("UC", "FID", "PL"):
            s = effective_start(p)
            if s is None:
                continue
            mask = idx >= s
            cols[lay][mask] += cap
            by_group[(g, lay)][mask] += cap
        elif lay == "OP":
            s = p["_plc"] if p["_plc"] is not None and not pd.isna(p["_plc"]) else pd.Timestamp("2000-01-01")
            e = pd.Timestamp(sv(p["_end"]) + "-01") if sv(p["_end"]) else None
            mask = (idx >= s) & ((idx < e) if e is not None else True)
            cols["OP"][mask] += cap
            by_group[(g, "OP")][mask] += cap
        elif lay == "OFF":
            of = pd.Timestamp(sv(p["_offline_from"]) + "-01") if sv(p["_offline_from"]) else pd.Timestamp("2000-01-01")
            s = p["_plc"] if p["_plc"] is not None and not pd.isna(p["_plc"]) else pd.Timestamp("2000-01-01")
            on = (idx >= s) & (idx < of)
            off = idx >= max(of, s)
            cols["OP"][on] += cap
            by_group[(g, "OP")][on] += cap
            cols["OFF"][off] += cap
            by_group[(g, "OFF")][off] += cap
    return idx, cols, by_group


def build_capacity_table(df):
    idx, c, bg = capacity_series(df)
    und = df[df["_kind"].notna()].copy()
    ud = {k: float(und[(und["capacity_value"].notna()) & (und["_plc"].isna()) & (und.apply(layer_of, axis=1) == k)]["capacity_value"].sum()) for k in ("UC", "FID", "PL")}
    t = pd.DataFrame(index=idx)
    t.index.name = "Month"
    t["Operating (mtpa)"] = c["OP"]
    t["Offline: damaged / shut-in (mtpa)"] = c["OFF"]
    t["Under construction, by expected month (cum.)"] = c["UC"]
    t["FID, by expected month (cum.)"] = c["FID"]
    t["Total sourced nameplate = operating + UC + FID"] = c["OP"] + c["UC"] + c["FID"]
    t["Planned / permitted with a date (cum.)"] = c["PL"]
    t["Total incl. dated planned"] = t["Total sourced nameplate = operating + UC + FID"] + c["PL"]
    t["Undated under construction (memo, constant)"] = ud["UC"]
    t["Undated FID (memo, constant)"] = ud["FID"]
    t["Undated planned / permitted (memo, constant)"] = ud["PL"]
    t["Total sourced nameplate (Bcf/d, DERIVED)"] = (c["OP"] + c["UC"] + c["FID"]) * MTPA_TO_BCFD
    return t, bg


def chart_tables(bg, idx):
    cn = pd.DataFrame(index=idx)
    cn.index.name = "Month"
    lab = {"OP": "operating", "UC": "under construction + FID (dated)", "PL": "planned / permitted (dated)"}
    for g in GROUP_ORDER:
        cn[f"{g} - operating"] = bg[(g, "OP")]
    for g in GROUP_ORDER:
        cn[f"{g} - UC + FID"] = bg[(g, "UC")] + bg[(g, "FID")]
    for g in GROUP_ORDER:
        cn[f"{g} - planned (dated)"] = bg[(g, "PL")]
    cn["Offline (damaged / shut-in)"] = sum(bg[(g, "OFF")] for g in GROUP_ORDER)
    return cn


def additions_by_year(df):
    years = list(range(2021, 2031))
    add = pd.DataFrame(0.0, index=years, columns=GROUP_ORDER)
    clo = pd.DataFrame(0.0, index=years, columns=GROUP_ORDER)
    for _, p in df.iterrows():
        cap = p["capacity_value"]
        if pd.isna(cap) or p["_kind"] is None:
            continue
        lay = layer_of(p)
        if lay in ("UC", "FID", "PL"):
            s = effective_start(p)
        elif lay == "OP" and not p["_startunk"] and (p["status"] == "COMMISSIONED_RECENT" or (p["_kind"] == "existing" and p["_plc"] is not None and not pd.isna(p["_plc"]) and p["_plc"] >= FIRST_M)):
            s = p["_plc"]
        else:
            s = None
        if s is not None and not pd.isna(s) and s.year in add.index:
            add.loc[s.year, p["_group"]] += cap
        if p["status"] == "CLOSED" and sv(p["_end"]):
            y = int(p["_end"][:4])
            if y in clo.index:
                clo.loc[y, p["_group"]] += cap
        if lay == "OFF" and sv(p["_offline_from"]) and int(p["_offline_from"][:4]) in clo.index and int(p["_offline_from"][:4]) >= 2021:
            clo.loc[int(p["_offline_from"][:4]), p["_group"]] += cap
    return add, clo


def by_country(df):
    d = df[df["_kind"].notna()].copy()
    d["lay"] = d.apply(layer_of, axis=1)
    d["cap"] = d["capacity_value"]
    end30 = LAST_M

    def agg(g):
        o = {}
        op = g[(g.lay == "OP") & (g.status != "CLOSED")]
        o["Operating now (mtpa)"] = op[op["_plc"].isna() | (op["_plc"] <= NOW_M)]["cap"].sum()
        o["Offline: damaged / shut-in (mtpa)"] = g[g.lay == "OFF"]["cap"].sum()
        for lay, nm in (("UC", "Under construction"), ("FID", "FID")):
            sub = g[g.lay == lay]
            dated = sub[sub["_plc"].notna()]
            o[f"{nm}: dated, online by Dec 2030 (mtpa)"] = dated.loc[np.array([effective_start(r) <= end30 for _, r in dated.iterrows()], dtype=bool), "cap"].sum()
            o[f"{nm}: undated (mtpa)"] = sub[sub["_plc"].isna()]["cap"].sum()
        sub = g[g.lay == "PL"]
        o["Planned / permitted: dated (mtpa)"] = sub[sub["_plc"].notna()]["cap"].sum()
        o["Planned / permitted: undated (mtpa)"] = sub[sub["_plc"].isna()]["cap"].sum()
        o["Suspended / cancelled projects (memo, mtpa)"] = g[(g.lay == "MEMO")]["cap"].sum()
        o["Nameplate Dec 2030, sourced = operating + dated UC + FID (mtpa)"] = (o["Operating now (mtpa)"] + o["Under construction: dated, online by Dec 2030 (mtpa)"]
                                                                          + o["FID: dated, online by Dec 2030 (mtpa)"])
        o["Nameplate Dec 2030 incl. dated planned (mtpa)"] = o["Nameplate Dec 2030, sourced = operating + dated UC + FID (mtpa)"] + o["Planned / permitted: dated (mtpa)"]
        o["Pipeline incl. undated UC + FID (mtpa)"] = (o["Nameplate Dec 2030, sourced = operating + dated UC + FID (mtpa)"] + o["Under construction: undated (mtpa)"]
                                                       + o["FID: undated (mtpa)"])
        o["Operating now (Bcf/d, DERIVED)"] = o["Operating now (mtpa)"] * MTPA_TO_BCFD
        o["Nameplate Dec 2030, sourced (Bcf/d, DERIVED)"] = o["Nameplate Dec 2030, sourced = operating + dated UC + FID (mtpa)"] * MTPA_TO_BCFD
        return pd.Series(o)

    t = d.groupby("country").apply(agg)
    t.insert(0, "Chart group", [group_of(c) for c in t.index])
    t = t.sort_values("Operating now (mtpa)", ascending=False)
    tot = t.drop(columns="Chart group").sum()
    tot.name = "WORLD TOTAL"
    t = pd.concat([t, pd.DataFrame([tot])])
    t.loc["WORLD TOTAL", "Chart group"] = ""
    t.index.name = "country"
    return t.round(3)


# ---------------------------------------------------------------------------------------------------------------------
def check_sources(prev_rel, full):
    """Change detection on the curated documents (HEAD, else a ranged GET) + newest REQ edition."""
    s = sess()
    rows = []
    prev = {}
    if prev_rel is not None:
        prev = {r["url"]: r for _, r in prev_rel.iterrows()}
    seen = set()
    for k, (pub, doc, url, asof) in list(SRC.items()) + list(XSRC.items()):
        if url in seen:
            continue
        seen.add(url)
        row = dict(document=f"{pub}: {doc}", url=url, etag="", last_modified="", size=np.nan, sha256="", checked=str(TODAY), result="")
        try:
            try:
                r = s.head(url, timeout=(15, 40), allow_redirects=True)
            except Exception:  # noqa: BLE001
                r = None
            if r is None or r.status_code >= 400 or not (r.headers.get("ETag") or r.headers.get("Last-Modified") or r.headers.get("Content-Length")):
                r = s.get(url, timeout=(15, 120), stream=True)
                if r.headers.get("ETag") or r.headers.get("Last-Modified") or r.headers.get("Content-Length"):
                    r.close()          # headers are enough for change detection
                else:
                    h = hashlib.sha256()
                    n = 0
                    for chunk in r.iter_content(1 << 20):
                        h.update(chunk)
                        n += len(chunk)
                        if n > 40 << 20:
                            break
                    row["sha256"] = h.hexdigest()
                    row["size"] = n
                row.setdefault("size", np.nan)
                if not row["sha256"]:
                    row["size"] = float(r.headers.get("Content-Length", "nan"))
            else:
                row["size"] = float(r.headers.get("Content-Length", "nan"))
            row["etag"] = r.headers.get("ETag", "")
            row["last_modified"] = r.headers.get("Last-Modified", "")
            ok = r.status_code < 400
            p = prev.get(url)
            if not ok:
                row["result"] = f"UNREACHABLE (HTTP {r.status_code})"
            elif p is None:
                row["result"] = "FIRST SEEN"
            else:
                same = ((row["etag"] and row["etag"] == p.get("etag")) or (row["last_modified"] and row["last_modified"] == p.get("last_modified"))
                        or (row["sha256"] and row["sha256"] == p.get("sha256")) or (not row["etag"] and not row["last_modified"] and not row["sha256"]
                                                                              and row["size"] == p.get("size")))
                row["result"] = "UNCHANGED" if same else "CHANGED - re-read the document and update the curated rows"
        except Exception as e:  # noqa: BLE001
            row["result"] = f"UNREACHABLE ({type(e).__name__})"
        rows.append(row)
        print("  watch", row["result"][:34], url[:90])
    try:
        r = s.get(REQ_LIST, timeout=40)
        eds = sorted(set(re.findall(r"resources-and-energy-quarterly-([a-z]+-20\d\d)", r.text)), key=lambda e: (e.split("-")[1], ["march", "june", "september", "december"].index(e.split("-")[0]) if e.split("-")[0] in ("march", "june", "september", "december") else 0))
        newest = eds[-1] if eds else ""
        rows.append(dict(document="DISR Resources and Energy Quarterly: newest edition in the listing", url=REQ_LIST, etag="", last_modified="", size=np.nan,
                         sha256="", checked=str(TODAY),
                         result=f"newest listed: {newest}; curated rows use {REQ_EDITION}" + ("" if newest == REQ_EDITION else " - NEW EDITION: re-read section 6 LNG and the nameplate map")))
    except Exception as e:  # noqa: BLE001
        rows.append(dict(document="DISR REQ listing", url=REQ_LIST, etag="", last_modified="", size=np.nan, sha256="", checked=str(TODAY),
                         result=f"UNREACHABLE ({type(e).__name__})"))
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------------------------------------------------
def read_prev():
    if not os.path.exists(OUT):
        return None, None, None, None
    try:
        proj = pd.read_excel(OUT, sheet_name="Projects")
        eia = pd.read_excel(OUT, sheet_name="EIA rows")
        rel = pd.read_excel(OUT, sheet_name="Releases")
        hist = pd.read_excel(OUT, sheet_name="History")
        for d in (proj, eia, rel, hist):
            d.drop(columns=[c for c in d.columns if str(c).startswith("Unnamed")], inplace=True)
        for c in ("release_date", "file"):
            if c in eia:
                eia[c] = eia[c].astype(str)
        return proj, eia, rel, hist
    except Exception as e:  # noqa: BLE001
        print("  previous workbook unreadable:", type(e).__name__, e)
        return None, None, None, None


NOTES = [
    "Global LNG liquefaction capacity - 'Global supply' section of the master future-capacity table",
    "",
    "Purpose",
    "Nameplate liquefaction capacity by plant / train (existing), and every project under construction, at FID or planned, with the expected start, plus closures. It feeds the Europe LNG-need forecast: the supply side the balancing item draws on.",
    "",
    "Units",
    "capacity_value is in mtpa as published (million tonnes of LNG per year). capacity_bcfd is DERIVED: 1 mtpa = 1.36 bcm/y of gas (DISR REQ note c) = 1.36 x 35.3147 / 365.25 = 0.1315 Bcf/d. The US figures are EIA baseload nameplate (peak in the notes); other countries are nominal capacity as published.",
    "Nameplate is not output: US plants run above baseload nameplate, Qatar and the UAE export through the Strait of Hormuz (disrupted in 2026), Russian volumes are sanctioned, Nigerian and Trinidad plants are feedgas-limited.",
    "",
    "Sheets",
    "Projects: the shared row schema (one row per plant / train / project; status OPERATING is an extension of the schema's status list for existing capacity; all other statuses are the schema's).",
    "History: every change of status / expected_online / capacity_value of a row between runs (old -> new, date seen).",
    "Capacity by month: Jan 2021 - Dec 2030 nameplate by status, no ramp-up. Yellow cell = ramp months (ASSUMPTION), used only in the 'with ramp' columns.",
    "By country: nameplate now and in Dec 2030 by status, per host country. Additions by year: per chart group, additions and closures. Placement: how each project's start month was fixed.",
    "Sources / Cross-check / Tried / Releases: documents used; GIIGNL and press figures beside the official totals; what was tried or is not sourced; change detection on every source file. EIA rows: the parsed EIA workbook (history store for the US rows).",
    "",
    "Rules",
    "A number comes from a document that was opened in full; search snippets are never used. Raw official / operator figures first (EIA, DISR, QatarEnergy, operator releases and reports). GIIGNL's 2026 Annual Report supplies a number only where no official figure was opened, and then it is named as the source_publisher and in the notes; elsewhere GIIGNL is a cross-check ('Cross-check' sheet).",
    "A project with no expected date in an opened source is NOT in the dated monthly series: it is carried in the 'undated' columns (By country, Capacity by month).",
    "Date placement when a source gives only a period: year -> July, H1 -> April, H2 -> October, quarter -> middle month. A source date already past in the run month is shown from the run month.",
    "Statuses: OPERATING / COMMISSIONED_RECENT count as operating from their start month. SUSPENDED existing capacity (damaged or shut) is 'Offline'. CLOSED capacity leaves the operating total at its end month.",
    "Charts: native Excel stacked charts on 'Chart - Nameplate' and 'Chart - Additions' (redrawn every run, no borders, mmm/yy), PNGs in output/PNG Charts.",
]
NOTE_TITLES = {"Purpose", "Units", "Sheets", "Rules"}


# ---------------------------------------------------------------------------------------------------------------------
def style_sheets(path, cap_cols):
    wb = load_workbook(path)
    yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    bold = Font(bold=True)
    for ws in wb.worksheets:
        if ws.title == "Units":
            continue
        for c in ws[1]:
            c.font = bold
        ws.freeze_panes = "B2"
        for i in range(1, ws.max_column + 1):
            w = 16
            h = str(ws.cell(1, i).value or "")
            if h in ("notes", "source_document", "evidence", "detail"):
                w = 60
            elif h in ("project", "source_publisher", "operator", "document", "result", "key"):
                w = 42
            elif h == "source_url" or h == "url":
                w = 50
            ws.column_dimensions[get_column_letter(i)].width = w
    wb.save(path)


def draw_charts(path, cn, add, clo):
    """Native Excel charts: nameplate by month (country x status shade) and additions by year (country)."""
    wb = load_workbook(path)
    for nm in ("Chart - Nameplate", "Chart - Additions"):
        if nm in wb.sheetnames:
            del wb[nm]
    # --- nameplate by month: stacked area, columns = group x layer -----------------------------------------------
    ws = wb.create_sheet("Chart - Nameplate", 1)
    ws.cell(1, 1, "Month").font = Font(bold=True)
    cols = list(cn.columns)
    for j, c in enumerate(cols, start=2):
        ws.cell(1, j, c).font = Font(bold=True)
    for i, (d, row) in enumerate(cn.iterrows(), start=2):
        ws.cell(i, 1, d.to_pydatetime()).number_format = "mmm/yy"
        for j, c in enumerate(cols, start=2):
            ws.cell(i, j, round(float(row[c]), 3))
    n = len(cn) + 1
    ch = AreaChart()
    ch.grouping = "stacked"
    ch.add_data(Reference(ws, min_col=2, max_col=1 + len(cols), min_row=1, max_row=n), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=n))
    for s, c in zip(ch.series, cols):
        if c.startswith("Offline"):
            col = "4D4D4D"
        else:
            g = c.split(" - ")[0]
            base = GROUP_COLOUR[g]
            col = base if c.endswith("operating") else xlsx_charts._lighten(base, 0.45 if "UC" in c else 0.75)
        s.graphicalProperties = GraphicalProperties(solidFill=col, ln=LineProperties(noFill=True))
    ch.title = "Global LNG liquefaction nameplate, mtpa: operating (solid), under construction + FID (mid), planned (light)"
    ch.y_axis.title = "mtpa"
    ch.x_axis.number_format = "mmm/yy"
    ch.x_axis.tickLblPos = "low"
    xlsx_charts.rotated_labels(ch.x_axis)
    ch.x_axis.tickLblSkip = 6
    ch.x_axis.tickMarkSkip = 6
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.legend.position = "b"
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.plot_area.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.width, ch.height = 34, 18
    xlsx_charts.tidy_layout(ch)
    ws.add_chart(ch, f"{get_column_letter(len(cols) + 3)}2")
    # --- additions by year -----------------------------------------------------------------------------------------
    ws2 = wb.create_sheet("Chart - Additions", 2)
    ws2.cell(1, 1, "Year").font = Font(bold=True)
    for j, g in enumerate(GROUP_ORDER, start=2):
        ws2.cell(1, j, g).font = Font(bold=True)
    for i, (y, row) in enumerate(add.iterrows(), start=2):
        ws2.cell(i, 1, str(y))
        for j, g in enumerate(GROUP_ORDER, start=2):
            ws2.cell(i, j, round(float(row[g]), 3))
    n2 = len(add) + 1
    b = BarChart()
    b.type = "col"
    b.grouping = "stacked"
    b.overlap = 100
    b.gapWidth = 40
    b.add_data(Reference(ws2, min_col=2, max_col=1 + len(GROUP_ORDER), min_row=1, max_row=n2), titles_from_data=True)
    b.set_categories(Reference(ws2, min_col=1, min_row=2, max_row=n2))
    for s, g in zip(b.series, GROUP_ORDER):
        s.graphicalProperties = GraphicalProperties(solidFill=GROUP_COLOUR[g], ln=LineProperties(noFill=True))
    b.title = "Global LNG liquefaction capacity additions by year, mtpa (sourced start dates; undated projects excluded)"
    b.y_axis.title = "mtpa added"
    b.x_axis.tickLblPos = "low"
    b.x_axis.delete = False
    b.y_axis.delete = False
    b.legend.position = "b"
    b.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    b.plot_area.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    b.width, b.height = 28, 15
    xlsx_charts.tidy_layout(b)
    ws2.add_chart(b, "K2")
    xlsx_charts.save_atomic(wb, path)


def add_cap_month_sheet_extras(path, ramp_cols):
    """Yellow ramp input + 'with ramp' formula columns on 'Capacity by month' (rows: header 1, data from 2)."""
    wb = load_workbook(path)
    ws = wb["Capacity by month"]
    ncol = ws.max_column
    yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    # ramp input in a fixed cell to the right of the table
    ic = ncol + 3
    ws.cell(1, ic, "Ramp months (ASSUMPTION, editable)").font = Font(bold=True)
    c = ws.cell(2, ic, RAMP_DEFAULT)
    c.fill = yellow
    ws.cell(3, ic, "Used only in the 'with ramp' columns: each addition after the run month rises linearly from 0 to nameplate over this many months, starting in its start month.").alignment = Alignment(wrap_text=True)
    ws.column_dimensions[get_column_letter(ic)].width = 36
    ref = f"${get_column_letter(ic)}$2"
    rc = wb["Ramp calc"]
    last_row = rc.max_row
    # 'Ramp calc': row 1 header, row 2 capacity, row 3 start month, data rows 4.. one per month
    nproj = rc.max_column - 3
    sumcols = (get_column_letter(rc.max_column - 1), get_column_letter(rc.max_column))
    hdr = ["Operating + UC + FID, with ramp (mtpa, ASSUMPTION)", "Incl. dated planned, with ramp (mtpa, ASSUMPTION)"]
    for k, h in enumerate(hdr):
        ws.cell(1, ncol + 1 + k, h).font = Font(bold=True)
        ws.column_dimensions[get_column_letter(ncol + 1 + k)].width = 24
    for r in range(2, ws.max_row + 1):
        rr = r + 2           # month row on Ramp calc (data from row 4)
        ws.cell(r, ncol + 1, f"=B{r}+'Ramp calc'!{sumcols[0]}{rr}")
        ws.cell(r, ncol + 2, f"=B{r}+'Ramp calc'!{sumcols[0]}{rr}+'Ramp calc'!{sumcols[1]}{rr}")
    # the ramp formulas on 'Ramp calc' reference the input cell
    for cidx in range(2, rc.max_column - 1):
        for r in range(4, last_row + 1):
            rc.cell(r, cidx, f"=IF(ISNUMBER({get_column_letter(cidx)}$3),{get_column_letter(cidx)}$2*MAX(0,MIN(1,((YEAR($A{r})-YEAR({get_column_letter(cidx)}$3))*12+MONTH($A{r})-MONTH({get_column_letter(cidx)}$3)+1)/'Capacity by month'!{ref})),0)")
            rc.cell(r, cidx).number_format = "0.00"
    xlsx_charts.save_atomic(wb, path)


def ramp_sheet(df):
    """Ramp calc frame: one column per dated future project (UC + FID, then planned), header rows 2 = capacity, 3 = start."""
    proj = df[(df["_kind"] == "project") & df["capacity_value"].notna() & df["_plc"].notna()].copy()
    proj["lay"] = proj.apply(layer_of, axis=1)
    proj = proj[proj["lay"].isin(["UC", "FID", "PL"])]
    # also the recent commissioning after the run month is handled by status UC/FID; operating ones are in column B
    order = pd.concat([proj[proj["lay"].isin(["UC", "FID"])], proj[proj["lay"] == "PL"]])
    return order


def write_all(df, hist, eia, rel, cap, by_ctry, add, clo, cn, placement):
    os.makedirs(OUT_DIR, exist_ok=True)
    proj_out = df[SCHEMA].copy()
    for c in ("capacity_value", "capacity_bcfd", "capacity_mw"):
        proj_out[c] = pd.to_numeric(proj_out[c], errors="coerce")
    sheets = {
        "Projects": proj_out.set_index("region"),
        "History": hist.set_index("key") if len(hist) else pd.DataFrame(columns=["field", "old", "new", "date_seen"], index=pd.Index([], name="key")),
        "Capacity by month": cap,
        "By country": by_ctry,
        "Additions by year": pd.concat([add.add_suffix(" (additions)"), clo.add_suffix(" (offline/closed)")], axis=1).rename_axis("year"),
        "Placement": placement.set_index("project"),
        "Sources": sources_frame().set_index("publisher"),
        "Cross-check": cross_frame(df).set_index("country"),
        "Tried": pd.DataFrame(TRIED, columns=["what", "url", "result"]).set_index("what"),
        "Releases": rel.set_index("document"),
        "EIA rows": eia.set_index("sheet"),
    }
    # Ramp calc skeleton (formulas added later)
    order = ramp_sheet(df)
    rc = pd.DataFrame(index=months())
    rc.index.name = "Month"
    for _, p in order.iterrows():
        rc[p["project"][:60]] = np.nan
    rc["Operating UC + FID, ramped (sum)"] = np.nan
    rc["Planned dated, ramped (sum)"] = np.nan
    sheets["Ramp calc"] = rc
    xlsx_notes.write_workbook(OUT, sheets, NOTES, NOTE_TITLES)
    # ---- Ramp calc: layout rows 1 header / 2 capacity / 3 start, data from row 4 -----------------------------------
    wb = load_workbook(OUT)
    ws = wb["Ramp calc"]
    ws.insert_rows(2, 2)          # rows 2,3 for capacity and start; data now start at row 4
    n_uc = int(order["lay"].isin(["UC", "FID"]).sum())
    for k, (_, p) in enumerate(order.iterrows()):
        col = 2 + k
        ws.cell(2, col, float(p["capacity_value"]))
        ws.cell(3, col, effective_start(p).to_pydatetime()).number_format = "mmm/yy"
    ws.cell(2, 1, "capacity (mtpa)")
    ws.cell(3, 1, "start month")
    # sum columns
    c_first, c_uc_last = 2, 1 + n_uc
    c_pl_first, c_pl_last = 2 + n_uc, 1 + len(order)
    sum1, sum2 = ws.max_column - 1, ws.max_column
    for r in range(4, ws.max_row + 1):
        if n_uc:
            ws.cell(r, sum1, f"=SUM({get_column_letter(c_first)}{r}:{get_column_letter(c_uc_last)}{r})")
        else:
            ws.cell(r, sum1, 0)
        if len(order) > n_uc:
            ws.cell(r, sum2, f"=SUM({get_column_letter(c_pl_first)}{r}:{get_column_letter(c_pl_last)}{r})")
        else:
            ws.cell(r, sum2, 0)
        ws.cell(r, 1).number_format = "mmm/yy"
    wb.save(OUT)
    add_cap_month_sheet_extras(OUT, None)
    wb = load_workbook(OUT)
    ws = wb["Capacity by month"]
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 1).number_format = "mmm/yy"
    wb.save(OUT)
    style_sheets(OUT, None)
    draw_charts(OUT, cn, add, clo)


def sources_frame():
    rows = []
    for k, (pub, doc, url, asof) in SRC.items():
        rows.append(dict(publisher=pub, role="Official / operator source of Projects numbers", document=doc, url=url, as_of_date=asof, used_for=SRC_USE.get(k, "")))
    rows.append(dict(publisher="U.S. Energy Information Administration", role="Official / operator source of Projects numbers",
                     document="U.S. Liquefaction Capacity workbook (newest quarterly file; sheets Existing & Under Construction, Approved)",
                     url=EIA_BASE.format(y="YYYY", q="Q"), as_of_date="see 'Releases'",
                     used_for="United States trains: baseload and peak nameplate, status, in-service dates; approved-not-built projects (cites FERC / DOE dockets)."))
    for k, (pub, doc, url, asof) in XSRC.items():
        rows.append(dict(publisher=pub, role="GIIGNL: number used only where no official figure was opened (labelled); otherwise cross-check. Others: dates / background", document=doc, url=url, as_of_date=asof, used_for=XSRC_USE.get(k, "")))
    return pd.DataFrame(rows)


SRC_USE = {
    "REQ": "Australian plant nameplates (map, 88 Mtpa total); GTA 2.3 Mtpa; Congo Nguya 2.4 Mtpa; Mozambique LNG 13.2 Mtpa from 2029; Qatar 77 -> 133 Mt by 2031 trajectory; Ras Laffan / Hormuz context.",
    "QE_DECK": "Qatar 77 mtpa (14 trains) now; NFE 32 (4 x 8), NFS 16 (2 x 8), NFW 16 (2 x 8); 142 mtpa by 2030.",
    "QE_AR24": "Qatar programme background (no per-train dates).",
    "CHEN_Q2": "Corpus Christi Stage 3 train dates; Sabine Pass + Corpus Christi operating capacity (~55 mtpa).",
    "GOLAR_Q2": "Esperanza 3.5 MTPA (H2 2028); fourth FLNG 3.5 MTPA (end-2029).",
    "ENI_FID": "Coral North 3.6 MTPA, 2028.",
    "NLNG_FF": "Nigeria 22 MTPA, Train 7 to 30 MTPA.",
    "NOVATEK_FID": "Arctic LNG 2 6.6 mtpa per train (design).",
    "SHELL_LNGC": "LNG Canada Phase 1 = 14 mtpa; Phase 2 FID, +14 mtpa, early 2030s.",
    "MOZ_SITE": "Mozambique LNG 13.12 MTPA.",
    "WOODFIBRE": "Woodfibre 2.1 mtpa, completion by 2027.",
    "CHEN_SP": "Sabine Pass Stage 5 (> 6 mtpa).",
}
XSRC_USE = {
    "GIIGNL": "Numbers where no official figure was opened (existing non-US plants outside Australia / Nigeria / Qatar total, damaged Ras Laffan trains, NWS Train 2 closure, Marsa, Ruwais, Abadi, Ksi Lisims, Hilli, ECA capacity, start years of LNG Canada / GTA / Arctic LNG 2); cross-checks on 'Cross-check'.",
    "WDS_Q2": "Pluto Train 2 / Scarborough dates (first LNG Q4 2026) - no capacity, so not in Projects ('Tried').",
    "WDS_H1": "NWS four onshore trains (note on the NWS row); Pluto Train 2 dates; Louisiana LNG 16.5 mtpa (its EIA row is the Projects source).",
    "TTE_Q2": "ECA LNG start-up early July 2026 - no capacity, so not in Projects ('Tried').",
    "PEMBINA_Q2": "Cedar LNG first exports late 2028 - no capacity, so not in Projects ('Tried').",
    "REUTERS_QE": "Press copy: NFE first train timing, Ras Laffan damage - 'Cross-check' only.",
}


def cross_frame(df):
    d = df[df["_kind"].notna() & df["capacity_value"].notna()]
    op = d[d.apply(layer_of, axis=1) == "OP"]
    op = op[op["status"] != "CLOSED"]
    rows = []
    for country, item, pub, val, text, cmp in CROSS:
        raw = np.nan
        if cmp == "WORLD":
            raw = float(op["capacity_value"].sum())
        elif cmp:
            raw = float(op[op["country"] == cmp]["capacity_value"].sum())
        rows.append(dict(country=country, item=item, publisher=pub, cross_check_value_mtpa=val if val is not None else np.nan,
                         raw_operating_in_Projects_mtpa=raw, difference_mtpa=(raw - val) if (val is not None and not np.isnan(raw)) else np.nan, text=text))
    t = pd.DataFrame(rows)
    return t


def placement_frame(df):
    rows = []
    for _, p in df[df["_kind"].notna()].iterrows():
        eff = effective_start(p)
        rows.append(dict(project=p["project"], country=p["country"], layer=layer_of(p), source_date_text=sv(p["_when"]) or "", rule=placement_rule(sv(p["_when"]) or "None"),
                         placed_month=eff.strftime("%Y-%m") if eff is not None and not pd.isna(eff) else "",
                         start_not_in_raw_source="yes - operating from run month" if p["_startunk"] else "", source_date_already_past="yes - shown from run month" if (p["_kind"] == "project" and p["_plc"] is not None and not pd.isna(p["_plc"]) and p["_plc"] < NOW_M) else "",
                         end_month=sv(p["_end"]) or "", offline_from=sv(p["_offline_from"]) or "", chart_group=p["_group"]))
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------------------------------------------------
def draw_pngs(cap, cn, add, clo):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    os.makedirs(PNG_DIR, exist_ok=True)
    plt.rcParams.update({"font.size": 9})
    fig, ax = plt.subplots(figsize=(14, 7.5))
    cols = list(cn.columns)
    colours = []
    for c in cols:
        if c.startswith("Offline"):
            colours.append("#4D4D4D")
        else:
            g = c.split(" - ")[0]
            base = "#" + GROUP_COLOUR[g]
            colours.append(base if c.endswith("operating") else "#" + xlsx_charts._lighten(GROUP_COLOUR[g], 0.45 if "UC" in c else 0.75))
    ax.stackplot(cn.index, [cn[c].values for c in cols], colors=colours, linewidth=0)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.set_ylabel("mtpa")
    ax.set_title("Global LNG liquefaction nameplate by month (mtpa): operating, under construction + FID, planned (light); sourced dates only, no ramp-up", loc="left", fontsize=11, fontweight="bold")
    import matplotlib.dates as mdates
    ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1, 7]))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b/%y"))
    ax.axvline(NOW_M, color="black", lw=0.8, ls=":")
    ax.text(NOW_M, ax.get_ylim()[1] * 0.98, " run month", fontsize=8, va="top")
    # legend: one row per group (operating colour), plus shade key
    from matplotlib.patches import Patch
    h = [Patch(color="#" + GROUP_COLOUR[g], label=g) for g in GROUP_ORDER] + [Patch(color="#4D4D4D", label="Offline (damaged / shut-in)")]
    h += [Patch(facecolor="#bbbbbb", label="solid = operating"), Patch(facecolor="#dddddd", label="mid = UC + FID (dated)"), Patch(facecolor="#f2f2f2", edgecolor="#bbbbbb", label="light = planned (dated)")]
    ax.legend(handles=h, ncol=6, loc="upper center", bbox_to_anchor=(0.5, -0.08), frameon=False)
    fig.text(0.01, 0.005, "Source: EIA U.S. liquefaction capacity workbook; DISR Resources and Energy Quarterly Sep 2026; QatarEnergy; operator releases; GIIGNL Annual Report 2026 (cross-check). "
             "Projects without a sourced date are not shown.", fontsize=7)
    fig.tight_layout(rect=(0, 0.02, 1, 1))
    p1 = os.path.join(PNG_DIR, f"{BASENAME}__Nameplate_by_month.png")
    fig.savefig(p1, dpi=130)
    plt.close(fig)
    fig, ax = plt.subplots(figsize=(12, 6.5))
    bottom = np.zeros(len(add))
    for g in GROUP_ORDER:
        ax.bar([str(y) for y in add.index], add[g].values, bottom=bottom, color="#" + GROUP_COLOUR[g], label=g, width=0.7)
        bottom += add[g].values
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.set_ylabel("mtpa added")
    ax.set_title("Global LNG liquefaction capacity additions by year (mtpa) - sourced start dates; projects without a date excluded", loc="left", fontsize=11, fontweight="bold")
    for i, v in enumerate(bottom):
        ax.text(i, v + 0.3, f"{v:.1f}", ha="center", fontsize=8)
    ax.legend(ncol=8, loc="upper center", bbox_to_anchor=(0.5, -0.07), frameon=False)
    fig.text(0.01, 0.005, "Source: as the nameplate chart. 2021-26 = capacity that came online (actual); 2027-30 = expected start dates, year-only dates placed mid-year.", fontsize=7)
    fig.tight_layout(rect=(0, 0.02, 1, 1))
    p2 = os.path.join(PNG_DIR, f"{BASENAME}__Additions_by_year.png")
    fig.savefig(p2, dpi=130)
    plt.close(fig)
    return p1, p2


# ---------------------------------------------------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--full", action="store_true", help="re-read every whole file")
    ap.add_argument("--no-check", action="store_true", help="skip the change-detection requests on the curated documents")
    a = ap.parse_args()
    print("Global LNG liquefaction", TODAY)
    prev_proj, prev_eia, prev_rel, prev_hist = read_prev()
    print("  previous workbook:", "found" if prev_proj is not None else "none (clean run)")
    eia, eia_rel = fetch_eia(prev_rel, prev_eia, a.full)
    df, hist_new = build_projects(eia, prev_proj)
    hist = pd.concat([prev_hist, hist_new], ignore_index=True) if prev_hist is not None and len(prev_hist) else hist_new
    rel_rows = [eia_rel]
    if a.no_check and prev_rel is not None:
        rel = prev_rel
    else:
        rel = pd.concat([pd.DataFrame(rel_rows), check_sources(prev_rel, a.full)], ignore_index=True)
    for c in ("document", "url", "etag", "last_modified", "size", "sha256", "checked", "result"):
        if c not in rel:
            rel[c] = ""
    rel = rel[["document", "url", "etag", "last_modified", "size", "sha256", "checked", "result"] + [c for c in rel.columns if c not in
                                                                                                     ("document", "url", "etag", "last_modified", "size", "sha256", "checked", "result")]]
    cap, bg = build_capacity_table(df)
    cn = chart_tables(bg, cap.index)
    add, clo = additions_by_year(df)
    bc = by_country(df)
    placement = placement_frame(df)
    cap_out = cap.copy()
    write_all(df, hist, eia, rel, cap_out, bc, add, clo, cn, placement)
    p1, p2 = draw_pngs(cap, cn, add, clo)
    print("wrote", OUT)
    print("PNGs", p1, p2)
    # ---- console summary --------------------------------------------------------------------------------------------
    now = cap.loc[NOW_M]
    e30 = cap.loc[LAST_M]
    print(f"  nameplate now ({NOW_M:%b/%y}): operating {now.iloc[0]:.1f} mtpa, offline {now.iloc[1]:.1f}")
    print(f"  Dec 2030: operating {e30.iloc[0]:.1f} + UC {e30.iloc[2]:.1f} + FID {e30.iloc[3]:.1f} = {e30.iloc[4]:.1f}; planned dated {e30.iloc[5]:.1f}; "
          f"undated UC {e30.iloc[7]:.1f} FID {e30.iloc[8]:.1f} planned {e30.iloc[9]:.1f}")
    print(bc[["Operating now (mtpa)", "Nameplate Dec 2030, sourced = operating + dated UC + FID (mtpa)"]].to_string())
    print(add.round(1).to_string())


if __name__ == "__main__":
    main()
