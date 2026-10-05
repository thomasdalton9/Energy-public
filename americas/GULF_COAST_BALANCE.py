"""
Gulf Coast gas balance, Texas + Louisiana, monthly Bcf/d, history from 2021 and forecast to Dec 2030 -> gulf_coast_gas_balance.xlsx.
Question it answers: is there enough Permian gas for the new Gulf LNG, or does the Gulf need more Haynesville (or other-region) supply?
Runs in GitHub Actions (.github/workflows/gulf_coast_balance.yml; EIA_API_KEY secret). Reads, never writes, the Texas workbooks.

  TEXAS       texas_production_forecast.xlsx ('Supply and demand' to 2023, 'Demand to 2033' from 2024: consumption by sector, Mexico pipeline,
              BASE data-centre gas burn, dry production base / takeaway-delayed, STEO basin series, takeaway table) and the LNG feedgas model of
              TEXAS_GAS_FORECAST.py run to Dec 2030 (base and 6-month delayed). All Texas inputs and their status labels are the Texas workbooks'.
  LOUISIANA   EIA API v2 (natural-gas/cons/sum, prod/sum, move/poe2; area SLA): consumption by sector, marketed / dry production, LNG exports
              of Sabine Pass, Cameron, Calcasieu Pass and Plaquemines (all four are Louisiana ports; feedgas = exports x 1.09 as for Texas).
              Forecast: existing plants = peak nameplate (EIA liquefaction capacity file) x 1.09 x 3-year seasonal utilisation; new trains from the
              editable 'Assump - LA LNG' tab (date status SOURCED only where an opened document gives it, else UNVERIFIED); base and 6-month delay;
              consumption = seasonal trend; production = STEO Haynesville x a Louisiana share + calibrated remainder.
  HAYNESVILLE STEO Haynesville (NGMPHA) is East Texas + Louisiana. Split: Texas 30% (the Texas workbook's assumption, kept so the two states
              add up to STEO) and Louisiana 70%; EIA state data give the implied Louisiana share (LA marketed / STEO Haynesville) as a check.
  COMBINED    production, demand incl. LNG and implied net outflow of the two states together, and the key-question table: growth in Gulf LNG
              feedgas (and other Gulf demand) from Dec 2025 against growth in Permian (takeaway-capped), Eagle Ford and Haynesville supply, and the
              extra supply needed to hold the outflow, base and delayed.
Beyond Dec 2027 the basin supply is an extension of EIA STEO (2028 by the STEO 2027/2026 growth, 2029-30 damped x0.5 a year, as the Texas
workbook), not an EIA forecast; 2029-30 is lighter on the charts.

Usage: python3 GULF_COAST_BALANCE.py [--out "output/Data and Chart Outputs/gulf_coast_gas_balance.xlsx"] [--full] [--fetch-only]
"""
import argparse
import os
import re
import sys

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import xlsx_notes  # noqa: E402
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import TEXAS_GAS as tg  # noqa: E402  (EIA helpers, read-only use)
import TEXAS_GAS_FORECAST as gf  # noqa: E402  (LNG model, read-only use)

OUT_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
DEFAULT_OUT = os.path.join(OUT_DIR, "gulf_coast_gas_balance.xlsx")
TEXAS_XLSX = os.path.join(OUT_DIR, "texas_gas_monthly.xlsx")
PROD_XLSX = os.path.join(OUT_DIR, "texas_production_forecast.xlsx")
LNG_XLSX = os.path.join(OUT_DIR, "lng_feedgas_daily.xlsx")
START, END = pd.Timestamp("2021-01-01"), pd.Timestamp("2030-12-01")
EXT_FROM = pd.Timestamp("2029-01-01")          # basin supply beyond the Texas workbook's 2028 horizon: damped extension
RAW = "Raw MMcf"
SECT = ["Electric power", "Industrial", "Residential", "Commercial", "Vehicle fuel"]
RATIO = 1.09                                    # feedgas / LNG exports, as TEXAS_GAS_FORECAST.py
LA_PORTS = {"YSPL": "Sabine Pass", "YCAM": "Cameron", "YCCPL": "Calcasieu Pass", "YPLAQ": "Plaquemines"}
SRC_EIA = "EIA U.S. liquefaction capacity file 2026 Q2 (opened 5 Oct 2026)"
AL = "Assump - LA LNG"
FILL_IN = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
FILL_HEAD = PatternFill("solid", start_color="DDEBF7", end_color="DDEBF7")
BOLD = Font(bold=True)

# EIA peak nameplate capacity of the operating Louisiana plants, Bcf/d of LNG (EIA liquefaction capacity file 2026 Q2: per-train peak x trains)
EXISTING = {"Sabine Pass": (6 * 0.7588767123287671, "6 trains x 0.759 Bcf/d peak (EIA 2026 Q2); EIA Today in Energy 15 Sep 2026: 3.6 nominal / 4.6 peak"),
            "Cameron": (3 * 0.66, "3 trains x 0.66 Bcf/d peak (EIA 2026 Q2)"),
            "Calcasieu Pass": (2 * 0.7903397260273972, "2 train groups x 0.790 Bcf/d peak (EIA 2026 Q2)"),
            "Plaquemines": (1.89 + 1.96, "Phase 1 1.89 + Phase 2 1.96 Bcf/d peak (EIA 2026 Q2; Venture Global: peak > 28 MTPA once complete)")}

# New Louisiana projects: (plant, train, EIA peak nameplate Bcf/d LNG, first-LNG month or None, ramp months, ramp-start util, steady util,
# date status, note). Feedgas nameplate = peak x 1.09 (how the Texas table's feedgas figures relate to EIA peak: Golden Pass 0.795 -> 0.865).
LA_TRAINS = [
    ("CP2 (Venture Global)", "Phase 1, Trains 1-26", 2.86, "2027-10", 12, 0.15, 0.95, "SOURCED",
     "EIA 2026 Q2 file: under construction, in service 2H2027 (month = mid-half, ASSUMED); ramp 12 months ASSUMED (Plaquemines took ~1 year from first LNG to full output; EIA 1 Sep 2026: exporting at full capacity)"),
    ("CP2 (Venture Global)", "Phase 2, Trains 27-36", 1.10, "2029-04", 8, 0.15, 0.95, "SOURCED",
     "EIA 2026 Q2 file: under construction, 1H2029 (month = mid-half, ASSUMED); ramp ASSUMED"),
    ("Woodside Louisiana LNG", "Phase 1, Trains 1-3", 2.326520547945205, "2029-07", 18, 0.10, 0.95, "SOURCED",
     "EIA 2026 Q2 file: under construction, 2029 (month = mid-year, ASSUMED); three trains staggered, ramp 18 months ASSUMED (a search snippet of Woodside's Q2 2026 report says first LNG 2029, 28% complete - not opened)"),
    ("Commonwealth LNG", "Trains 1-6", 1.1845263157894736, "2030-07", 12, 0.10, 0.95, "SOURCED",
     "EIA 2026 Q2 file: under construction, 2030 (month = mid-year, ASSUMED); six trains, ramp 12 months ASSUMED"),
    ("Delfin FLNG", "Vessel 1", 0.6, "2030-07", 4, 0.30, 0.95, "SOURCED",
     "EIA 2026 Q2 file: under construction, 2030 (month = mid-year, ASSUMED); floating vessel; DOE extended its start deadline to 1 Jun 2031"),
    # pre-FID or without a sourced date: no first-LNG month, so nothing is forecast until the owner types one
    ("Cameron LNG", "Train 4", 0.96, None, 8, 0.15, 0.95, "UNVERIFIED",
     "EIA 2026 Q2 'Approved' tab: new single larger train, EPC awarded to Bechtel, FID TBA; DOE start deadline Mar 2033. No date -> not in the forecast"),
    ("Woodside Louisiana LNG", "Phase 2, Trains 4-5", 1.4542250958, None, 12, 0.10, 0.95, "UNVERIFIED",
     "EIA 2026 Q2 'Approved' tab: limited notice to proceed, FID TBA. No date -> not in the forecast"),
    ("Lake Charles LNG (Energy Transfer)", "Trains 1-3", 2.1734342465753427, None, 12, 0.10, 0.95, "UNVERIFIED",
     "EIA 2026 Q2 'Approved' tab: Energy Transfer announced on 18 Dec 2025 it suspended development (open to third parties). No date -> not in the forecast"),
    ("Delfin FLNG", "Vessels 2-3", 1.0533333333, None, 4, 0.30, 0.95, "UNVERIFIED",
     "EIA 2026 Q2 'Approved' tab: EPC awarded to Samsung Heavy / Black & Veatch, FID TBA. No date -> not in the forecast"),
    ("Sabine Pass Stage 5 (Cheniere)", "Expansion", None, None, 12, 0.10, 0.95, "UNVERIFIED",
     "FERC page (opened 5 Oct 2026): FERC staff issued the final EIS on 25 Sep 2026. Not in EIA's 2026 Q2 file; capacity, FID and dates not found -> not in the forecast"),
    ("Plaquemines expansion (Venture Global)", "Expansion", None, None, 12, 0.10, 0.95, "UNVERIFIED",
     "Venture Global site (opened 5 Oct 2026): 'additional plans for expansion'. Not in EIA's file; no capacity or dates -> not in the forecast"),
]


# --------------------------------------------------------------------------- Louisiana EIA pull
def fetch(store, full):
    new = {}

    def put(col, period, v):
        if v is not None:
            new.setdefault(col, {})[pd.Timestamp(period + "-01")] = v

    for prefix, route, procs in (("cons|", "cons/sum/data/", tg.CONS), ("prod|", "prod/sum/data/", tg.PROD),
                                 ("stor|", "stor/sum/data/", tg.STOR)):
        st = tg.start_for(store, prefix, full)
        data = tg.rows(route, {"duoarea": ["SLA"], "process": list(procs)}, st)
        print(f"  LA {prefix} from {st}: {len(data)} rows", flush=True)
        for r in data:
            put(prefix + procs[r["process"]], r["period"], tg.volume(r))
    st = tg.start_for(store, "ENG|", full)
    sids = [f"NGM_EPG0_ENG_{k}-Z00_MMCF" for k in LA_PORTS]
    data = tg.rows("move/poe2/data/", {"series": sids}, st)
    print(f"  LA ENG| from {st}: {len(data)} rows", flush=True)
    for r in data:
        m = re.match(r"NGM_EPG0_ENG_(Y\w+?)-Z00_MMCF", str(r.get("series", "")))
        if m and m[1] in LA_PORTS:
            put("ENG|" + LA_PORTS[m[1]], r["period"], tg.volume(r))
    add = pd.DataFrame({c: pd.Series(v) for c, v in new.items()}).sort_index()
    if store.empty:
        merged = add
    else:
        merged = store.reindex(store.index.union(add.index)).reindex(columns=store.columns.union(add.columns))
        merged.update(add)
    merged.index.name = "Month"
    return merged.sort_index().dropna(how="all")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--full", action="store_true")
    ap.add_argument("--fetch-only", action="store_true")
    a = ap.parse_args()
    store = tg.load_store(a.out)
    raw = fetch(store, a.full)
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    if a.fetch_only:
        xlsx_notes.write_workbook(a.out, {RAW: raw}, ["Notes", "Fetch-only run of americas/GULF_COAST_BALANCE.py: Louisiana EIA store only."], set())
        print(raw.tail(6).T.to_string())
        return
    import GULF_MODEL as gm  # noqa: E402  (placeholder, replaced below)
    gm.run(a.out, raw)


if __name__ == "__main__":
    main()
