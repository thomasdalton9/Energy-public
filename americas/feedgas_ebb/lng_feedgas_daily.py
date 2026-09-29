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
    ("Sabine Pass", "transco", "TRANSCO", "9009310", "Transco (Gulf Trace) -> Sabine Pass, Lighthouse Road"),
    ("Sabine Pass", "km", "NGPL", "46622", "NGPL -> Sabine Pass Liquefaction"),
    # Creole Trail's own delivery meter into the plant. Its receipts
    # include the Texas Eastern (25,000) and Trunkline (259,976) Creole
    # Trail deliveries previously counted here, plus LEAP and Acadian -
    # those upstream meters are dropped so nothing is counted twice.
    ("Sabine Pass", "cheniere", "200", "CT200111", "Creole Trail -> Sabine Pass Liquefaction"),
    ("Sabine Pass", "cheniere", "200", "SPLNGD", "Creole Trail -> Sabine Pass LNG"),
    # Gator Express's own delivery meter into the plant = TGP + Texas Eastern
    # + Columbia Gulf receipts (checked equal on 27 Sep 2026), with history
    ("Plaquemines", "quorum", "VGPPB1IPWS:2", "VGPQD", "Gator Express -> Plaquemines LNG (whole plant)"),
    # Wilkinson Bayou is in Plaquemines Parish - Columbia Gulf's feed into
    # Gator Express (Venture Global's filings name Columbia Gulf, TGP and
    # Texas Eastern as the plant's supply pipes). Counterparty not yet
    # confirmed from Columbia Gulf's location file (it timed out).
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
    ("Freeport", "gulfsouth", "GS", "24329", "Gulf South -> Stratton Ridge (to Freeport LNG)"),
    ("Elba Island", "km", "EEC", "660700", "Elba Express -> Elba Liquefaction, Chatham"),
    # Golden Pass Pipeline's delivery meter into the terminal (gasnom, like
    # Cameron Interstate) - Gulf Run's deliveries into GPP partly go elsewhere
    ("Golden Pass", "gasnom", "goldenpass", "1097217", "Golden Pass Pipeline -> Golden Pass LNG terminal"),
    # Cove Point pipeline's delivery into the plant (BHE GT&S posting).
    ("Cove Point", "bhe", "cpl", "10001", "Cove Point pipeline -> Cove Point plant"),
]

# Approximate first-feedgas month per plant (commissioning, a little
# before first cargo) - blanks before these are "not operating yet", not
# missing data. Used to keep pre-start days out of the progress log's
# "missing" list, out of the calibration (plus RAMP_MONTHS of start-up
# after), and to stop an unmetered plant's EIA placeholder being applied
# before it existed. Edit if you have better dates.
PLANT_START = {}  # filled from TRAIN_STARTS below: each plant's first train
RAMP_MONTHS = 3

# Nameplate liquefaction capacity by phase, as published by the
# operators (million tonnes per annum of LNG). Start-up = first LNG (month,
# approximate); the date FERC staff authorised fuel gas / hazardous
# fluids into that phase's facilities is kept for reference, with the
# letter's eLibrary accession number. Plants routinely run above nameplate (Sabine Pass, Calcasieu
# Pass, Plaquemines), so utilisation over 100% is normal. Edit as phases
# are added.
PLANT_CAPACITY = [  # plant, phase, trains, nameplate mtpa, FERC fuel gas authorised, FERC accession, first LNG
    ("Sabine Pass", "Trains 1-6", 6, 30.0, date(2015, 9, 23), "20150923-3012", date(2016, 2, 1)),
    ("Cove Point", "Train 1", 1, 5.25, date(2017, 8, 31), "20170831-3071", date(2018, 3, 1)),
    ("Corpus Christi", "Stage 1-2 (Trains 1-3)", 3, 15.0, date(2018, 8, 16), "20180816-3034", date(2018, 11, 1)),
    ("Corpus Christi", "Stage 3 (7 midscale trains)", 7, 10.0, date(2024, 7, 24), "20240724-3089", date(2024, 12, 1)),
    ("Cameron", "Trains 1-3", 3, 12.0, date(2018, 11, 5), "20181105-3030", date(2019, 5, 1)),
    ("Freeport", "Trains 1-3", 3, 15.0, date(2019, 2, 19), "20190219-3025", date(2019, 8, 1)),
    ("Elba Island", "10 small-scale units", 10, 2.5, date(2019, 2, 1), "20190201-3039", date(2019, 10, 1)),
    ("Calcasieu Pass", "18 midscale trains (9 blocks)", 18, 10.0, date(2021, 4, 21), "20210421-3018", date(2022, 1, 1)),
    ("Plaquemines", "Phase 1", 18, 13.3, date(2024, 4, 23), "20240423-3041", date(2024, 12, 1)),
    ("Plaquemines", "Phase 2", 18, 6.7, date(2025, 1, 27), "20250127-3027", date(2025, 3, 1)),
    ("Golden Pass", "Trains 1-3", 3, 18.1, date(2025, 3, 19), "20250319-3021", date(2025, 11, 1)),
]

# Train-by-train start-ups: the date FERC staff granted each liquefaction
# train / block its feed gas (or hazardous fluids - Venture Global's
# wording) - the point it can start making LNG - with the letter's
# eLibrary accession number. Elba units 1-5 only have "commence service"
# letters. Nameplate mtpa is the phase's split evenly across its trains.
# Plant start-ups (PLANT_START) and the capacity chart step up from these.
TRAIN_STARTS = [  # plant, phase, train/block, nameplate mtpa, start, FERC accession, basis
    ("Sabine Pass", "Trains 1-6", "Train 1", 5.0, date(2015, 11, 19), "20151119-3080", "feed gas"),
    ("Sabine Pass", "Trains 1-6", "Train 2", 5.0, date(2016, 4, 18), "20160418-3037", "feed gas"),
    ("Sabine Pass", "Trains 1-6", "Train 3", 5.0, date(2016, 11, 8), "20161108-3029", "feed gas"),
    ("Sabine Pass", "Trains 1-6", "Train 4", 5.0, date(2017, 6, 1), "20170601-4015", "feed gas"),
    ("Sabine Pass", "Trains 1-6", "Train 5", 5.0, date(2018, 9, 6), "20180906-3061", "feed gas"),
    ("Sabine Pass", "Trains 1-6", "Train 6", 5.0, date(2021, 9, 22), "20210922-3034", "feed gas"),
    ("Cove Point", "Train 1", "Train 1", 5.25, date(2017, 8, 31), "20170831-3071", "hazardous fluids"),
    ("Corpus Christi", "Stage 1-2", "Train 1", 5.0, date(2018, 8, 16), "20180816-3034", "feed gas"),
    ("Corpus Christi", "Stage 1-2", "Train 2", 5.0, date(2019, 3, 11), "20190311-3000", "feed gas"),
    ("Corpus Christi", "Stage 1-2", "Train 3", 5.0, date(2020, 9, 24), "20200924-3032", "feed gas"),
    ("Corpus Christi", "Stage 3", "Midscale Train 1", 1.4286, date(2024, 12, 23), "20241223-3028", "feed gas"),
    ("Corpus Christi", "Stage 3", "Midscale Train 2", 1.4286, date(2025, 6, 5), "20250605-3048", "feed gas"),
    ("Corpus Christi", "Stage 3", "Midscale Train 3", 1.4286, date(2025, 8, 28), "20250828-3093", "feed gas"),
    ("Corpus Christi", "Stage 3", "Midscale Train 4", 1.4286, date(2025, 10, 20), "20251020-3039", "feed gas"),
    ("Corpus Christi", "Stage 3", "Midscale Train 5", 1.4286, date(2026, 2, 9), "20260209-3069", "feed gas"),
    ("Corpus Christi", "Stage 3", "Midscale Train 6", 1.4286, date(2026, 4, 24), "20260424-3056", "feed gas"),
    ("Corpus Christi", "Stage 3", "Midscale Train 7", 1.4286, date(2026, 7, 28), "20260728-3030", "feed gas"),
    ("Cameron", "Trains 1-3", "Train 1", 4.0, date(2019, 4, 5), "20190405-3083", "feed gas"),
    ("Cameron", "Trains 1-3", "Train 2", 4.0, date(2019, 11, 26), "20191126-3050", "feed gas"),
    ("Cameron", "Trains 1-3", "Train 3", 4.0, date(2020, 4, 21), "20200421-3085", "feed gas"),
    ("Freeport", "Trains 1-3", "Train 1", 5.0, date(2019, 7, 19), "20190719-3031", "hazardous fluids"),
    ("Freeport", "Trains 1-3", "Train 2", 5.0, date(2019, 11, 7), "20191107-3043", "hazardous fluids"),
    ("Freeport", "Trains 1-3", "Train 3", 5.0, date(2020, 3, 9), "20200309-3010", "hazardous fluids"),
    ("Elba Island", "10 MMLS units", "MMLS 1", 0.25, date(2019, 9, 30), "20190930-4000", "commence service"),
    ("Elba Island", "10 MMLS units", "MMLS 2", 0.25, date(2020, 1, 16), "20200116-3052", "commence service"),
    ("Elba Island", "10 MMLS units", "MMLS 3", 0.25, date(2019, 11, 26), "20191126-3018", "commence service"),
    ("Elba Island", "10 MMLS units", "MMLS 4", 0.25, date(2019, 12, 30), "20191230-3015", "commence service"),
    ("Elba Island", "10 MMLS units", "MMLS 5", 0.25, date(2020, 3, 2), "20200302-3018", "commence service"),
    ("Elba Island", "10 MMLS units", "MMLS 6", 0.25, date(2020, 2, 25), "20200225-3038", "hazardous fluids"),
    ("Elba Island", "10 MMLS units", "MMLS 7", 0.25, date(2020, 8, 5), "20200805-3007", "hazardous fluids"),
    ("Elba Island", "10 MMLS units", "MMLS 8", 0.25, date(2020, 5, 1), "20200501-3022", "feed gas"),
    ("Elba Island", "10 MMLS units", "MMLS 9", 0.25, date(2020, 6, 12), "20200612-3039", "feed gas"),
    ("Elba Island", "10 MMLS units", "MMLS 10", 0.25, date(2020, 7, 14), "20200714-3051", "hazardous fluids"),
    ("Calcasieu Pass", "9 blocks", "Block 1 (1A)", 0.5556, date(2022, 1, 12), "20220112-3006", "hazardous fluids"),
    ("Calcasieu Pass", "9 blocks", "Block 1 (1B)", 0.5556, date(2022, 1, 21), "20220121-3038", "hazardous fluids"),
    ("Calcasieu Pass", "9 blocks", "Block 2", 1.1111, date(2022, 1, 27), "20220127-3066", "hazardous fluids"),
    ("Calcasieu Pass", "9 blocks", "Block 3", 1.1111, date(2022, 2, 14), "20220214-3028", "hazardous fluids"),
    ("Calcasieu Pass", "9 blocks", "Block 4", 1.1111, date(2022, 2, 28), "20220228-3070", "hazardous fluids"),
    ("Calcasieu Pass", "9 blocks", "Block 5", 1.1111, date(2022, 4, 6), "20220406-3027", "hazardous fluids"),
    ("Calcasieu Pass", "9 blocks", "Block 6", 1.1111, date(2022, 4, 20), "20220420-3060", "hazardous fluids"),
    ("Calcasieu Pass", "9 blocks", "Block 7", 1.1111, date(2022, 5, 25), "20220525-3027", "hazardous fluids"),
    ("Calcasieu Pass", "9 blocks", "Block 8", 1.1111, date(2022, 6, 10), "20220610-3028", "hazardous fluids"),
    ("Calcasieu Pass", "9 blocks", "Block 9", 1.1111, date(2022, 7, 22), "20220722-3003", "hazardous fluids"),
    ("Plaquemines", "Phase 1", "Block 1", 1.4778, date(2024, 11, 21), "20241121-3079", "hazardous fluids"),
    ("Plaquemines", "Phase 1", "Block 2", 1.4778, date(2024, 12, 5), "20241205-3051", "hazardous fluids"),
    ("Plaquemines", "Phase 1", "Block 3", 1.4778, date(2024, 12, 17), "20241217-3076", "hazardous fluids"),
    ("Plaquemines", "Phase 1", "Block 4", 1.4778, date(2024, 12, 23), "20241223-3090", "hazardous fluids"),
    ("Plaquemines", "Phase 1", "Block 5", 1.4778, date(2025, 1, 10), "20250110-3008", "hazardous fluids"),
    ("Plaquemines", "Phase 1", "Block 6", 1.4778, date(2025, 1, 10), "20250110-3008", "hazardous fluids"),
    ("Plaquemines", "Phase 1", "Block 7", 1.4778, date(2025, 1, 28), "20250128-3019", "hazardous fluids"),
    ("Plaquemines", "Phase 1", "Block 8", 1.4778, date(2025, 2, 10), "20250210-3044", "hazardous fluids"),
    ("Plaquemines", "Phase 1", "Block 9", 1.4778, date(2025, 2, 27), "20250227-3064", "hazardous fluids"),
    ("Plaquemines", "Phase 2", "Block 10", 0.7444, date(2025, 4, 23), "20250423-3066", "hazardous fluids"),
    ("Plaquemines", "Phase 2", "Block 11", 0.7444, date(2025, 4, 24), "20250424-3027", "hazardous fluids"),
    ("Plaquemines", "Phase 2", "Block 12", 0.7444, date(2025, 6, 5), "20250605-3017", "hazardous fluids"),
    ("Plaquemines", "Phase 2", "Block 13", 0.7444, date(2025, 5, 8), "20250508-3059", "hazardous fluids"),
    ("Plaquemines", "Phase 2", "Block 14", 0.7444, date(2025, 7, 16), "20250716-3014", "hazardous fluids"),
    ("Plaquemines", "Phase 2", "Block 15", 0.7444, date(2025, 9, 18), "20250918-3052", "hazardous fluids"),
    ("Plaquemines", "Phase 2", "Block 16", 0.7444, date(2025, 8, 12), "20250812-3010", "hazardous fluids"),
    ("Plaquemines", "Phase 2", "Block 17", 0.7444, date(2025, 10, 21), "20251021-3029", "hazardous fluids"),
    ("Plaquemines", "Phase 2", "Block 18", 0.7444, date(2025, 8, 27), "20250827-3055", "hazardous fluids"),
    ("Golden Pass", "Trains 1-3", "Train 1", 6.0333, date(2025, 11, 1), "", "first LNG (no FERC train letter found)"),
]
for _plant, _, _, _, _start, _, _ in TRAIN_STARTS:
    if _start:
        PLANT_START[_plant] = min(PLANT_START.get(_plant, _start), _start)

# Expected start-ups over the next two years: trains under construction,
# dated from each developer's latest public guidance (first LNG where
# given, else a month inside the guided window - 'basis' says which).
# These are forecasts, not FERC letters: when a train gets its feed-gas
# letter, move it to TRAIN_STARTS with the real date. Existing plants'
# new trains stack in their own band; new plants share one band.
FUTURE_HORIZON = date(2028, 9, 30)
FUTURE_AS_OF = "Sep-2026"
FUTURE_TRAINS = [  # plant, phase, train/block, nameplate mtpa, expected first LNG, basis
    ("Golden Pass", "Trains 1-3", "Train 2", 6.0333, date(2026, 11, 1),
     "ExxonMobil: start-up fall 2026; FERC cleared Train 2 commissioning Jun-2026"),
    ("Golden Pass", "Trains 1-3", "Train 3", 6.0333, date(2027, 7, 1),
     "mechanical completion targeted 2Q 2027; EIA has it exporting in 2027"),
    ("Rio Grande", "Phase 1 (Trains 1-3)", "Train 1", 5.8667, date(2027, 4, 1),
     "NextDecade: first LNG 1H 2027 (commissioning from late 2026)"),
    ("Rio Grande", "Phase 1 (Trains 1-3)", "Train 2", 5.8667, date(2027, 10, 1),
     "EIA: Trains 1-2 exporting in 2027 - month estimated"),
    ("Rio Grande", "Phase 1 (Trains 1-3)", "Train 3", 5.8667, date(2028, 7, 1),
     "estimate: ~9 months after Train 2 (tracking ahead of guaranteed dates)"),
    ("Port Arthur", "Phase 1 (Trains 1-2)", "Train 1", 6.5, date(2027, 6, 1),
     "Sempra: Train 1 COD 2027; EIA: Phase 1 exporting in 2027 - month estimated"),
    ("Port Arthur", "Phase 1 (Trains 1-2)", "Train 2", 6.5, date(2028, 1, 1),
     "Sempra: Train 2 2028 - month estimated"),
    *[("CP2", "Phase 1 (13 blocks)", f"Block {i + 1}", round(14.4 / 13, 4),
       (pd.Timestamp(2027, 11, 1) + pd.DateOffset(months=i)).date(),
       "Venture Global: first LNG late 2027 (Phase 1 nameplate 14.4 mtpa); "
       "then a block a month, as Plaquemines ramped - estimate")
      for i in range(13)],
    ("Corpus Christi", "Midscale Trains 8-9", "Midscale Train 8", 1.5, date(2028, 6, 1),
     "Cheniere: substantial completion 2H 2028 (>3 mtpa for both) - month estimated"),
    ("Corpus Christi", "Midscale Trains 8-9", "Midscale Train 9", 1.5, date(2028, 9, 1),
     "Cheniere: substantial completion 2H 2028 - month estimated"),
]
FUTURE_TRAINS = [t for t in FUTURE_TRAINS if t[4] <= FUTURE_HORIZON]
BCF_PER_MT_LNG = 48.0  # 1 tonne of LNG ~ 48 Mcf of gas, so 1 mtpa ~ 0.13 Bcf/d


def capacity_table():
    rows = []
    for plant, phase, trains, mtpa, fuel_gas, accession, first in PLANT_CAPACITY:
        export = mtpa * BCF_PER_MT_LNG / 365
        rows.append({"plant": plant, "phase": phase, "trains": trains, "nameplate_mtpa": mtpa,
                     "nameplate_bcfd_lng_out": round(export, 2),
                     "nameplate_bcfd_feedgas": round(export * FEEDGAS_PER_EXPORT, 2),
                     "ferc_fuel_gas_authorised": pd.Timestamp(fuel_gas), "ferc_accession": accession,
                     "first_lng_start_up": pd.Timestamp(first)})
    df = pd.DataFrame(rows)
    totals = df.groupby("plant", sort=False)[["nameplate_mtpa", "nameplate_bcfd_lng_out", "nameplate_bcfd_feedgas"]].sum()
    total_rows = [{"plant": pl, "phase": "Plant total", "trains": None, **totals.loc[pl].round(2).to_dict(),
                   "ferc_fuel_gas_authorised": df.loc[df.plant == pl, "ferc_fuel_gas_authorised"].min(),
                   "first_lng_start_up": df.loc[df.plant == pl, "first_lng_start_up"].min()}
                  for pl in totals.index if (df.plant == pl).sum() > 1]  # only plants built in phases
    us = {"plant": "US total", "phase": "", "trains": None,
          **df[["nameplate_mtpa", "nameplate_bcfd_lng_out", "nameplate_bcfd_feedgas"]].sum().round(2).to_dict()}
    out = pd.concat([df, pd.DataFrame(total_rows), pd.DataFrame([us])], ignore_index=True)
    order = {pl: i for i, pl in enumerate(dict.fromkeys(df.plant))}
    out["_o"] = out.plant.map(order).fillna(len(order))
    out["_t"] = (out.phase == "Plant total").astype(int)
    return out.sort_values(["_o", "_t"], kind="stable").drop(columns=["_o", "_t"]).set_index("plant")
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
    "Sabine Pass": "Near complete - delivery meters of Creole Trail, Kinder Morgan Louisiana, NGPL and Transco's Gulf Trace lateral (Lighthouse Road).",
    "Plaquemines": "Complete - Gator Express's own delivery meter into the plant (its Tennessee Gas, Texas Eastern and Columbia Gulf receipts), with full history.",
    "Cameron": "Complete - Cameron Interstate's delivery meter into the terminal plus Columbia Gulf's direct feed.",
    "Calcasieu Pass": "Mostly complete - ANR (Grand Chenier XPress) and Texas Eastern deliveries into TransCameron; Sabine Pipe Line's, if any, not yet seen.",
    "Corpus Christi": "Mostly complete - Cheniere Corpus Christi Pipeline's delivery into the plant, which includes intrastate Permian gas it receives; gas delivered straight to the plant by the intrastate ADCC pipeline is not seen.",
    "Freeport": "Mostly metered - all three interstate feeds (Gulf South's Stratton Ridge delivery to Freeport LNG, Texas Eastern at Stratton Ridge and into BIG Pipeline); Texas intrastate supply is invisible, so use the calibrated estimate. Texas Eastern's Stratton Ridge meter also serves Dow's Freeport complex.",
    "Elba Island": "Elba Liquefaction meter on Elba Express.",
    "Golden Pass": "Golden Pass Pipeline's delivery meter into the terminal; gas brought straight to the plant by Kinder Morgan's intrastate Trident line, if any, is not seen.",
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


TRANSCO_API = "https://www.1line.williams.com/portal-service/api/portal-service/v1/public/oac/"
TRANSCO_CYCLES = [(8, "ID3"), (5, "Post"), (4, "ID2")]  # latest intraday first


def fetch_transco(page, pipe, gas_day):
    """Williams 1Line (Transco) OA report: a public API - POST the query,
    then GET the report view (an HTML table) for the gas day and cycle."""
    day = f"{gas_day:%m/%d/%Y}"
    req = page.context.request
    for cycle, label in TRANSCO_CYCLES:
        q = req.post(TRANSCO_API + "query", timeout=120000, data={
            "buid": 80, "mapId": 0, "startDate": day, "endDate": day, "cycle": cycle, "locationIds": "", "locationType": "All"})
        if not (q.ok and (q.json().get("data") or {}).get("recordCount")):
            continue
        html = req.get(TRANSCO_API + "report/view", timeout=180000, params={
            "buid": 80, "startDate": day, "endDate": day, "cycle": cycle, "zoneId": 0, "sortType": "HDR"}).text()
        frame = parse_transco_report(html, gas_day)
        print(f"  Transco: {len(frame)} delivery points, gas day {gas_day}, cycle {label}", flush=True)
        return frame
    raise RuntimeError(f"no Transco posting for {gas_day}")


def parse_transco_report(html, gas_day):
    m = re.search(r"Effective Gas Day:\s*(?:<[^>]+>\s*)*(\d\d/\d\d/\d{4})", html)
    if not m or m.group(1) != f"{gas_day:%m/%d/%Y}":
        raise RuntimeError(f"Transco report is for {m.group(1) if m else '?'}, wanted {gas_day}")
    rows = []
    for r in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S):
        cells = [" ".join(re.sub(r"<[^>]+>", " ", c).split()) for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", r, re.S)]
        if len(cells) >= 9 and cells[1] == "Delivery Location":
            rows.append((cells[0], cells[4], cells[8]))
    if not rows:
        raise RuntimeError("no delivery rows in the Transco report")
    df = pd.DataFrame(rows, columns=["loc", "name", "tsq"])
    return points_frame(df["loc"], df["name"], df["tsq"])


QUORUM_API = "https://web-prd.myquorumcloud.com/{site}/OpAvailPosting/GetOpAvailPostings?tspno={tsp}"


def fetch_quorum(page, pipe, gas_day):
    """Quorum myquorumcloud informational postings (Venture Global's
    TransCameron and Gator Express): the OA page's own JSON call, for the
    gas day, keeping the latest cycle posted. pipe = "<site>:<tspno>"."""
    site, tsp = pipe.split(":")
    r = page.context.request.post(QUORUM_API.format(site=site, tsp=tsp), timeout=120000, form={
        "sort": "", "group": "", "filter": "", "CycleId": "", "GasDay": f"{gas_day:%Y-%m-%d}", "LocId": "", "LocPurpCode": ""})
    rows = [d for d in (r.json().get("Data") or []) if str(d.get("GasDay", ""))[:10] == f"{gas_day:%Y-%m-%d}"]
    if not rows:
        raise RuntimeError(f"no Quorum {pipe} posting for {gas_day}")
    last = max(d["CycleId"] for d in rows)
    rows = [d for d in rows if d["CycleId"] == last]
    df = pd.DataFrame(rows)
    print(f"  Quorum {df['TspNm'].iloc[0]}: {len(df)} points, gas day {gas_day}, cycle {df['CycleDescr'].iloc[0]}", flush=True)
    return points_frame(df["LocId"], df["LocNm"], df["TotalSchdQty"])


GULFSOUTH_LIST = "https://reporting.prod.bwpmlp.org/infopost/infopostdetails"
GULFSOUTH_DOC = "https://reporting.prod.bwpmlp.org/infopost/postings?postingsDocumentId={doc}"
_GULFSOUTH_POSTINGS = []  # Operational Capacity postings, newest first, paged in as needed


def fetch_gulfsouth(page, pipe, gas_day):
    """Boardwalk GasQuest (Gulf South): the Operational Capacity posting
    list is a JSON call; each posting links a CSV. Take the gas day's
    Intraday 3 (final) posting, else its latest."""
    req = page.context.request
    wanted = f"{gas_day:%m/%d/%Y}"

    def day_postings():
        return [p for p in _GULFSOUTH_POSTINGS if p.get("description", "").startswith(wanted)]
    while not day_postings():
        oldest = _GULFSOUTH_POSTINGS[-1]["description"][:10] if _GULFSOUTH_POSTINGS else None
        if oldest and datetime.strptime(oldest, "%m/%d/%Y").date() < gas_day:
            break
        r = req.post(GULFSOUTH_LIST, timeout=60000, data={
            "infoPostID": 1, "tspId": 1, "pageNumber": len(_GULFSOUTH_POSTINGS) // 100 + 1, "pageSize": 100,
            "sortBy": "datetimePostingEffective", "groupCode": "INFOPOST", "sortDescending": True})
        got = (r.json() or {}).get("postings") or []
        if not got:
            break
        _GULFSOUTH_POSTINGS.extend(got)
    posts = day_postings()
    if not posts:
        raise RuntimeError(f"no Gulf South posting for {gas_day}")
    post = next((p for p in posts if "Intraday 3" in p["description"]), posts[0])
    doc = next(f["infoPostTrackerID"] for f in post["reportFiles"] if f["fileName"].lower().endswith(".csv"))
    text = req.get(GULFSOUTH_DOC.format(doc=doc), timeout=120000, headers={
        "Origin": "https://www.gasquest.com", "Referer": "https://www.gasquest.com/", "Accept": "text/csv,*/*"}).text().lstrip("\ufeff")
    if not text.startswith("TSP Name"):
        raise RuntimeError(f"Gulf South doc {doc} ({post['description']}) isn't the CSV: {text[:300]!r}")
    df = pd.read_csv(io.StringIO(text), dtype=str)
    df = df[df["Loc Purp Desc"].astype(str).str.startswith("Delivery")]
    print(f"  Gulf South: {len(df)} delivery points, {post['description']}", flush=True)
    return points_frame(df["Loc"], df["Loc Name"], df["Total Scheduled Quantity"])


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
            "tc": fetch_tc, "bhe": fetch_bhe, "transco": fetch_transco, "quorum": fetch_quorum,
            "gulfsouth": fetch_gulfsouth}


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


def capture_tc_only(out):
    """--tc-only: save TC's latest posting (ANR, Columbia Gulf) into the
    workbook's TC cache sheet and change nothing else. Run a few times a day
    so a missed morning run doesn't lose TC's day (TC keeps no history) and
    a later cycle for the same gas day overwrites an earlier one."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(ignore_https_errors=True, viewport={"width": 1600, "height": 1200})
        page = context.new_page()
        for pipe in TC_REPORTS:
            try:
                posted, frame = fetch_tc_latest(page, pipe)
            except Exception as e:
                print(f"  WARNING: TC {pipe} failed ({type(e).__name__}: {str(e)[:300]})", file=sys.stderr, flush=True)
                continue
            for point in POINTS:
                if point[1] == "tc" and point[2] == pipe and point[3] in frame.index:
                    TC_AHEAD.setdefault(posted, {})[point_column(point)] = frame.at[point[3], "scheduled_dth"]
        browser.close()
    if not TC_AHEAD:
        print("No TC posting captured - workbook unchanged.", flush=True)
        return
    import openpyxl
    cache = load_tc_cache(out)
    ahead = pd.DataFrame.from_dict(TC_AHEAD, orient="index")
    cache = ahead if cache is None else ahead.combine_first(cache)  # newest capture wins for the same day
    cache = cache.sort_index().tail(14)
    wb = openpyxl.load_workbook(out)
    position = wb.sheetnames.index(TC_CACHE_SHEET) if TC_CACHE_SHEET in wb.sheetnames else len(wb.sheetnames)
    if TC_CACHE_SHEET in wb.sheetnames:
        del wb[TC_CACHE_SHEET]
    ws = wb.create_sheet(TC_CACHE_SHEET, position)
    ws.append(["gas_day"] + list(cache.columns))
    for day, row in cache.iterrows():
        ws.append([pd.Timestamp(day).to_pydatetime()] + [None if pd.isna(v) else float(v) for v in row])
    for (cell,) in ws.iter_rows(min_row=2, max_col=1):
        cell.number_format = "dd-mmm-yyyy"
    ws.column_dimensions["A"].width = 14
    wb.save(out)
    print(f"TC cache updated for gas day(s) {sorted(TC_AHEAD)} - rest of the workbook untouched.", flush=True)


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


def pull(gas_day, tc_cache=None, skip=()):
    """One row of point-level scheduled Dth for the gas day. Points on a
    pipeline that failed to load (or is in skip) are left blank, not zero;
    row.attrs["failed"] lists the pipelines that were tried and failed."""
    needed = [x for x in sorted({(p[1], p[2]) for p in POINTS}) if x not in skip]
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
    df = pd.DataFrame([row], index=pd.Index([gas_day], name="gas_day"))
    df.attrs["failed"] = set(needed) - set(results)
    return df


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


def monthly_check(plants_daily, eia, points=None):
    """Metered monthly average vs EIA-implied feedgas, per plant. Only
    days with every one of the plant's meters present count - a share
    worked out while a meter was missing (e.g. TC's, which have no
    history) would over-scale the plant once that meter is back."""
    daily = plants_daily.drop(columns=["Total"], errors="ignore").copy()
    if points is not None:
        for plant in daily.columns:
            cols = [point_column(p) for p in POINTS if p[0] == plant and point_column(p) in points.columns]
            if cols:
                daily[plant] = daily[plant].where(points[cols].notna().all(axis=1))
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


HISTORY_START = date(2020, 1, 1)
NO_DATA = "no data yet (EIA runs 2-3 months behind)"


def best_estimate_daily(est, eia, points=None):
    """One daily series per plant from HISTORY_START: the metered and
    calibrated estimate where the daily pull has it, otherwise the plant's
    EIA monthly feedgas spread ratably across every day of that month (the
    pipelines only keep ~2 years of daily postings, so EIA is the history).
    'Source' says which: metered, EIA monthly (ratable), or mixed."""
    if eia is None or eia.empty:
        return pd.DataFrame()
    plants = list(dict.fromkeys(p[0] for p in POINTS))
    eia_end = pd.Period(eia.index[-1], "M").end_time.date()
    last = max([eia_end] + ([max(est.index)] if not est.empty else []))
    days = pd.date_range(HISTORY_START, last, freq="D")
    months = days.to_period("M").astype(str)
    flat = pd.DataFrame({pl: (eia[pl].reindex(months).values if pl in eia.columns else float("nan")) for pl in plants},
                        index=days.date)
    metered = est.reindex(columns=plants).reindex(flat.index)
    # Provisional: months EIA hasn't published yet carry each plant's latest
    # EIA month forward - used only for plant days missing meters, and
    # replaced by the real month once EIA publishes it.
    carried = pd.DataFrame({pl: (eia[pl].dropna().iloc[-1] if pl in eia.columns and eia[pl].notna().any() else float("nan"))
                            for pl in plants}, index=flat.index)
    carried = carried.where(flat.isna() & (pd.Series(months, index=flat.index) > eia.index[-1]).values[:, None])
    provisional = pd.DataFrame(False, index=flat.index, columns=plants)
    if points is not None:
        # A plant day missing some of its meters (Kinder Morgan, Cheniere
        # and Cameron Interstate keep ~90 days; TC has no history)
        # understates the plant: use EIA's ratable figure for that plant
        # instead - the month's own where EIA has it, else the latest
        # EIA month carried forward (provisional).
        pts = points.reindex(flat.index)
        for pl in plants:
            if pl not in PLANT_START:
                continue
            live = pd.Series([d >= PLANT_START[pl] for d in flat.index], index=flat.index)
            cols = [point_column(p) for p in POINTS if p[0] == pl and point_column(p) in pts.columns]
            if cols:
                partial = pts[cols].isna().any(axis=1) & live
                use_month = partial & flat[pl].notna()
                use_carry = partial & flat[pl].isna() & carried[pl].notna()
                metered.loc[use_month | use_carry, pl] = float("nan")
                flat.loc[use_carry, pl] = carried.loc[use_carry, pl]
                provisional.loc[use_carry, pl] = True
    out = metered.combine_first(flat)
    used_met, used_eia = metered.notna().any(axis=1), (metered.isna() & flat.notna()).any(axis=1)
    used_prov = provisional.any(axis=1)
    source = pd.Series(NO_DATA, index=out.index)
    source[used_eia] = "EIA monthly (ratable)"
    source[used_met] = "mixed (EIA ratable for unmetered plants)"
    source[used_met & used_prov] = "mixed (provisional: latest EIA month for unmetered plants)"
    source[used_met & ~used_eia] = "metered"
    out["Total"] = out[plants].sum(axis=1, min_count=1)
    out["Source"] = source
    out.index.name = "gas_day"
    return out.round(3)

# 'Feedgas vs capacity' chart: nameplate capacity stacked by plant in
# start-up order (Cove Point and Elba share a slot - eight series is the
# most the palette keeps distinguishable), total feedgas as a line.
CAPACITY_GROUPS = [
    ("Sabine Pass", ["Sabine Pass"]), ("Cove Point + Elba", ["Cove Point", "Elba Island"]),
    ("Corpus Christi", ["Corpus Christi"]), ("Cameron", ["Cameron"]), ("Freeport", ["Freeport"]),
    ("Calcasieu Pass", ["Calcasieu Pass"]), ("Plaquemines", ["Plaquemines"]), ("Golden Pass", ["Golden Pass"]),
    ("New plants (Rio Grande, Port Arthur, CP2)", ["Rio Grande", "Port Arthur", "CP2"]),
]
# a ninth hue wouldn't stay distinguishable - new plants get a neutral grey
CHART_COLORS = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "4A3AA7", "E34948", "9A9A94"]
CHART_LINE_COLOR = "0B0B0B"
FEEDGAS_LINE = "Total feedgas"


def train_table():
    df = pd.DataFrame(TRAIN_STARTS, columns=["plant", "phase", "train", "nameplate_mtpa", "start_up", "ferc_accession", "basis"])
    df["nameplate_bcfd_feedgas"] = (df["nameplate_mtpa"] * BCF_PER_MT_LNG / 365 * FEEDGAS_PER_EXPORT).round(3)
    df["start_up"] = pd.to_datetime(df["start_up"])
    return df.set_index("plant")[["phase", "train", "start_up", "nameplate_mtpa", "nameplate_bcfd_feedgas", "basis", "ferc_accession"]]


def future_table():
    df = pd.DataFrame(FUTURE_TRAINS, columns=["plant", "phase", "train", "nameplate_mtpa", "expected_first_lng", "basis"])
    df["nameplate_bcfd_feedgas"] = (df["nameplate_mtpa"] * BCF_PER_MT_LNG / 365 * FEEDGAS_PER_EXPORT).round(3)
    df["expected_first_lng"] = pd.to_datetime(df["expected_first_lng"])
    df["cumulative_added_bcfd"] = df.sort_values("expected_first_lng")["nameplate_bcfd_feedgas"].cumsum().round(2)
    df = df.sort_values("expected_first_lng")
    return df.set_index("plant")[["phase", "train", "expected_first_lng", "nameplate_mtpa", "nameplate_bcfd_feedgas",
                                  "cumulative_added_bcfd", "basis"]]


def chart_data(best):
    """Daily nameplate capacity (feedgas Bcf/d) per group, stepping up as
    each train starts - FERC-dated trains, then expected ones out to
    FUTURE_HORIZON - plus the best-estimate total (blank in the future)."""
    index = [d.date() for d in pd.date_range(min(best.index), max(max(best.index), FUTURE_HORIZON))]
    days = pd.Series(pd.to_datetime(index), index=index)
    trains = [(t[0], t[3], t[4]) for t in TRAIN_STARTS] + [(t[0], t[3], t[4]) for t in FUTURE_TRAINS]
    out = {}
    for group, plants in CAPACITY_GROUPS:
        cap = pd.Series(0.0, index=index)
        for plant, mtpa, start in trains:
            if plant in plants and start:
                cap += (days >= pd.Timestamp(start)) * mtpa * BCF_PER_MT_LNG / 365 * FEEDGAS_PER_EXPORT
        out[group] = cap.round(3)  # column header = legend label (plant name only)
    df = pd.DataFrame(out, index=index)
    df[FEEDGAS_LINE] = best["Total"].reindex(index)
    df.index.name = "gas_day"
    return df


CHART_FONT_PT = 12
CHART_TITLE_PT = 16
CHART_FONT = "Aptos"            # same faces as the repo's other native charts
CHART_TITLE_FONT = "Aptos Display"


def _font(size_pt=CHART_FONT_PT, bold=False, face=CHART_FONT):
    """Text properties for a chart element: one font and size throughout."""
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import CharacterProperties, Font, Paragraph, ParagraphProperties
    cp = CharacterProperties(sz=int(size_pt * 100), b=bold, latin=Font(typeface=face))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def _title(text, size_pt=CHART_FONT_PT, face=CHART_FONT):
    from openpyxl.chart.text import Text
    from openpyxl.chart.title import Title
    from openpyxl.drawing.text import CharacterProperties, Font, Paragraph, ParagraphProperties, RegularTextRun
    cp = CharacterProperties(sz=int(size_pt * 100), b=True, latin=Font(typeface=face))
    rich = Text(rich=_font(size_pt, bold=True, face=face))
    rich.rich.p = [Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])]
    return Title(tx=rich, overlay=False)


def add_capacity_chart(wb, data_sheet="Chart data", title="Feedgas vs capacity"):
    """Excel Quick Layout 1 style (chart title + legend), legend at the
    bottom, the units as the primary y-axis title, every piece of chart
    text at CHART_FONT_PT."""
    from openpyxl.chart import AreaChart, LineChart, Reference
    from openpyxl.chart.axis import DateAxis
    ws = wb[data_sheet]
    n = ws.max_row
    area = AreaChart()
    area.grouping = "stacked"
    area.title = _title("US LNG feedgas vs nameplate capacity", CHART_TITLE_PT, CHART_TITLE_FONT)
    area.y_axis.title = _title("Gcf/d")  # primary y-axis label carries the units (Gcf/d = Bcf/d)
    area.y_axis.majorGridlines.spPr = None
    area.add_data(Reference(ws, min_col=2, max_col=1 + len(CAPACITY_GROUPS), min_row=1, max_row=n), titles_from_data=True)
    dates = Reference(ws, min_col=1, min_row=2, max_row=n)
    area.set_categories(dates)
    for series, color in zip(area.series, CHART_COLORS):
        series.graphicalProperties.solidFill = color
        series.graphicalProperties.line.solidFill = "FFFFFF"
        series.graphicalProperties.line.width = 6350  # 0.5pt surface gap between bands
    line = LineChart()
    line.add_data(Reference(ws, min_col=2 + len(CAPACITY_GROUPS), min_row=1, max_row=n), titles_from_data=True)
    feed = line.series[0]
    feed.graphicalProperties.line.solidFill = CHART_LINE_COLOR
    feed.graphicalProperties.line.width = 22860  # 1.8pt
    feed.smooth = False
    line.y_axis.axId = area.y_axis.axId  # one shared axis
    area += line
    area.x_axis = DateAxis(crossAx=100)
    area.x_axis.number_format = "mmm-yy"
    area.x_axis.majorTimeUnit = "months"
    area.x_axis.majorUnit = 6
    area.display_blanks = "gap"  # no feedgas line in the future
    area.x_axis.title = None
    area.legend.position = "b"
    area.legend.overlay = False
    area.legend.txPr = _font()
    area.x_axis.txPr = _font()
    area.y_axis.txPr = _font()
    area.x_axis.delete = False
    area.y_axis.delete = False
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    area.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))  # no border round the chart
    area.width, area.height = 40, 22
    cs = wb.create_sheet(title, 1)
    cs.add_chart(area, "A1")
    cs["A46"] = ("Stacked areas: operator nameplate liquefaction capacity by plant, as feedgas Bcf/d, stepping up train by train "
                 "on each train's FERC feed-gas date ('Train start-ups' tab). "
                 "Line: 'Best estimate daily' total - EIA monthly spread ratably before the daily pull, metered after. "
                 "Plants routinely run above nameplate, so the line can sit above the stack. "
                 f"Beyond the line's end the stack is EXPECTED capacity to {FUTURE_HORIZON:%b-%Y}: trains under construction, "
                 f"on developer guidance as of {FUTURE_AS_OF} ('Future capacity' tab) - forecasts, not FERC dates.")


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
    lines += ["", "PLANT START DATES (first train's FERC feed-gas letter - see Train start-ups; blanks before these = not operating yet)"]
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
              "", "BEST ESTIMATE DAILY",
              f"'Best estimate daily': one row per gas day from {HISTORY_START:%d %b %Y}. Where the daily pipeline pull has the day, "
              "the metered/calibrated estimate ('Estimated (calibrated)'); before that - pipelines only keep about two years of "
              "daily postings - each plant's EIA monthly feedgas spread ratably across the month, so history has no day-to-day "
              "shape. 'Source' says which was used. A plant missing some of its meters on a day (Kinder Morgan keeps only "
              "~90 days of postings; TC has none) uses EIA's ratable figure for that month where EIA has it ('mixed'), or, "
              "for months EIA hasn't published yet, its latest EIA month carried forward ('provisional') until EIA catches up.",
              "", "PLANT CAPACITY",
              "'Plant capacity': operator-published nameplate liquefaction capacity by phase (mtpa of LNG) and first-LNG month. "
              f"Bcf/d = mtpa x {BCF_PER_MT_LNG:.0f} Bcf per million tonnes / 365 ('lng_out'); 'feedgas' grosses that up by "
              f"{FEEDGAS_PER_EXPORT:.2f} like the EIA calibration. Plants often run above nameplate, so metered feedgas above it is normal.",
              f"'Future capacity': trains under construction expected to start by {FUTURE_HORIZON:%b-%Y}, dated from developer "
              f"guidance as of {FUTURE_AS_OF} ('basis' says whether the month is guided or estimated). The chart's stack "
              "continues past today with these - they are forecasts and slip; each moves to 'Train start-ups' once FERC "
              "letters it.",
              "", "TC ENERGY METERS (ANR, COLUMBIA GULF)",
              "TC eConnects only serves its latest posting - by each morning run, the next gas day's Timely cycle. "
              f"Each run caches it ('{TC_CACHE_SHEET}') and the next run uses it for that day, so TC's meters "
              "(Calcasieu Pass via ANR; Cameron and Plaquemines via Columbia Gulf) are Timely nominations, not the final "
              "cycle, and have no history before the cache started. A missed run leaves that day's TC meters blank.",
              "", "POINTS", "Each 'Points (Dth)' column is 'Plant | meter (pipeline code, location id)'.", "",
              "SOURCE", "Kinder Morgan (pipeline2.kindermorgan.com) and Enbridge LINK (rtba.enbridge.com) "
              "Operationally Available Capacity postings, Total Scheduled Quantity column; gasnom.com (Cameron Interstate), "
              "Energy Transfer Messenger+, Cheniere LNG Connection, TC eConnects, BHE GT&S, Williams 1Line (Transco), "
              "Venture Global's pipelines (Quorum) and Boardwalk GasQuest (Gulf South) equivalents - all operator-hosted."]
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
    parser.add_argument("--only", help="comma-separated platforms (e.g. transco) - pull just these, leaving the "
                                       "workbook's other meters as they are (to backfill a newly added meter)")
    parser.add_argument("--tc-only", action="store_true", help="just save TC's latest posting into the TC cache "
                                                                 "sheet (the extra intraday runs); nothing else changes")
    args = parser.parse_args()
    if args.tc_only:
        capture_tc_only(args.out)
        return
    yesterday = datetime.now(ZoneInfo("America/Chicago")).date() - timedelta(days=1)
    gas_day = date.fromisoformat(args.gas_day) if args.gas_day else yesterday
    if args.dump:
        dump(gas_day)
        return
    only = {x.strip() for x in args.only.split(",")} if args.only else None
    not_wanted = {(p[1], p[2]) for p in POINTS if only and p[1] not in only}
    if args.probe:
        probe([date.fromisoformat(d.strip()) for d in args.probe.split(",")], not_wanted)
        return

    if args.date_from:
        # Newest first: pipelines keep a limited history, so once one has
        # failed GIVE_UP_AFTER days running it won't have older days either
        # - stop asking (a missing day can cost minutes in timeouts).
        first, last = date.fromisoformat(args.date_from), date.fromisoformat(args.date_to) if args.date_to else yesterday
        days = [last - timedelta(days=i) for i in range((last - first).days + 1)]
    else:
        days = [gas_day]
    GIVE_UP_AFTER = 3
    streak, dead = {}, set(not_wanted)

    # Backfills save every SAVE_EVERY days, so an interrupted run keeps
    # what it has (re-running skips nothing - it just overwrites).
    SAVE_EVERY = int(os.environ.get("FEEDGAS_SAVE_EVERY", "7"))
    pending, retrieved = [], 0
    tc_cache = load_tc_cache(args.out)
    started = datetime.now()
    progress_log = os.path.splitext(args.out)[0] + "_progress.log"
    for i, day in enumerate(days, 1):
        print(f"Pulling LNG feedgas points for gas day {day} ({i}/{len(days)})...", flush=True)
        row = pull(day, tc_cache, dead)
        for pipe in {(p[1], p[2]) for p in POINTS} - dead:
            streak[pipe] = streak.get(pipe, 0) + 1 if pipe in row.attrs.get("failed", ()) else 0
            if len(days) > 1 and streak[pipe] >= GIVE_UP_AFTER:
                dead.add(pipe)
                print(f"  {pipe[0]} {pipe[1]}: failed {GIVE_UP_AFTER} days running - not asking it for days before {day}", flush=True)
        if row.notna().sum(axis=1).iloc[0] == 0:
            print(f"  No points retrieved for {day} - skipped.", file=sys.stderr)
        else:
            pending.append(row)
            retrieved += 1
        line = progress_line(progress_log, i, len(days), day, row, started)
        if pending and (i % SAVE_EVERY == 0 or i == len(days)):
            save(args.out, pd.concat(pending))
            pending = []
        if os.environ.get("FEEDGAS_PUSH_EACH_DAY"):  # set by the workflow for a watched backfill
            push_progress([args.out, progress_log], line)
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
    return line


def push_progress(paths, line):
    """Commit and push the workbook and progress log with the day's
    progress line as the message, so a backfill running on GitHub can be
    followed day by day from the commit history. Best effort."""
    import subprocess
    run = lambda *cmd: subprocess.run(cmd, capture_output=True, text=True)
    run("git", "add", "-f", *paths)
    if run("git", "diff", "--cached", "--quiet").returncode == 0:
        return
    run("git", "commit", "-m", f"Feedgas backfill: {' '.join(line.split())}")
    for _ in range(3):
        run("git", "pull", "--rebase", "-X", "theirs")
        if run("git", "push").returncode == 0:
            return
    print("  (progress push failed - carrying on)", flush=True)


def probe(days, skip=()):
    """For each gas day, which pipelines return a posting - pipelines are
    only required to keep ~3 years of postings, so this shows how far
    back a backfill will actually find data. Nothing is saved."""
    needed = sorted({(p[1], p[2]) for p in POINTS} - set(skip))
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
    check = monthly_check(plants, eia, points) if not eia.empty else pd.DataFrame()
    est, share = estimated(plants, check, eia)
    if share:
        print(f"  seen shares used: {share}")

    best = best_estimate_daily(est, eia, points)
    sheets = {"Best estimate daily": best} if not best.empty else {}
    sheets.update({"Bcfd by plant": plants, "Estimated (calibrated)": est, "Plant capacity": capacity_table(),
                   "Train start-ups": train_table(), "Future capacity": future_table(), "Points (Dth)": points})
    if not best.empty:
        sheets["Chart data"] = chart_data(best)
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
                               "TC ENERGY METERS (ANR, COLUMBIA GULF)", "PLANT CAPACITY",
                               "BEST ESTIMATE DAILY",
                               "PLANT START DATES (first train's FERC feed-gas letter - see Train start-ups; blanks before these = not operating yet)"})
    import openpyxl
    wb = openpyxl.load_workbook(args.out)
    for name in ("Best estimate daily", "Bcfd by plant", "Estimated (calibrated)", "Points (Dth)", TC_CACHE_SHEET):
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        for (cell,) in ws.iter_rows(min_row=2, max_col=1):
            cell.number_format = "dd-mmm-yyyy"
        ws.column_dimensions["A"].width = 14
    if "Chart data" in wb.sheetnames:
        for (cell,) in wb["Chart data"].iter_rows(min_row=2, max_col=1):
            cell.number_format = "dd-mmm-yyyy"
        add_capacity_chart(wb)
    ws = wb["Train start-ups"]
    for row in ws.iter_rows(min_row=2):
        row[3].number_format = "dd-mmm-yyyy"
    for col, width in zip("ABCDEFGH", (16, 14, 20, 13, 15, 22, 44, 16)):
        ws.column_dimensions[col].width = width
    ws = wb["Future capacity"]
    for row in ws.iter_rows(min_row=2):
        row[3].number_format = "mmm-yyyy"
    for col, width in zip("ABCDEFGH", (16, 22, 18, 18, 15, 22, 21, 90)):
        ws.column_dimensions[col].width = width
    ws = wb["Plant capacity"]
    for row in ws.iter_rows(min_row=2):
        row[6].number_format = "dd-mmm-yyyy"
        row[8].number_format = "mmm-yyyy"
        if row[1].value == "Plant total" or row[0].value == "US total":
            for cell in row:
                cell.font = openpyxl.styles.Font(bold=True)
    for col, width in zip("ABCDEFGHI", (16, 30, 8, 16, 22, 22, 24, 16, 16)):
        ws.column_dimensions[col].width = width
    wb.save(args.out)

    print(f"\nSaved {args.out} - {len(points)} gas day(s).")
    print(plants.tail().to_string())


if __name__ == "__main__":
    main()
