"""
Africa master workbook: one file with every African dataset this repo pulls from a raw source (grid operator,
regulator, ministry, national oil company, statistics office - NO Ember), and Dashboard front pages carrying all
their charts. Same layout as the North America / South & Central America masters (south_america/SOUTH_AMERICA_MASTER.py,
whose table/chart/dashboard code this reuses).

  Dashboard          - gas and oil: Nigeria (NNPC Ltd monthly report)
  Dashboard - Power  - generation by source, then lake levels and capacity where a raw feed exists:
                       South Africa (Eskom), Ghana (Energy Commission: weekly WEM statistics, annual statistics with
                       Akosombo / Bui water-year level charts), Nigeria (NERC quarterly report), Cameroon (ARSEL
                       monthly energy balance, 2025 so far). Where raw data has gaps, Ember fills them as a labelled
                       fallback: every Ember series is named "... - Ember (fallback)" (tables, charts, Sources).
  <CC> <chart> data  - the table each Dashboard chart plots (CC = ZA, GH, NG ... : one tab group per country)
  <CC> <dataset> raw - the full data sheet(s) from each source workbook
  Sources            - where each dataset comes from, units and notes

Reads (doesn't refetch) the workbooks the scheduled pulls write to "output/Data and Chart Outputs/". A missing
input is listed on the Dashboard and skipped rather than stopping the rest.

TO ADD A COUNTRY: append its workbook to the lists below -
  DATASETS            gas / oil / other non-power workbooks   (Dashboard)
  RAW_POWER_DATASETS  generation workbooks                    (Dashboard - Power)
  OTHER_POWER_DATASETS  demand, water levels, gas burn, ...   (Dashboard - Power)
  CAPACITY_DATASETS   installed capacity workbooks            (Dashboard - Power)
each as (country code, country, workbook, raw sheet / tuple of sheets / "*", short name), add a SOURCES entry
(publisher + link), and register the workbook's chart specs in add_charts.py REGISTRY (or MASTER_SPECS here).

Usage: python3 AFRICA_MASTER.py [--out "output/Data and Chart Outputs/Master Outputs/africa_master.xlsx"]
"""
import argparse
import os
import sys

import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
import add_charts  # noqa: E402
import xlsx_charts  # noqa: E402
import SOUTH_AMERICA_MASTER as sam  # noqa: E402

DATA_DIR = sam.DATA_DIR

# (country code, country, workbook, raw sheet / tuple of raw sheets / "*", short dataset name)
DATASETS = [
    ("NG", "Nigeria", "nigeria_nnpc_gas_monthly.xlsx", "Monthly", "NNPC gas and oil"),
]
RAW_POWER_DATASETS = [
    ("ZA", "South Africa", "south_africa_generation_mix_daily.xlsx",
     ("Data", "System", "Weekly EAF", "Pumped storage", "Load shedding"), "power"),
    ("GH", "Ghana", "ghana_wem_weekly_generation_daily.xlsx", ("Daily", "Plants_GWh"), "WEM weekly power"),
    ("NG", "Nigeria", "nigeria_power_generation_quarterly.xlsx", ("Quarterly", "Plants"), "power"),
    ("CM", "Cameroon", "cameroon_arsel_energy_balance.xlsx",
     ("Plant injections", "Substation off-take", "HT customers"), "ARSEL energy balance"),
]
OTHER_POWER_DATASETS = [
    # annual generation by type plus Akosombo / Bui water-year level charts
    ("GH", "Ghana", "ghana_energy_commission_power.xlsx", ("Annual", "Akosombo", "Bui"), "annual statistics and lake levels"),
    # EMBER (FALLBACK) gap fills: only where the raw feed above has a hole (see ember_* below)
    ("NG", "Nigeria", "west_africa_power_by_type.xlsx", ("Nigeria",), "Ember fallback monthly"),
    ("NG", "Nigeria", "west_africa_power_by_type_annual.xlsx", ("Nigeria",), "Ember fallback annual"),
    ("CM", "Cameroon", "west_africa_power_by_type_annual.xlsx", ("Cameroon",), "Ember fallback annual"),
]
CAPACITY_DATASETS = []
# Countries with NO raw feed yet: Ember (fallback) annual data only, on their own dashboard. Replace an entry with raw
# datasets in the lists above when a raw source is found.
EMBER_ONLY_DATASETS = [(code, c, "west_africa_power_by_type_annual.xlsx", (c,), "Ember fallback annual") for code, c in [
    ("SN", "Senegal"), ("CI", "Cote d'Ivoire"), ("MR", "Mauritania"), ("ML", "Mali"), ("BF", "Burkina Faso"),
    ("GN", "Guinea"), ("SL", "Sierra Leone"), ("LR", "Liberia"), ("BJ", "Benin"), ("TG", "Togo"), ("NE", "Niger"),
    ("GM", "Gambia"), ("GW", "Guinea-Bissau"), ("CV", "Cape Verde"), ("GQ", "Equatorial Guinea")]]
HYDRO_DATASETS = []
HYDRO_EXTRA = {}
DASHBOARD_ONLY = {}
EMBER_TAG = "Ember (fallback)"
EMBER = {"west_africa_power_by_type.xlsx", "west_africa_power_by_type_annual.xlsx"}
OPERATORS = {}
GAS_BCFD = True
CTX = {}   # the country whose Ember chart is being built (main() sets it before each collect call)

# Cameroon ARSEL plants -> fuel group (plant injections into the grid)
CM_HYDRO = ["Songloulou", "Edea 3", "Lagdo", "Nachtigal", "Memve'ele", "Lom Pangar"]
CM_GAS = ["Kribi (gas)"]
CM_OIL = ["Limbe HFO", "Oyomabang", "Dibamba"]


def _ember_annual(d):
    """Ember yearly sheet -> Hydro / Gas / Other (incl. solar, wind, bio, other fossil) in GWh, named as Ember fallback."""
    d = d.set_index(pd.to_datetime(d["Year"].astype(int).astype(str) + "-01-01")).drop(columns="Year")
    g = add_charts.power_mix(d)
    g = g.loc[:, (g.fillna(0) != 0).any(axis=0)]
    return g.rename(columns=lambda c: f"{c} - {EMBER_TAG}")


def ember_annual(path):
    """Annual generation by type from Ember (FALLBACK) for the country in CTX. Cameroon: the raw ARSEL months
    (2025 so far) are drawn as separate, differently named columns next to Ember's earlier years."""
    country = CTX["country"]
    g = _ember_annual(add_charts.read(path, country))
    only = country in {d[1] for d in EMBER_ONLY_DATASETS}
    title = (f"{country} power generation by type: Ember (fallback, annual) - no raw source found yet" if only else
             f"{country} power generation by type, annual: {EMBER_TAG} (Ember; no raw annual feed)")
    if country == "Cameroon":
        arsel = os.path.join(os.path.dirname(path), "cameroon_arsel_energy_balance.xlsx")
        try:
            inj = add_charts.by_date(add_charts.read(arsel, "Plant injections"), "month").apply(pd.to_numeric, errors="coerce") / 1000
            y = inj.groupby(inj.index.year)
            full = y.count().min(axis=1) >= 12
            rows = {}
            for yr in full[full].index:
                part = inj[inj.index.year == yr]
                rows[pd.Timestamp(f"{yr}-01-01")] = {
                    "Hydro - ARSEL (raw)": part[[c for c in CM_HYDRO if c in part]].sum().sum(),
                    "Gas - ARSEL (raw)": part[[c for c in CM_GAS if c in part]].sum().sum(),
                    "Other (oil) - ARSEL (raw)": part[[c for c in CM_OIL if c in part]].sum().sum()}
            if rows:
                g = g.join(pd.DataFrame(rows).T, how="outer")
                title = ("Cameroon power generation by type, annual: Ember (fallback) to 2024, ARSEL (raw) plant injections "
                         "for 2025 (grid injections only - not like-for-like with Ember)")
        except Exception:  # noqa: BLE001 - ARSEL workbook missing: Ember only
            pass
    return [add_charts.spec("Ember annual", g.round(0), title, "GWh per year", "stacked_bar", "%Y")]


def ember_monthly(path):
    """Nigeria monthly gas / hydro generation from Ember (FALLBACK): NERC's raw feed is quarterly only."""
    d = add_charts.by_date(add_charts.read(path, "Nigeria"), "Month")
    g = d[["Hydro_GWh", "Gas_GWh"]].rename(columns=lambda c: f"{c.replace('_GWh', '')} - {EMBER_TAG}").dropna(how="all")
    return [add_charts.spec("Nigeria", g, "Nigeria monthly power generation by type: Ember (fallback) - NERC raw feed is quarterly only",
                            "GWh per month", "stacked_bar")]


MASTER_SPECS = {"west_africa_power_by_type_annual.xlsx": ember_annual,
                "west_africa_power_by_type.xlsx": ember_monthly}
DASHBOARD_ONLY = {}

SOURCES = {
    "south_africa_generation_mix_daily.xlsx": (
        "Eskom Data Portal, Station Build Up (hourly generation by station, rolling last 7 days; this repo archives it "
        "every 3 days, so history starts when the pull started)", "https://www.eskom.co.za/dataportal/"),
    "ghana_wem_weekly_generation_daily.xlsx": (
        "Energy Commission of Ghana, Weekly Wholesale Electricity Market (WEM) Statistics (market-operator dispatch data)",
        "https://www.energycom.gov.gh/index.php/planning/weekly-wholesale-electricity-market-wem-statistics"),
    "ghana_energy_commission_power.xlsx": (
        "Energy Commission of Ghana, National Energy Statistics (GRIDCo, VRA, ECG and IPP returns)",
        "https://www.energycom.gov.gh/index.php/planning/energy-statistics"),
    "nigeria_power_generation_quarterly.xlsx": (
        "NERC (Nigerian Electricity Regulatory Commission), Quarterly Reports: NISO-metered generation of grid-connected plants only "
        "(QUARTERLY average, 2019Q1-; embedded/captive generation not included; thermal booked to gas)",
        "https://nerc.gov.ng/resource-category/nerc-reports/"),
    "nigeria_nnpc_gas_monthly.xlsx": (
        "NNPC Ltd, Monthly Report Summary (gas production and sales, crude and condensate production; SHORT HISTORY, from Feb-2025)",
        "https://www.nnpcgroup.com/insights"),
    "cameroon_arsel_energy_balance.xlsx": (
        "ARSEL (Agence de Regulation du Secteur de l'Electricite, Cameroon) / SONATREL-ENEO monthly energy balance "
        "(plant injections, substation off-take, HT customers). Jan-Dec 2025 so far, published with an 8+ month lag",
        "https://arsel-cm.org/bilan-energetique-mensuel/"),
    "west_africa_power_by_type.xlsx": (
        "Ember (fallback) monthly electricity data - gap fill only where the raw feed is missing or coarser (Nigeria monthly)",
        "https://ember-energy.org/data/monthly-electricity-data/"),
    "west_africa_power_by_type_annual.xlsx": (
        "Ember (fallback) yearly electricity data - gap fill only where the raw feed is missing (Cameroon before the ARSEL months; "
        "Nigeria before NERC's 2019 start and for the latest year; and the Ember-only countries on 'Dashboard - Power (Ember)')",
        "https://ember-energy.org/data/yearly-electricity-data/"),
}

CAVEATS = [
    "Cameroon ARSEL data is raw but covers Jan-Dec 2025 only and is published with an 8+ month lag; Ember (fallback) fills earlier years",
    "Nigeria NERC data is QUARTERLY (grid-connected plants only, from 2019Q1); Ember (fallback) gives the monthly view",
    "Ember (fallback) series are labelled '- Ember (fallback)' and never replace raw data where raw exists",
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(DATA_DIR, "Master Outputs", "africa_master.xlsx"))
    ap.add_argument("--data-dir", default=DATA_DIR)
    args = ap.parse_args()
    cfg = sys.modules[__name__]

    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    dash2 = wb.create_sheet("Dashboard - Power")
    used = {"Dashboard", "Dashboard - Power", "Dashboard - Power (Ember)", "Sources"}
    sources = []

    gas = sam.collect(wb, DATASETS, args.data_dir, used, sources, cfg=cfg)
    # one block per country (dashboard order = order of first appearance in the lists), generation then levels/capacity
    pd_all = RAW_POWER_DATASETS + OTHER_POWER_DATASETS + CAPACITY_DATASETS
    order = {}
    for d in pd_all:
        order.setdefault(d[1], len(order))
    power = ([], [], [])
    for d in sorted(pd_all, key=lambda d: order[d[1]]):   # stable: keeps list order within a country
        CTX["country"] = d[1]
        part = sam.collect(wb, [d], args.data_dir, used, sources, cfg=cfg)
        for i in range(3):
            power[i].extend(part[i])
    power[2].extend(CAVEATS)

    # Ember-only countries: separate dashboard, every chart labelled Ember (fallback, annual)
    emb = ([], [], [])
    for d in EMBER_ONLY_DATASETS:
        CTX["country"] = d[1]
        part = sam.collect(wb, [d], args.data_dir, used, sources, cfg=cfg)
        for i in range(3):
            emb[i].extend(part[i])
    emb[2].append("ALL charts on this dashboard are Ember (fallback, annual): these countries have no raw source in the repo "
                  "yet; each is replaced by a raw section when a grid operator / ministry feed is found")

    sam.draw_dashboard(dash, "Africa energy - gas and oil dashboard", *gas)
    sam.draw_dashboard(dash2, "Africa energy - power generation (South Africa, Ghana, Nigeria)", *power)

    dash3 = wb.create_sheet("Dashboard - Power (Ember)", 2)
    sam.draw_dashboard(dash3, "Africa - Ember (fallback, annual) power generation: countries without a raw source yet", *emb)

    src = wb.create_sheet("Sources")
    src.append(["Country", "Dataset", "Workbook", "Publisher", "Link", "Units and notes (from the source workbook)"])
    for cell in src[1]:
        cell.font = Font(bold=True)
    for s in sources:
        src.append(list(s))
        if s[4]:
            src.cell(row=src.max_row, column=5).hyperlink = s[4]
    for col, w in (("A", 14), ("B", 24), ("C", 38), ("D", 40), ("E", 50), ("F", 160)):
        src.column_dimensions[col].width = w

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_charts.save_atomic(wb, args.out)
    print(f"Saved {args.out}: {len(gas[0])} gas charts, {len(power[0])} power charts; tabs {wb.sheetnames}")
    for label, m in (("gas", gas[2]), ("power", power[2])):
        if m:
            print(f"missing ({label}):", m)


if __name__ == "__main__":
    main()
