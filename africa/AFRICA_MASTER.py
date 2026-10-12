"""
Africa master workbook: one file with every African dataset this repo pulls from a raw source (grid operator,
regulator, ministry, national oil company, statistics office - NO Ember), and Dashboard front pages carrying all
their charts. Same layout as the North America / South & Central America masters (south_america/SOUTH_AMERICA_MASTER.py,
whose table/chart/dashboard code this reuses).

  Dashboard          - gas and oil: Nigeria (NNPC Ltd monthly report); Morocco (GME flow at Tarifa = gas burned at
                       Tahaddart and Ain Beni Mathar from 2023, ministry gas-for-power, plant heat rates, import price)
  Dashboard - Power  - generation by source, then lake levels and capacity where a raw feed exists:
                       South Africa (Eskom), Ghana (Energy Commission: weekly WEM statistics, annual statistics with
                       Akosombo / Bui water-year level charts), Nigeria (NERC quarterly report), Cameroon (ARSEL
                       monthly energy balance, 2025 so far), Egypt daily peak / minimum grid load (MOEE, from Oct
                       2026), Morocco gas-fired generation by plant (HCP, annual) and a monthly ESTIMATE from the
                       Tarifa gas flow against Ember (fallback). Where raw data has gaps, Ember fills them as a labelled
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
import types

import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
import add_charts  # noqa: E402
import xlsx_charts  # noqa: E402
import SOUTH_AMERICA_MASTER as sam  # noqa: E402
sys.path.insert(0, os.path.join(ROOT, "middle_east", "capacity"))
import EAST_MED_AFRICA_CAPACITY_COMPLETE as emac  # noqa: E402
import capacity_factors  # noqa: E402

DATA_DIR = sam.DATA_DIR

# (country code, country, workbook, raw sheet / tuple of raw sheets / "*", short dataset name)
DATASETS = [
    ("NG", "Nigeria", "nigeria_nnpc_gas_monthly.xlsx", "Monthly", "NNPC gas and oil"),
    # Morocco gas (africa/MOROCCO_GAS.py): the gas charts here, the generation charts on Dashboard - Power (MA_* below)
    ("MA", "Morocco", "morocco_gas_daily.xlsx", ("Tarifa monthly", "Ministry YTD", "Comtrade monthly", "JODI (fallback)", "Checks"),
     "gas"),
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
    # daily peak / minimum load from the Ministry of Electricity's home page (no history at source: builds from Oct 2026)
    ("EG", "Egypt", "egypt_grid_load_daily.xlsx", "Daily", "grid load"),
    # Morocco gas-fired plants: HCP per-plant generation, monthly ESTIMATE from Tarifa gas vs Ember (FALLBACK)
    ("MA", "Morocco", "morocco_gas_daily.xlsx", ("Plants annual", "Estimate", "Estimate annual"), "gas-fired plants"),
    # EMBER (FALLBACK) gap fills: only where the raw feed above has a hole (see ember_* below)
    ("NG", "Nigeria", "west_africa_power_by_type.xlsx", ("Nigeria",), "Ember fallback monthly"),
    ("NG", "Nigeria", "west_africa_power_by_type_annual.xlsx", ("Nigeria",), "Ember fallback annual"),
    ("CM", "Cameroon", "west_africa_power_by_type_annual.xlsx", ("Cameroon",), "Ember fallback annual"),
]
# Complete monthly installed capacity (middle_east/capacity/EAST_MED_AFRICA_CAPACITY_COMPLETE.py): official annual stocks where reachable
# (Ghana: Energy Commission; Morocco: ONEE via HCP; South Africa: Eskom for nuclear and pumped storage), Ember yearly capacity (fallback)
# elsewhere, interpolated/held to monthly, every filled cell flagged in that workbook. CAPACITY_DATASETS (raw-source countries, on
# 'Dashboard - Power') and EMBER_CAPACITY (Ember-only countries, on 'Dashboard - Power (Ember)') are built in main().
CAP_FILE = emac.OUT_AF
CAPACITY_COUNTRIES = [("ZA", "South Africa"), ("GH", "Ghana"), ("NG", "Nigeria"), ("CM", "Cameroon"), ("EG", "Egypt"), ("MA", "Morocco")]
CAPACITY_DATASETS = [(c, n, CAP_FILE, (n, f"{n} flags"), "capacity") for c, n in CAPACITY_COUNTRIES]
# Countries with NO raw feed yet: Ember (fallback) annual data only, on their own dashboard. Replace an entry with raw
# datasets in the lists above when a raw source is found.
EMBER_ONLY_DATASETS = [(code, c, "west_africa_power_by_type_annual.xlsx", (c,), "Ember fallback annual") for code, c in [
    ("SN", "Senegal"), ("CI", "Cote d'Ivoire"), ("MR", "Mauritania"), ("ML", "Mali"), ("BF", "Burkina Faso"),
    ("GN", "Guinea"), ("SL", "Sierra Leone"), ("LR", "Liberia"), ("BJ", "Benin"), ("TG", "Togo"), ("NE", "Niger"),
    ("GM", "Gambia"), ("GW", "Guinea-Bissau"), ("CV", "Cape Verde"), ("GQ", "Equatorial Guinea")]]
EMBER_CAPACITY = [(c, n, CAP_FILE, (n, f"{n} flags"), "capacity") for c, n, *_ in
                  [(d[0], d[1]) for d in EMBER_ONLY_DATASETS]]
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


# Morocco: one workbook feeds both dashboards; these chart names go to Dashboard - Power, the rest to the gas Dashboard
MA_POWER_CHARTS = ("Plants generation", "Estimate vs Ember")
MA_GAS_CHARTS = ("Tarifa daily", "Gas to power YTD", "Tarifa monthly", "Nm3 per MWh", "Comtrade price")
MA_WB = "morocco_gas_daily.xlsx"
ENTSOG_URL = "https://transparency.entsog.eu/#/points/data?points=ITP-00082"
SPEC_SOURCES = {   # per-chart source line where one workbook combines several publishers
    (MA_WB, "Tarifa daily"): ("ENTSOG transparency platform: Enagas physical flow at Tarifa (ITP-00082); Bcf/d DERIVED at "
                              "11.7 GWh per million Nm3", ENTSOG_URL),
    (MA_WB, "Gas to power YTD"): ("Ministry of Energy Transition 'Chiffres cles' (ONEE fuel use) vs Tarifa exit (ENTSOG / Enagas, "
                                  "DERIVED)", "https://www.mem.gov.ma/Pages/rapports.html"),
    (MA_WB, "Tarifa monthly"): ("ENTSOG (Enagas) and Enagas monthly statistical bulletin",
                                "https://www.enagas.es/en/technical-management-system/energy-data/publications/gas-statistical-bulletin/"),
    (MA_WB, "Nm3 per MWh"): ("HCP Annuaire statistique du Maroc (ONEE tables); Nm3/MWh DERIVED",
                             "https://www.hcp.ma/downloads/Annuaires-statistiques-du-Maroc-Format-Excel_t22392.html"),
    (MA_WB, "Comtrade price"): ("UN Comtrade, Morocco imports HS 271121 (quantity = kWh); US$/MMBtu DERIVED",
                                "https://comtradeplus.un.org/"),
    (MA_WB, "Plants generation"): ("HCP Annuaire statistique du Maroc (ONEE: thermal generation by plant)",
                                   "https://www.hcp.ma/downloads/Annuaires-statistiques-du-Maroc-Format-Excel_t22392.html"),
    (MA_WB, "Estimate vs Ember"): ("ESTIMATE: Tarifa gas (ENTSOG / Enagas) / HCP heat rate; Ember (fallback) for comparison",
                                   ENTSOG_URL),
}
_sam_source_of = sam.source_of


def _source_of(fname, spec_name=None, cfg=None):
    """The shared source line, or this master's per-chart source (SPEC_SOURCES) where one workbook has several."""
    if spec_name is not None and (fname, spec_name) in SPEC_SOURCES:
        return SPEC_SOURCES[(fname, spec_name)]
    return _sam_source_of(fname, spec_name, cfg)


sam.source_of = _source_of   # this process only: collect() looks the function up in its own module

MASTER_SPECS = {"west_africa_power_by_type_annual.xlsx": ember_annual,
                "west_africa_power_by_type.xlsx": ember_monthly}
MASTER_SPECS.update({(CAP_FILE, (d[1], f"{d[1]} flags")): emac.cap_spec(d[1], f"{d[1]} installed capacity, monthly (see Basis)")
                     for d in CAPACITY_DATASETS + EMBER_CAPACITY})
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
    "egypt_grid_load_daily.xlsx": (
        "Ministry of Electricity and Renewable Energy (MOEE), Egypt: home-page 'Grid Loads' panel (yesterday's peak and "
        "minimum load, today's expected peak; the page keeps no history, so the series starts Oct 2026)",
        "http://www.moee.gov.eg/english_new/home.aspx"),
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
    "morocco_gas_daily.xlsx": (
        "ENTSOG / Enagas (GME flow at Tarifa, daily) and Enagas monthly statistical bulletin; HCP Annuaire statistique (ONEE "
        "per-plant generation and gas burn); Ministry of Energy Transition and Sustainable Development ('Chiffres cles', annual "
        "key figures); UN Comtrade (import value); JODI Gas (FALLBACK); Ember (FALLBACK). DERIVED = computed here (11.7 GWh "
        "per million Nm3, Nm3/MWh, Bcf/d, US$/MMBtu); ESTIMATE = monthly gas-fired generation from Tarifa gas",
        "https://transparency.entsog.eu/#/points/data?points=ITP-00082"),
    "west_africa_power_by_type.xlsx": (
        "Ember (fallback) monthly electricity data - gap fill only where the raw feed is missing or coarser (Nigeria monthly)",
        "https://ember-energy.org/data/monthly-electricity-data/"),
    "west_africa_power_by_type_annual.xlsx": (
        "Ember (fallback) yearly electricity data - gap fill only where the raw feed is missing (Cameroon before the ARSEL months; "
        "Nigeria before NERC's 2019 start and for the latest year; and the Ember-only countries on 'Dashboard - Power (Ember)')",
        "https://ember-energy.org/data/yearly-electricity-data/"),
}

CAVEATS = [
    "Morocco: from 2023 the GME flow at Tarifa (Spain -> Morocco, regasified LNG) = the gas burned at Tahaddart and Ain Beni Mathar "
    "(within 1% of the ministry's gas-for-power a year); before Nov 2021 Tarifa carried Algerian transit gas to Spain. Monthly "
    "gas-fired generation is an ESTIMATE (Tarifa gas / the HCP heat rate, about 190 Nm3 per MWh); Ember (fallback) runs 14-18% below "
    "the official HCP figures",
    "Cameroon ARSEL data is raw but covers Jan-Dec 2025 only and is published with an 8+ month lag; Ember (fallback) fills earlier years",
    "Nigeria NERC data is QUARTERLY (grid-connected plants only, from 2019Q1); Ember (fallback) gives the monthly view",
    "Ember (fallback) series are labelled '- Ember (fallback)' and never replace raw data where raw exists",
    "Installed capacity: a gap-free MONTHLY series built from ANNUAL year-end stocks, one basis per country and fuel - official annual for "
    "Ghana (Energy Commission, grid-connected; thermal shown as Gas), Morocco (ONEE via HCP; gas turbines under Other) and South Africa "
    "(Eskom Integrated Report 2025: nuclear and pumped storage only), Ember yearly capacity (FALLBACK) for every other fuel and country "
    "(Nigeria: NERC gives available, not installed, capacity). Wind, solar, bioenergy linear between stocks, others held, flat after the last "
    "stock; blank = no source (e.g. no coal row in Ember). Capacity factors only where a COMPLETE monthly generation feed exists (Ghana WEM, "
    "Cameroon ARSEL 2025, Morocco gas-fired estimate); South Africa's daily feed has too little history yet",
]


def cap_cfg(d):
    """collect() config whose source line for the capacity workbook is this country's (official annual vs Ember fallback fuels)."""
    keys = ("HYDRO_DATASETS", "HYDRO_EXTRA", "DASHBOARD_ONLY", "EMBER", "OPERATORS", "MASTER_SPECS")
    ns = types.SimpleNamespace(**{k: globals()[k] for k in keys})
    ns.SOURCES = {**SOURCES, CAP_FILE: emac.source_line(os.path.join(DATA_DIR, CAP_FILE), d[1], CAP_FILE)}
    return ns


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

    gas = sam.collect(wb, DATASETS, args.data_dir, used, sources, skip=MA_POWER_CHARTS, cfg=cfg)
    # one block per country (dashboard order = order of first appearance in the lists), generation then levels/capacity
    pd_all = RAW_POWER_DATASETS + OTHER_POWER_DATASETS + CAPACITY_DATASETS
    cap_path = os.path.join(args.data_dir, CAP_FILE)
    order = {}
    for d in pd_all:
        order.setdefault(d[1], len(order))
    power = ([], [], [])
    for d in sorted(pd_all, key=lambda d: order[d[1]]):   # stable: keeps list order within a country
        CTX["country"] = d[1]
        part = sam.collect(wb, [d], args.data_dir, used, sources, skip=MA_GAS_CHARTS if d[2] == MA_WB else (),
                           cfg=cap_cfg(d) if d[2] == CAP_FILE else cfg)
        for i in range(3):
            power[i].extend(part[i])
    gens = emac.generation_frames(args.data_dir, "AF") if os.path.exists(cap_path) else {}
    if gens:
        _chart, _row, cf_missing, cf_charts = capacity_factors.add_capacity_factor_sheets(
            wb, used, sam.sheet_name, gens, {n: (cap_path, n) for n in gens}, sam.CHART_W, sam.CHART_H,
            "Generation as on this dashboard (Ghana: Energy Commission WEM; Cameroon: ARSEL plant injections; Morocco: gas-fired ESTIMATE) / "
            "installed capacity: official annual or Ember yearly (fallback), interpolated to months (see the capacity workbook's Basis)", "Africa")
        if _chart:   # countries differ too much in coverage for a summed regional chart: country charts only
            wb.remove(wb[_row[3]])
            used.discard(_row[3])
        for (ch, row) in cf_charts[1:] if _chart else cf_charts:
            power[0].append(ch)
            power[1].append(row)
        power[2].extend(cf_missing)
    power[2].extend(CAVEATS)

    # Ember-only countries: separate dashboard, every chart labelled Ember (fallback, annual)
    emb = ([], [], [])
    for d in EMBER_ONLY_DATASETS:
        CTX["country"] = d[1]
        part = sam.collect(wb, [d], args.data_dir, used, sources, cfg=cfg)
        for i in range(3):
            emb[i].extend(part[i])
    for d in (EMBER_CAPACITY if os.path.exists(cap_path) else []):
        part = sam.collect(wb, [d], args.data_dir, used, sources, cfg=cap_cfg(d))
        for i in range(3):
            emb[i].extend(part[i])
    emb[2].append("ALL charts on this dashboard are Ember (fallback, annual): these countries have no raw source in the repo "
                  "yet; each is replaced by a raw section when a grid operator / ministry feed is found")

    sam.draw_dashboard(dash, "Africa energy - gas and oil dashboard (Nigeria, Morocco)", *gas)
    sam.draw_dashboard(dash2, "Africa energy - power generation (South Africa, Ghana, Nigeria, Cameroon, Egypt, Morocco)", *power)

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
