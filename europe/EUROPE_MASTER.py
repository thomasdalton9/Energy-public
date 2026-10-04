"""
Europe master workbook: one file with every European dataset this repo pulls, and Dashboard front pages carrying
all their charts - the same layout as the South & Central America, North America and Australia + NZ masters
(south_america/SOUTH_AMERICA_MASTER.py, whose table/chart/dashboard code this reuses).

  Dashboard                    - gas: EU gas storage (GIE AGSI+, water year, EU and the main countries) and LNG
                                 terminal send-out / inventory (GIE ALSI); Ireland gas demand, supply and sector use (GNI)
  Dashboard - Power            - Europe power generation by source and installed capacity (sum of the countries below),
                                 day-ahead prices, net imports/exports by country, then one generation chart and one
                                 capacity chart per country (ENTSO-E), Ireland (EirGrid), Turkey (EPIAS), Cyprus
                                 (TSOC), the Rhine at Kaub, and finally a supply/demand BALANCE chart per country
                                 (generation by fuel + net imports + pumped storage/batteries net, against load)
  Dashboard - Capacity factors - capacity factor by generation type, Europe and per country
  <CC> <chart> data            - the table each Dashboard chart plots
  <CC> <dataset> raw           - the full data sheet(s) from each source workbook
  Sources                      - where each dataset comes from, units and notes

Sources: ENTSO-E Transparency Platform (the TSOs' own statutory reporting) is the primary source for power across
Europe - in a settled-week check it matched SMARD (Germany) and RTE (France) to within 0.5% - labelled as such on
every chart's source line. National feeds replace it where they are better (Ireland EirGrid, Turkey EPIAS, Cyprus
TSOC are national). Gas storage and LNG are GIE's aggregation of the operators' own reports.
GB is not in the ENTSO-E generation set (no GB data after 2020); it needs its own feed (Elexon BMRS).

Reads (doesn't refetch) the workbooks the scheduled pulls write to "output/Data and Chart Outputs/". A missing input
is listed on the Dashboard and skipped rather than stopping the rest.

Usage: python3 EUROPE_MASTER.py [--out "output/Data and Chart Outputs/europe_master.xlsx"]
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
sys.path.insert(0, os.path.join(ROOT, "europe"))
import add_charts  # noqa: E402
import capacity_factors  # noqa: E402
import daily_shape  # noqa: E402
import fundamentals  # noqa: E402
import xlsx_charts  # noqa: E402
import SOUTH_AMERICA_MASTER as sam  # noqa: E402
from europe_countries import COUNTRIES  # noqa: E402

DATA_DIR = sam.DATA_DIR
FUELS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Other"]   # same order/colours as the other dashboards
START = "2021-01-01"
MIN_SHARE = 0.95   # a country joins the Europe total if it has this share of the months since START

ENTSOE = ("ENTSO-E Transparency Platform (TSOs' reporting): actual generation per production type; Great Britain: Elexon BMRS + NESO; "
          "Ireland: EirGrid + Ember",
          "https://transparency.entsoe.eu/")

# (code, country, workbook, raw sheet / tuple of raw sheets / "*", short dataset name)
DATASETS = [
    ("EU", "Europe", "eu_gas_storage_daily.xlsx", "Daily", "gas storage"),
    ("EU", "Europe", "eu_lng_terminals_daily.xlsx", "Daily", "LNG terminals"),
    ("IE", "Ireland", "ireland_gas_combined_daily.xlsx", "*", "gas"),
    ("IE", "Ireland", "ireland_gni_transparency_daily.xlsx", "*", "gas by sector"),
    ("EU", "Europe", "europe_gas_flows_daily.xlsx", "*", "ENTSOG"),
    ("GB", "Great Britain", "gb_gas_nts_daily.xlsx", "Daily", "NTS gas"),
    ("EU", "Europe", "europe_tso_gas_demand_daily.xlsx", "Daily", "TSO gas consumption (DE FR ES DK PT)"),
    ("EU", "Europe", "europe_tso_gas_demand_cee_daily.xlsx", "Daily", "TSO gas consumption (AT CZ LT)"),
    ("EU", "Europe", "europe_tso_gas_demand_extra_daily.xlsx", "*", "TSO gas consumption (PL RO HR FI ES)"),
    ("EU", "Europe", "eurostat_gas_monthly.xlsx", "Monthly", "Eurostat gas (benchmark)"),
    ("NO", "Norway", "norway_gassco_gas_flows_daily.xlsx", "Daily", "Gassco gas exports"),
    ("EU", "Europe", "europe_biomethane_operators.xlsx", "*", "biomethane injection (FR DK NL)"),
    ("NL", "Netherlands", "netherlands_cbs_gas_monthly.xlsx", "Monthly", "CBS gas balance"),
    ("EU", "Europe", "europe_biomethane_statistics.xlsx", "*", "biomethane statistics (GB AT SE, EU annual)"),
]
RAW_POWER_DATASETS = (
    [(code, name, f"{slug}_power_generation_daily.xlsx", "Daily", "power") for code, (name, slug, _) in COUNTRIES.items()]
    + [("GB", "Great Britain", "great_britain_power_generation_daily.xlsx", "Daily", "power"),
       ("CH", "Switzerland (Swissgrid)", "switzerland_swissgrid_power_daily.xlsx", "Daily", "power"),
       ("NL", "Netherlands (CBS)", "netherlands_cbs_power_daily.xlsx", "Daily", "power"),
       ("DK", "Denmark (Energinet load)", "denmark_energinet_load_daily.xlsx", "Daily", "consumption"),
       ("IE-EG", "Ireland (EirGrid)", "ireland_smartgrid_15min.xlsx", (), "power"),
       ("IE", "Ireland (Ember)", "ember_europe_power_monthly.xlsx", "*", "Ember"),
       ("IE", "Ireland (EirGrid)", "ireland_eirgrid_system_data.xlsx", "Daily", "EirGrid"),
       ("TR", "Turkey", "turkey_generation_mix_dashboard_daily.xlsx", "*", "power"),
       ("CY", "Cyprus", "cyprus_generation_mix_daily.xlsx", "*", "power")])
CAPACITY_DATASETS = [(code, name, f"{slug}_power_capacity.xlsx", "Monthly", "capacity")
                     for code, (name, slug, _) in COUNTRIES.items()]
PRICE_DATASETS = [("EU", "Europe", "europe_power_prices_daily.xlsx", "Daily", "power prices"),
                  ("EU", "Europe", "europe_cross_border_flows_daily.xlsx", "*", "flows")]
FLOWS_FILE = "europe_cross_border_flows_daily.xlsx"
# Raw national generation feeds that replace a country's ENTSO-E workbook (same Daily layout)
GEN_OVERRIDE = {"Switzerland": "switzerland_swissgrid_power_daily.xlsx", "Netherlands": "netherlands_cbs_power_daily.xlsx"}
CH_BALANCE_SRC = ("Swissgrid via the Swiss Federal Office of Energy (production by carrier, national consumption); physical imports/exports and pumping consumption from the BFE monthly electricity balance",
                  "https://www.energiedashboard.ch")
NL_BALANCE_SRC = ("Statistics Netherlands (CBS) electricity balance (production by source incl. rooftop solar); load = CBS consumption incl. losses; cross-border flows from ENTSO-E (they match CBS imports/exports)",
                  "https://opendata.cbs.nl/ODataApi/odata/84575NED")
# Raw national load that replaces ENTSO-E's 'actual total load' (generation by fuel and flows stay ENTSO-E)
LOAD_OVERRIDE = {"Denmark": "denmark_energinet_load_daily.xlsx"}
DK_BALANCE_SRC = ("ENTSO-E generation and flows; load = Energinet settlement gross consumption (incl. grid losses and power-to-heat), which ENTSO-E's Danish load omits",
                  "https://www.energidataservice.dk/tso-electricity/ProductionConsumptionSettlement")
GB_FILE = "great_britain_power_generation_daily.xlsx"
GB_BALANCE_SRC = ("Elexon BMRS (metered generation, interconnectors) and NESO (national demand, embedded wind and solar)",
                  "https://bmrs.elexon.co.uk/")
HYDRO_DATASETS = [("DE", "Germany (Rhine)", "rhine_kaub_level_daily.xlsx", "Data", "river level")]
HYDRO_EXTRA = {}
DASHBOARD_ONLY = {}
EMBER = set()
OPERATORS = {}


def rhine(path):
    """rhine_kaub_level_daily.xlsx draws its own chart (REGISTRY entry is None); the master needs a spec."""
    d = add_charts.by_date(add_charts.read(path, "Data"), "date")
    return [{"name": "Rhine", "water_year": d["level_cm"], "title": "Rhine water level at Kaub (PEGELONLINE)",
             "units": "cm (gauge datum)", "y_decimals": 0}]


MASTER_SPECS = {"rhine_kaub_level_daily.xlsx": rhine}

SOURCES = {
    "eu_gas_storage_daily.xlsx": ("Gas Infrastructure Europe, AGSI+ (storage operators' own reports)", "https://agsi.gie.eu/"),
    "eu_lng_terminals_daily.xlsx": ("Gas Infrastructure Europe, ALSI (LNG terminal operators' own reports)", "https://alsi.gie.eu/"),
    "europe_gas_flows_daily.xlsx": ("ENTSOG Transparency Platform: physical flows at interconnection points (TSOs' own reporting)",
                                    "https://transparency.entsog.eu/"),
    "ireland_gas_combined_daily.xlsx": ("Gas Networks Ireland (GNI) transparency and open data",
                                        "https://www.gasnetworks.ie/corporate/gas-regulation/transparency/"),
    "ireland_gni_transparency_daily.xlsx": ("Gas Networks Ireland (GNI) transparency pages",
                                            "https://www.gasnetworks.ie/corporate/gas-regulation/transparency/"),
    "gb_gas_nts_daily.xlsx": ("National Gas Transmission Data Portal (NTS demand by sector and supply by entry point)",
                              "https://data.nationalgas.com/find-gas-data"),
    "gb_storage_sites_daily.xlsx": ("National Gas Transmission Data Portal (storage stock, inflow and outflow by site)",
                                    "https://data.nationalgas.com/find-gas-data"),
    "europe_tso_gas_demand_daily.xlsx": ("Gas TSOs' own series: Trading Hub Europe (DE), ODRE / GRTgaz-Teréga-RTE (FR), Enagás (ES), "
                                         "Energinet (DK), REN DataHub (PT)", "https://www.tradinghub.eu/"),
    "europe_tso_gas_demand_cee_daily.xlsx": ("Gas TSOs' own series: AGGM (AT), NET4GAS CAMS system balance (CZ), Amber Grid (LT)",
                                             "https://platform.aggm.at/"),
    "europe_tso_gas_demand_extra_daily.xlsx": ("Gas TSOs' own series: Gaz-System (PL), Transgaz (RO), Plinacro (HR), Gasgrid Finland (FI); "
                                               "Enagás monthly statistical bulletin (ES 2021-22)", "https://www.gasgrid.fi/"),
    "norway_gassco_gas_flows_daily.xlsx": ("Gassco: Norwegian gas flows by delivery destination (daily, mcm/d converted at 11.2 GWh per mcm)",
                                           "https://gassco.eu/"),
    "europe_biomethane_statistics.xlsx": ("Biomethane: DESNZ Energy Trends (GB), AGGM (AT), Energimyndigheten (SE), Eurostat nrg_bal_c annual "
                                          "biogases blended into natural gas (all EU27, used where there is no operator series)",
                                          "https://ec.europa.eu/eurostat/databrowser/view/nrg_bal_c"),
    "europe_biomethane_operators.xlsx": ("Biomethane injected into the gas grids: ODRE (France, daily), Energinet Gasflow (Denmark, daily), "
                                         "CBS StatLine 86103NED (Netherlands, monthly, includes a little refinery-gas conversion)",
                                         "https://odre.opendatasoft.com/"),
    "netherlands_cbs_gas_monthly.xlsx": ("Statistics Netherlands (CBS) StatLine 86103NED natural gas balance (monthly): Dutch production and total consumption "
                                         "replace ENTSOG's in the Netherlands and Germany + Netherlands gas balances",
                                         "https://opendata.cbs.nl/ODataApi/odata/86103NED"),
    "eurostat_gas_monthly.xlsx": ("Eurostat nrg_cb_gasm monthly natural gas balance (validation benchmark only, not used in the charts)",
                                  "https://ec.europa.eu/eurostat/databrowser/view/nrg_cb_gasm"),
    "ireland_eirgrid_system_data.xlsx": ("EirGrid / SONI System and Renewable Data Reports (Ireland and Northern Ireland system data)",
                                         "https://www.eirgrid.ie/grid/system-and-renewable-data-reports"),
    "ember_europe_power_monthly.xlsx": ("Ember monthly electricity data (fallback for Ireland, where the ENTSO-E all-island feed is "
                                        "incomplete; CC-BY-4.0)", "https://ember-energy.org/data/monthly-electricity-data/"),
    "great_britain_power_generation_daily.xlsx": ("Elexon BMRS (FUELHH) and NESO historic demand data (national demand, "
                                                  "embedded wind and solar)", "https://bmrs.elexon.co.uk/"),
    "ireland_smartgrid_15min.xlsx": ("EirGrid / SONI Smart Grid Dashboard", "https://www.smartgriddashboard.com/"),
    "turkey_generation_mix_dashboard_daily.xlsx": ("EPIAS Transparency Platform (Turkey)", "https://seffaflik.epias.com.tr/"),
    "cyprus_generation_mix_daily.xlsx": ("Transmission System Operator Cyprus (TSOC)", "https://tsoc.org.cy/"),
    "rhine_kaub_level_daily.xlsx": ("PEGELONLINE (German federal waterways administration), Kaub gauge",
                                    "https://www.pegelonline.wsv.de/"),
    "europe_power_prices_daily.xlsx": ("ENTSO-E Transparency Platform: day-ahead prices", "https://transparency.entsoe.eu/"),
    "europe_cross_border_flows_daily.xlsx": ("ENTSO-E Transparency Platform: cross-border physical flows",
                                             "https://transparency.entsoe.eu/"),
}
GAS_BALANCE_SRC = ("ENTSOG (production, pipeline flows, consumption), TSO consumption series (Germany THE, France ODRE, Spain Enagas, Denmark Energinet, Portugal REN, Austria AGGM, Czechia NET4GAS, Lithuania Amber Grid, Finland Gasgrid, Great Britain National Gas, Ireland GNI), Elering (Estonia flows incl. Karksi, Balticconnector), Gassco (Norway to Great Britain), biomethane injection (France ODRE, Denmark Energinet, Netherlands CBS, Austria AGGM, other countries Eurostat annual), GIE ALSI (LNG send-out), GIE AGSI+ (storage)",
                   "https://transparency.entsog.eu/")
GAS_FLOWS_FILE = "europe_gas_flows_daily.xlsx"
BALANCE_SRC = ("ENTSO-E Transparency Platform: generation, load and cross-border physical flows (Great Britain: Elexon BMRS + NESO; "
               "Ireland: EirGrid + Ember)",
               "https://transparency.entsoe.eu/")
for _code, (_name, _slug, _zones) in COUNTRIES.items():
    SOURCES[f"{_slug}_power_generation_daily.xlsx"] = ENTSOE
    SOURCES[f"{_slug}_power_capacity.xlsx"] = ("ENTSO-E Transparency Platform: installed capacity per production type",
                                               "https://transparency.entsoe.eu/")

SOURCES["denmark_energinet_load_daily.xlsx"] = ("Energinet, Energi Data Service: ProductionConsumptionSettlement (gross consumption, production, exchanges)",
                                                "https://www.energidataservice.dk/tso-electricity/ProductionConsumptionSettlement")


MIN_COVERAGE = 0.5          # a month whose generation/load ratio is under this share of the country's typical ratio is not usable
LATE_START_OK = pd.Timestamp("2022-03-01")   # a country whose usable history starts by here still joins the Europe totals


def covered_months(daily_mwh):
    """Boolean Series by month start: generation reported (fuels, excluding pumped storage and batteries) is at least
    MIN_COVERAGE of load. Catches a feed that only started reporting part-way through (e.g. Sweden's hydro and nuclear
    begin in Dec 2021, so Jan-Nov 2021 would show about 20% of load)."""
    fuel = [c for c in daily_mwh if c.endswith("_MWh") and c not in (
        "Total_MWh", "Load_MWh", "PumpedStorage_MWh", "Storage_MWh", "PumpedStorageConsumption_MWh", "StorageCharging_MWh")]
    gen = daily_mwh[fuel].sum(axis=1, min_count=1).resample("MS").sum(min_count=1)
    load = daily_mwh["Load_MWh"].resample("MS").sum(min_count=1)
    ratio = gen / load
    good = ratio >= MIN_COVERAGE * ratio.median()   # relative: import-heavy countries (Baltics) legitimately sit well below 1
    first = good.idxmax() if good.any() else None
    # only the leading months before the feed was complete are dropped; seasonal dips later (hydro) are real
    return pd.Series([bool(first is not None and d >= first) for d in ratio.index], index=ratio.index)


def core_months(frames, start, notes):
    """Countries that can be summed, and the months they all have. A country joins if its usable history starts by
    LATE_START_OK and is >= MIN_SHARE complete from its own first month; the totals then start when the last of them
    joins. The others are listed as not included."""
    last = max(f.index.max() for f in frames.values())
    keep = {}
    for name, f in frames.items():
        span = pd.date_range(f.index.min(), last, freq="MS")
        share = len(f.index.intersection(span)) / len(span)
        if f.index.min() <= LATE_START_OK and share >= MIN_SHARE:
            keep[name] = f
        else:
            notes.append(f"NOT INCLUDED: {name} (usable months {f.index.min():%b/%y}-{f.index.max():%b/%y}, "
                         f"{share:.0%} complete from its first month)")
    if not keep:
        return {}, []
    first = max(max(f.index.min() for f in keep.values()), pd.Timestamp(start))
    late = [f"{n} from {f.index.min():%b/%y}" for n, f in keep.items() if f.index.min() > pd.Timestamp(start) and f.index.min() == first]
    if late:
        notes.append(f"Totals start {first:%b/%y} because " + ", ".join(late)
                     + f" (earlier months of the feed are missing or under {MIN_COVERAGE:.0%} of the country's typical share of load)")
    months = sorted(set.intersection(*(set(f.index[f.index >= first]) for f in keep.values())))
    return keep, months


def monthly_gwh(path):
    """Standard Daily sheet (MWh) -> complete months, GWh, in the dashboard fuel groups (storage/pumping left out)."""
    d = add_charts.by_date(add_charts.read(path, "Daily"), "date")
    d = d[[c for c in d.columns if str(c).endswith("_MWh")]].apply(pd.to_numeric, errors="coerce")
    m = d.resample("MS").sum(min_count=1) / 1000.0
    if "Load_MWh" in d:
        ok = covered_months(d)
        m = m[m.index.isin(ok.index[ok])]
    last = d.dropna(how="all").index.max()
    if last < last + pd.offsets.MonthEnd(0):   # drop the month in progress
        m = m[m.index < last.to_period("M").to_timestamp()]
    m = m[m.index >= START].rename(columns=lambda c: c.replace("_MWh", "_GWh"))
    m = m.rename(columns={"Oil_GWh": "Other Fossil_GWh", "Other_GWh": "Other Renewables_GWh"})
    return add_charts.power_mix(m)


def ember_ireland(data_dir):
    """Ember monthly generation for the Republic of Ireland, GWh, in the dashboard fuel groups."""
    d = add_charts.by_date(add_charts.read(os.path.join(data_dir, "ember_europe_power_monthly.xlsx"), "Ireland"), "Month")
    d = d[[c for c in d.columns if str(c).endswith("_GWh") and c not in ("Total_GWh", "Demand_GWh", "NetImports_GWh")]]
    return add_charts.power_mix(d.apply(pd.to_numeric, errors="coerce"))


def ireland_monthly(data_dir):
    """Republic of Ireland, GWh per month: (generation by fuel group, EirGrid demand). Fuel split from Ember (it agrees with
    EirGrid's own totals within about 2%); months after Ember's last (it lags 3-4 months) come from EirGrid's monthly summary -
    its wind and solar, hydro and the bioenergy/other-renewable share of the same month a year earlier, and the thermal
    remainder split into gas, coal and oil/other fossil by Ember's trailing 12-month shares. Returns (mix, demand, note)."""
    em = ember_ireland(data_dir)
    eg = add_charts.by_date(add_charts.read(os.path.join(data_dir, "ireland_eirgrid_system_data.xlsx"), "Monthly"), "Month")
    gen, dem = eg["IE_Generation"], eg["IE_Demand"]
    em_last = em.index.max()
    thermal = [c for c in ("Gas", "Coal", "Other") if c in em]
    shares = em[thermal].tail(12).sum()
    shares = shares / shares.sum()
    rows = {}
    for d in gen.index[gen.index > em_last]:
        prev = d - pd.DateOffset(years=1)
        base = em.loc[prev] if prev in em.index else None
        if base is None or pd.isna(gen[d]):
            continue
        wind, solar = eg["IE_Wind"].get(d, float("nan")), eg["IE_Solar"].get(d, float("nan"))
        hydro = base.get("Hydro", 0.0)
        rest = gen[d] - wind - solar - hydro
        row = {"Hydro": hydro, "Wind": wind, "Solar": solar, "Nuclear": 0.0}
        for c in thermal:
            row[c] = rest * shares[c]
        rows[d] = row
    ext = pd.DataFrame.from_dict(rows, orient="index").reindex(columns=em.columns).fillna(0.0) if rows else pd.DataFrame(columns=em.columns)
    mix = pd.concat([em, ext]).sort_index()
    note = (f"Ireland (Republic of Ireland): fuel split from Ember to {em_last:%b/%y}; later months from EirGrid's monthly system data "
            f"(wind, solar, total generation; thermal split by Ember's trailing 12-month shares)") if len(ext) else ""
    return mix, dem.reindex(mix.index), note


def europe_generation(data_dir, frames_out=None):
    """Sum of the ENTSO-E countries with a near-complete record, GWh per month, over the months they all have."""
    frames, notes = {}, []
    for code, (name, slug, _) in COUNTRIES.items():
        fname = GEN_OVERRIDE.get(name, f"{slug}_power_generation_daily.xlsx")
        try:
            m = monthly_gwh(os.path.join(data_dir, fname))
            frames[name] = m[m.sum(axis=1) > 0]
        except Exception as e:  # noqa: BLE001
            notes.append(f"NOT INCLUDED: {name} ({type(e).__name__}: {e})")
    if not frames:
        return pd.DataFrame(), notes
    last_all = max(f.index.max() for f in frames.values())
    # Great Britain (Elexon + NESO) is not in ENTSO-E generation; Ireland's ENTSO-E all-island feed covers only part of demand,
    # so Ireland (Republic) comes from Ember, which lags a few months: later months repeat the same month of the previous year.
    try:
        frames["Great Britain"] = monthly_gwh(os.path.join(data_dir, GB_FILE))
    except Exception as e:  # noqa: BLE001
        notes.append(f"NOT INCLUDED: Great Britain ({type(e).__name__}: {e})")
    try:
        mix, _dem, note = ireland_monthly(data_dir)
        frames.pop("Ireland (all-island SEM)", None)
        frames["Ireland"] = mix
        if note:
            notes.append(note)
    except Exception as e:  # noqa: BLE001
        notes.append(f"Ireland from Ember/EirGrid not available ({type(e).__name__}: {e}); ENTSO-E all-island feed used")
    if frames_out is not None:
        frames_out.update(frames)
    keep, months = core_months(frames, START, notes)
    if not months:
        return pd.DataFrame(), notes + ["no month common to every country yet"]
    total = sum(f.reindex(index=months, columns=FUELS).fillna(0) for f in keep.values())
    total.index.name = "date"
    notes.insert(0, f"{len(keep)} countries summed (ENTSO-E, {months[0]:%b/%y}-{months[-1]:%b/%y}): " + ", ".join(keep))
    return total, notes


def europe_capacity(data_dir):
    """Annual installed capacity, GW, summed over the countries with figures for the year (ENTSO-E, 1 January)."""
    frames, notes = {}, []
    for code, (name, slug, _) in COUNTRIES.items():
        path = os.path.join(data_dir, f"{slug}_power_capacity.xlsx")
        try:
            d = add_charts.by_date(add_charts.read(path, "Monthly"), "date")
            frames[name] = add_charts.capacity_groups(pd.DataFrame(
                {f: d.get(f"{f}_MW") for f in add_charts.CAPACITY_FUELS}, index=d.index)) / 1000.0
        except Exception as e:  # noqa: BLE001
            notes.append(f"NOT INCLUDED: {name} ({type(e).__name__}: {e})")
    if not frames:
        return pd.DataFrame(), notes
    first = min(f.index.min() for f in frames.values())
    last = max(f.index.max() for f in frames.values())
    rng = pd.date_range(first, last, freq="YS")
    core = {}
    for name, f in frames.items():
        share = len(f.index.intersection(rng)) / len(rng)
        if share >= MIN_SHARE:
            core[name] = f
        else:
            notes.append(f"NOT INCLUDED: {name} (figures for only {share:.0%} of {first:%Y}-{last:%Y}: "
                         f"{f.index.min():%Y}-{f.index.max():%Y})")
    years = sorted(set.intersection(*(set(f.index) for f in core.values()))) if core else []
    rows, used_years = [sum(core[n].loc[y] for n in core) for y in years], years
    frames = core
    if not rows:
        return pd.DataFrame(), notes + ["no year with figures for the countries counted"]
    total = pd.DataFrame(rows, index=pd.DatetimeIndex(used_years))
    total.index.name = "date"
    notes.insert(0, f"{len(frames)} countries (ENTSO-E annual installed capacity, 1 January; storage excluded)")
    return total, notes


def total_chart(wb, used, power, pos, df_total, notes, sheet, title, units, src, note_head, date_format="%Y-%m",
                label="Europe", line_cols=()):
    """Chart + data tab; pos=None appends to the dashboard lists, otherwise inserts at pos."""
    ws = wb.create_sheet(sam.sheet_name(sheet, used))
    df, n_bars = xlsx_charts.prepare(df_total, line_cols)
    xlsx_charts.write_table(ws, df, date_format)
    ws.cell(row=1, column=df.shape[1] + 4, value=note_head)
    for i, note in enumerate(notes, start=2):
        ws.cell(row=i, column=df.shape[1] + 4, value=note)
    chart = xlsx_charts.build_chart(ws, df, n_bars, title, units, "stacked_bar", date_format, width=sam.CHART_W,
                                    height=sam.CHART_H, gridlines=False, inner=xlsx_charts.DASHBOARD_INNER)
    missing = [n.split(":", 1)[1].split("(")[0].strip() for n in notes if n.startswith("NOT INCLUDED")]
    row = (label, title + (f" (not counted: {', '.join(missing)})" if missing else ""),
           df.index.max().strftime("%b/%y" if date_format != "%Y" else "%Y"), ws.title, *src)
    if pos is None:
        power[0].append((chart, (src[0], src[1])))
        power[1].append(row)
    else:
        power[0].insert(pos, (chart, (src[0], src[1])))
        power[1].insert(pos, row)


def monthly_cover(daily, min_cover=0.9):
    """Daily frame -> monthly totals scaled to the full month, blank where fewer than min_cover of the days exist."""
    g = daily.resample("MS")
    n = g.count()
    dim = pd.Series(n.index.days_in_month, index=n.index)
    out = g.sum(min_count=1).mul(dim, axis=0) / n.where(n > 0)
    return out.where(n.div(dim, axis=0) >= min_cover)


BAL_ITEM = "Balancing item (implied embedded supply / losses)"
# Raw cross-border flows ENTSO-E does not publish, added to the country's ENTSO-E net imports (daily GWh = MWh/1000)
NETIMP_EXTRA = {"Spain": ("spain_ree_exchanges_daily.xlsx", "MA_AD_NetImports_MWh")}
ES_BALANCE_SRC = ("ENTSO-E generation, load and France/Portugal flows; net imports also include Red Electrica's physical exchanges with Morocco and Andorra "
                  "(no ENTSO-E bidding zone; 2.1-4.0 TWh a year of exports)", "https://www.ree.es/en/datos/intercambios")
# Pumping consumption ENTSO-E does not publish (generation is reported, consumption is not), from national statistics; subtracted from 'Pumped & battery (net)'
PUMP_CONS_OVERRIDE = {"Slovenia": "slovenia_sistat_pumping_daily.xlsx"}
SI_BALANCE_SRC = ("ENTSO-E generation, load and flows; pumping consumption of the Avce pumped-storage plant (not published by ENTSO-E) from the Statistical Office of Slovenia "
                  "(SiStat 1817602S, annual, spread over the days with ENTSO-E's pumped output)", "https://pxweb.stat.si/SiStatData/pxweb/en/Data/")
BALANCE_COLS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Other", "Net imports",
                "Pumped & battery (net)", "Load"]


# Why reported supply can fall short of load for a country (ENTSO-E reports what TSOs meter; the rest is not in the feed).
KNOWN_GAPS = {
    "Switzerland": "Generation is Swissgrid's own production by carrier (storage hydro is gross of pumped-storage output); pumping consumption, "
                   "imports and exports are BFE's monthly electricity balance spread over the days. ENTSO-E's Swiss hydro was incomplete (supply/load ~70%) and is no longer used.",
    "Netherlands": "Generation is CBS monthly production by source (including rooftop solar) spread evenly over the days. Load is CBS consumption incl. distribution losses (ENTSO-E load is 10% lower in 2021-22, equal within 1% in 2024-25), so supply vs load closes by construction; flows are ENTSO-E and match CBS.",
    "Germany": "Industrial self-generation and part of the small PV and biomass fleet are not in the ENTSO-E per-type feed (Eurostat/Destatis counts 441 TWh net generation in 2024 against 429 TWh in ENTSO-E), so supply is 95-98% of load; a monthly Eurostat check is in germany_eurostat_power_daily.xlsx but is not used here.",
    "Italy": "ENTSO-E generation (about 215-220 TWh in 2024) omits embedded and self-consumed generation (Eurostat/Terna count 263.5 TWh) and ENTSO-E load (273 TWh) omits the matching demand (Terna 312 TWh); supply is 95-98% of load. Eurostat checks are in italy_eurostat_power_daily.xlsx but are not used here.",
    "Great Britain": (
        "Load now includes station load (power stations' own use, 4.7-4.8 TWh a year = NESO TSD - ND - pumping - interconnector exports, half-hourly): Elexon FUELHH metered output is gross of it while "
        "NESO national demand excludes it, which had left supply 2-3% above load (8 TWh in 2025). Supply is now 100.2-101.2% of load in 2021-25."),
    "Bulgaria": "ENTSO-E load (37.9 TWh in 2025) is about 3 TWh above Eurostat consumption incl. losses (34.9 TWh) while generation (36.7 TWh) equals Eurostat's and net exports match (1.3 TWh), so supply is 93% of load in 2025: a load-definition difference.",
    "Romania": "ENTSO-E load (53.6 TWh in 2025) is about 3 TWh above Eurostat available-to-market (50.4 TWh) while generation matches (46.7 vs 47.4 TWh), so supply is 96% of load in 2025: a load-definition difference.",
    "Montenegro": (
        "The ENTSO-E physical flow on the Bosnia-Montenegro border (2.96 TWh BA>ME, 0.59 ME>BA in 2025) does not close either side's balance: Montenegro is oversupplied (2025 +36%) and Bosnia undersupplied (-13%). "
        "Checked in a probe (discovery_archive/europe/BALKAN_FLOWS_PROBE.py): the sign and zone mapping are right (A11 is labelled by direction), the finalised SCHEDULES (A09) are lower on BA>ME (2.23 TWh) but higher on RS>ME (1.59 vs 0.47 TWh physical): "
        "scheduled Serbia-Montenegro energy physically loops through Bosnia, so the physical BA>ME flow is partly transit. Montenegro's 2025 net imports are 2.1 TWh physical against 1.0 TWh scheduled, which would bring it to about 100%, but Bosnia's -13% is unchanged under schedules "
        "(net -2.2 vs -2.3 TWh), so it is a Bosnian generation/load coverage issue and physical flows are kept (the correct quantity for a physical balance). Montenegro's TSO (CGES) has no feed reachable from GitHub "
        "(discovery_archive/europe/CGES_PROBE.py), and 2025 generation is low because the Pljevlja coal plant was out Apr-Nov. Treat the Balkan (BA, ME, MK, XK, RS) balances as indicative."),
    "Bosnia and Herzegovina": "See Montenegro: the Bosnia-Montenegro physical flow looks overstated; supply is 13% below load in 2025 (ratio was 97-102% before).",
    "North Macedonia": "Small system with unreliable ENTSO-E load/flow reporting (supply 89-91% of load in 2023-24, 100% in 2025).",
    "Kosovo": "KOSTT generation is metered at the plant and load includes distribution losses and theft; supply is about 5% below load.",
    "Slovenia": "Supply is about 4% above load every year: ENTSO-E Slovenian load excludes some demand that generation and flows cover (grid losses/closed distribution systems).",
    "Serbia": "Supply is about 4% above load every year, consistent with a load definition that is net of transmission losses.",
    "Lithuania": "Supply is 3-4% above load since 2023, after the Baltic synchronisation changed the metered border flows.",
    "Denmark": "Load is Energinet's settlement gross consumption (incl. grid losses and 2.7 TWh of power-to-heat in 2025). ENTSO-E's Danish load is 4-7% lower, which made supply look 4-7% too high; "
               "ENTSO-E net imports match Energinet's exchanges (7.4 TWh in 2025). Remaining gap: ENTSO-E generation is about 1 TWh above Energinet's production.",
    "Poland": "ENTSO-E load (164 TWh in 2024) is a gross figure incl. station own use, above Eurostat consumption incl. losses (155 TWh) while generation is close (157 vs 155 TWh); supply is 94-95% of load in 2022-23 and 98-100% since.",
    "Slovakia": "Net imports exclude double-counted Ukraine flows (ENTSO-E reports the same tie-lines under three Ukraine zones); supply now matches load within 1%.",
    "Finland": "2021-22 imports from Russia are not in the ENTSO-E flow data used here.",
    "Lithuania": "Imports from Belarus/Russia before 2022 are not in the ENTSO-E flow data used here.",
    "Ireland": "Republic of Ireland only (Northern Ireland is in the UK). Net imports are EirGrid demand less generation, so supply equals load by construction; "
               "the fuel split is Ember's, with EirGrid's monthly totals after Ember's last month.",
}


def coverage_notes(name, b):
    """Supply-to-load ratio over the last 12 months shown, plus the known reason for any shortfall."""
    cols = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Other", "Net imports", "Pumped & battery (net)"]
    tail = b.tail(12)
    ratio = tail[cols].sum().sum() / tail["Load"].sum()
    out = [f"Supply (generation + net imports + pumped/battery net) is {ratio:.1%} of load over the last 12 months shown.",
           f"The '{BAL_ITEM}' bar is load minus that supply: positive = embedded or self-consumed generation the feed does not carry, negative = losses or "
           "demand the supply side does not cover. It is an implied figure, not measured data, and makes the bars add up to load."]
    if name in KNOWN_GAPS:
        out.append(KNOWN_GAPS[name])
    return out


def country_balance(gen_path, net_imports, load_override=None):
    """Monthly GWh supply/demand balance for one country: generation by fuel group, net imports, pumped storage and
    batteries (discharge minus pumping/charging) and load - over the days that have load and flows."""
    d = add_charts.by_date(add_charts.read(gen_path, "Daily"), "date").apply(pd.to_numeric, errors="coerce")

    def col(name):
        return d[name] if name in d else pd.Series(float("nan"), index=d.index)

    day = pd.DataFrame({
        "Hydro": col("Hydro_MWh"), "Gas": col("Gas_MWh"), "Wind": col("Wind_MWh"), "Solar": col("Solar_MWh"),
        "Coal": col("Coal_MWh"), "Nuclear": col("Nuclear_MWh"),
        "Other": d[[c for c in ("Bioenergy_MWh", "Oil_MWh", "Other_MWh") if c in d]].sum(axis=1, min_count=1),
        "Pumped & battery (net)": (col("PumpedStorage_MWh").fillna(0) + col("Storage_MWh").fillna(0)
                                   - col("PumpedStorageConsumption_MWh").fillna(0) - col("StorageCharging_MWh").fillna(0)),
        "Load": col("Load_MWh")}) / 1000.0
    if load_override is not None:   # raw national load (GWh/day) replaces the ENTSO-E load
        day["Load"] = load_override.reindex(day.index)
    day["Net imports"] = net_imports
    day = day.dropna(subset=["Load", "Net imports", "Gas"]).fillna(0)
    day = day[day.index >= START]
    if day.empty:
        return pd.DataFrame()
    m = monthly_cover(day[BALANCE_COLS], 0.75).dropna(how="all")
    fuels = m[["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Other"]].sum(axis=1)
    r = fuels / m["Load"]
    good = r >= MIN_COVERAGE * r.median()
    m = m[m.index >= (good.idxmax() if good.any() else m.index.min())]   # leading months before a feed was complete (e.g. Sweden hydro/nuclear before Dec 2021)
    last = day.index.max()
    if last < last + pd.offsets.MonthEnd(0):   # drop the month in progress
        m = m[m.index < last.to_period("M").to_timestamp()]
    return m.dropna(how="any")


GAS_BAL_COLS = ["Production", "Pipeline imports", "LNG send-out", "Storage withdrawals", "Pipeline exports",
                "Storage injections", "Consumption"]
EU27_GAS = ["AT", "BE", "BG", "HR", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT", "LV", "LT", "LU", "NL", "PL",
            "PT", "RO", "SK", "SI", "ES", "SE"]
GAS_NAMES = {"AT": "Austria", "BE": "Belgium", "BG": "Bulgaria", "HR": "Croatia", "CZ": "Czechia", "DK": "Denmark",
             "EE": "Estonia", "FI": "Finland", "FR": "France", "DE": "Germany", "GR": "Greece", "HU": "Hungary",
             "IE": "Ireland", "IT": "Italy", "LV": "Latvia", "LT": "Lithuania", "LU": "Luxembourg", "NL": "Netherlands",
             "PL": "Poland", "PT": "Portugal", "RO": "Romania", "SK": "Slovakia", "SI": "Slovenia", "ES": "Spain",
             "SE": "Sweden", "UK": "United Kingdom"}


def _col(df, name):
    return df[name] if name in df else pd.Series(float("nan"), index=df.index)


def _monthly_twh(day, line_floor=12):
    """Daily GWh frame -> complete months in TWh (months with >= 90% of days), blank rows dropped."""
    m = monthly_cover(day).dropna(how="all") / 1000.0
    last = day.dropna(how="all").index.max()
    if len(day) and last < last + pd.offsets.MonthEnd(0):   # drop the month in progress
        m = m[m.index < last.to_period("M").to_timestamp()]
    return m.dropna(how="any") if len(m) >= line_floor else pd.DataFrame()


def gas_country_balance(bal, cc, storage, lng, cons_override=None, norway_to=None, biomethane=None, extra_imports=None,
                      extra_exports=None, prod_adjust=None, prod_override=None, import_floor=None, export_floor=None, until=None):
    """Monthly TWh gas balance for one ENTSOG country: production, pipeline imports, LNG send-out (ALSI) and storage
    withdrawals (AGSI+) as supply; pipeline exports and storage injections as negatives; consumption (distribution +
    final consumers) as a line. Supply less the negatives should land near the consumption line; the gap is the
    unreported or unclassified flow."""
    a = "GB" if cc == "UK" else cc
    day = pd.DataFrame(index=bal.index)
    day["Production"] = _col(bal, f"{cc}_production_GWhd").fillna(0)
    day["Pipeline imports"] = _col(bal, f"{cc}_imports_GWhd")
    lng_s = _col(lng, f"{a}_sendout_GWhd").reindex(bal.index)
    day["LNG send-out"] = lng_s.where(lng_s.notna(), _col(bal, f"{cc}_lng_GWhd")).fillna(0)
    out_s, in_s = _col(storage, f"{a}_withdrawal_GWhd").reindex(bal.index), _col(storage, f"{a}_injection_GWhd").reindex(bal.index)
    day["Storage withdrawals"] = out_s.where(out_s.notna(), _col(bal, f"{cc}_storage_out_GWhd")).fillna(0)
    day["Pipeline exports"] = -_col(bal, f"{cc}_exports_GWhd").fillna(0)
    day["Storage injections"] = -in_s.where(in_s.notna(), _col(bal, f"{cc}_storage_in_GWhd")).fillna(0)
    own = pd.concat([_col(bal, f"{cc}_distribution_GWhd"), _col(bal, f"{cc}_final_consumers_GWhd")], axis=1).sum(axis=1, min_count=1)
    day["Consumption"] = cons_override.reindex(bal.index).combine_first(own) if cons_override is not None else own
    if import_floor is not None:        # the country's own TSO border-entry series, where it exceeds ENTSOG's (points ENTSOG lacks or under-reports)
        day["Pipeline imports"] = pd.concat([day["Pipeline imports"], import_floor.reindex(bal.index)], axis=1).max(axis=1)
    if export_floor is not None:
        day["Pipeline exports"] = -pd.concat([-day["Pipeline exports"], export_floor.reindex(bal.index)], axis=1).max(axis=1)
    if norway_to is not None:
        # ENTSOG's pipeline imports for Great Britain are the St Fergus + Easington entry points, which carry UK North Sea gas as well
        # as Norwegian gas. Gassco's flows to Great Britain become the pipeline-import line; the rest of those terminals is UK
        # production (supply total unchanged). Days Gassco lacks take that month's mean (months with none keep ENTSOG's mixed split).
        nrw = norway_to.reindex(bal.index)
        nrw = nrw.fillna(nrw.groupby([nrw.index.year, nrw.index.month]).transform("mean"))
        ok = nrw.notna() & day["Pipeline imports"].notna()
        day.loc[ok, "Production"] = day["Production"] + day["Pipeline imports"] - nrw
        day.loc[ok, "Pipeline imports"] = nrw
    if extra_imports is not None:       # Norwegian gas ENTSOG does not report (Emden): added to pipeline imports
        day["Pipeline imports"] = day["Pipeline imports"] + extra_imports.reindex(bal.index).fillna(0)
    if extra_exports is not None:       # flows the ENTSOG classification drops (e.g. GB -> Ireland at Moffat): added to pipeline exports
        day["Pipeline exports"] = day["Pipeline exports"] - extra_exports.reindex(bal.index).fillna(0)
    if prod_adjust is not None:         # gas that leaves the grid and re-enters as 'production' (Hungary's blending): taken off production
        adj = prod_adjust.reindex(bal.index).fillna(0).where(day["Production"] > 0, 0.0)    # only on days the production entry is reported
        day["Production"] = (day["Production"] - adj).clip(lower=0)
    if prod_override is not None:       # national production statistic (Netherlands, CBS) in place of ENTSOG's production entries
        po = prod_override.reindex(bal.index)
        day["Production"] = po.where(po.notna(), day["Production"])
    cols = GAS_BAL_COLS
    if biomethane is not None and biomethane.notna().any():
        day["Biomethane"] = biomethane.reindex(bal.index).fillna(0)
        cols = GAS_BAL_COLS[:-1] + ["Biomethane", "Consumption"]
    day = day.dropna(subset=["Pipeline imports", "Consumption"])
    day = day[day.index >= "2021-10-01"]
    if until is not None:               # a correcting feed that ends (Elering's Karksi flow): later days are not balanced
        day = day[day.index <= until]
    return _monthly_twh(day[cols]) if len(day) else pd.DataFrame()


POINT_FIXES_FILE = "entsog_point_fixes_daily.xlsx"


def gb_site_storage(data_dir, nts_sto):
    """GB storage flows, GWh/d, from the National Gas Data Portal's site-level data (GB_STORAGE_SITES_DAILY.py): the nine sites'
    summed outflow and inflow. The NTS aggregate (`nts_sto`) overstates net withdrawals by about 5 TWh a year (it implies a net
    withdrawal in 2022-23 when the portal's stock rose 11 TWh), so it is only the last resort. Days before the site items start
    (Oct 2024) use the day-to-day change in the portal's total stock level (net: a fall is a withdrawal, a rise an injection)."""
    st = _sheet_or_empty(os.path.join(data_dir, "gb_storage_sites_daily.xlsx"), "Daily", "date")
    if not len(st) or "portal_total_stock" not in st:
        return nts_sto
    idx = nts_sto.index.union(st.index)
    wd, inj = st["total_outflow"].reindex(idx), st["total_inflow"].reindex(idx)
    d = -st["portal_total_stock"].reindex(idx).diff()   # stock labelled d is the end-of-day stock, so d - (d-1) is day d's net
    wd = wd.where(wd.notna(), d.clip(lower=0))
    inj = inj.where(inj.notna(), (-d).clip(lower=0))
    out = pd.DataFrame({"GB_withdrawal_GWhd": wd, "GB_injection_GWhd": inj})
    for c, n in (("GB_withdrawal_GWhd", "GB_withdrawal_GWhd"), ("GB_injection_GWhd", "GB_injection_GWhd")):
        out[c] = out[c].where(out[c].notna(), nts_sto[n].reindex(idx))
    return out


NL_CBS_GAS_FILE = "netherlands_cbs_gas_monthly.xlsx"


def nl_cbs_gas(data_dir, bal):
    """Dutch production and total consumption from CBS StatLine 86103NED (NETHERLANDS_CBS_GAS.py), GWh/d on the days of `bal`, or None.
    Each CBS month keeps its official total and takes its day-to-day shape from ENTSOG's Dutch production entries / consumption exits
    (daily_shape.reshape; an even spread where ENTSOG has no usable shape). Days after CBS's last month take ENTSOG's value (production entries; distribution +
    final-consumer exits) times the CBS/ENTSOG ratio of the last twelve CBS months, so the series carries on without a step."""
    mon = _sheet_or_empty(os.path.join(data_dir, NL_CBS_GAS_FILE), "Monthly", "month")
    if not len(mon) or not {"Production_GWh", "Consumption_GWh"} <= set(mon.columns):
        return None
    out = pd.DataFrame(index=bal.index)
    ent = {"Production": _col(bal, "NL_production_GWhd"),
           "Consumption": pd.concat([_col(bal, "NL_distribution_GWhd"), _col(bal, "NL_final_consumers_GWhd")], axis=1).sum(axis=1, min_count=1)}
    end = mon.index.max() + pd.offsets.MonthEnd(0)
    for k, col in (("Production", "Production_GWh"), ("Consumption", "Consumption_GWh")):
        m = mon[col].dropna()
        e = ent[k]
        per_day = daily_shape.reshape(m, e, None, label=f"CBS NL gas {k}")   # official CBS month total, ENTSOG's daily shape
        cbs = per_day.reindex(bal.index)
        last12 = per_day.index[per_day.index > end - pd.DateOffset(months=12)]
        both = pd.DataFrame({"c": per_day.reindex(last12), "e": e.reindex(last12)}).dropna()
        ratio = both["c"].sum() / both["e"].sum() if len(both) > 90 and both["e"].sum() > 0 else float("nan")
        tail = (e * ratio).where(bal.index > end)
        out[k] = cbs.combine_first(tail)
    if "Imports_via_Denmark_GWh" in mon:   # Danish North Sea gas landed at Den Helder (NOGAT): in no ENTSOG border row and no other country's balance
        m = mon["Imports_via_Denmark_GWh"].dropna()
        per_day = (m / m.index.days_in_month).reindex(pd.date_range(m.index.min(), m.index.max() + pd.offsets.MonthEnd(0), freq="D"), method="ffill")
        tail = pd.Series(per_day[per_day.index > end - pd.DateOffset(months=12)].mean(), index=bal.index).where(bal.index > end)
        out["Via Denmark"] = per_day.reindex(bal.index).combine_first(tail)
    return out


def nord_stream(data_dir):
    """Nord Stream 1 gas entering Germany at Greifswald (NEL + OPAL entries), GWh/d, or None. ENTSOG lists no far side for those points
    (ENTSOG_POINT_FIXES_DAILY.py), so the main pull drops them: 620 TWh in 2021 and 314 TWh in 2022 of Russian pipeline gas."""
    fx = _sheet_or_empty(os.path.join(data_dir, POINT_FIXES_FILE), "Daily", "date")
    cols = [c for c in ("DE_greifswald_nel", "DE_greifswald_opal") if c in fx]
    return fx[cols].sum(axis=1, min_count=1) if cols else None


def emden_entsog(data_dir):
    """Norwegian gas entering at Emden (EPT1), GWh/d, from ENTSOG's own entry points: {"DE": OGE + GUD, "NL": GTS} or None. ENTSOG does
    publish Emden (ENTSOG_POINT_FIXES_DAILY.py) but, with no far side listed, the country classification drops it. Thyssengas reports the
    same flow as GUD, so it is left out. 2022: 345 TWh against 513 from the Gassco-based emden_gap, which runs too high in 2021-22."""
    fx = _sheet_or_empty(os.path.join(data_dir, POINT_FIXES_FILE), "Daily", "date")
    if not {"DE_emden_oge", "DE_emden_gud", "NL_emden_gts"} <= set(fx.columns):
        return None
    return {"DE": fx[["DE_emden_oge", "DE_emden_gud"]].sum(axis=1, min_count=1), "NL": fx["NL_emden_gts"]}


def lu_imports(bal):
    """Luxembourg's pipeline imports from Belgium (GWh/d): the larger of ENTSOG's Bras-Petange entry and Luxembourg's own delivered volumes."""
    cons = pd.concat([bal["LU_distribution_GWhd"], bal["LU_final_consumers_GWhd"]], axis=1).sum(axis=1, min_count=1)
    return pd.concat([bal["LU_imports_GWhd"], cons], axis=1).max(axis=1)


ELERING_FILE = "elering_gas_daily.xlsx"


def estonia_gas_balance(data_dir):
    """Estonia's gas balance (monthly TWh) from Elering's own flows (ELERING_GAS_DAILY.py): imports and exports at Balticconnector (Finland) and
    Karksi (Latvia) against the gas delivered to Estonian consumers. ENTSOG has no Karksi row, so its Estonian balance missed the Latvian gas
    (Estonian consumption plus the Balticconnector exports to Finland) and ran -267% to +52%. Elering stopped publishing Karksi after Nov 2025, so
    later days are left out."""
    d = _sheet_or_empty(os.path.join(data_dir, ELERING_FILE), "Daily", "date")
    if not len(d) or "EE_karksi" not in d:
        return pd.DataFrame()
    pts = [c for c in ("EE_balticconnector", "EE_karksi", "EE_narva", "EE_varska") if c in d]
    f = d[pts].fillna(0.0)
    day = pd.DataFrame(index=d.index)
    day["Production"] = 0.0
    day["Pipeline imports"] = f.clip(lower=0).sum(axis=1)
    day["LNG send-out"] = 0.0
    day["Storage withdrawals"] = 0.0
    day["Pipeline exports"] = f.clip(upper=0).sum(axis=1)
    day["Storage injections"] = 0.0
    day["Consumption"] = d["EE_consumption"]
    day = day[d["EE_karksi"].notna()].dropna(subset=["Consumption"])
    day = day[day.index >= "2021-10-01"]
    return _monthly_twh(day) if len(day) else pd.DataFrame()


def lv_storage_from_stock(data_dir):
    """Latvia's (Inčukalns) storage flows in GWh/d from the day-to-day change in AGSI+'s stock level: AGSI's own injection/withdrawal columns
    disagree with the stock by 8.5 TWh in 2023 and 6.3 in 2024 (the stock is the reliable series)."""
    s = _sheet_or_empty(os.path.join(data_dir, "eu_gas_storage_daily.xlsx"), "Daily", "date")
    if "LV_TWh" not in s:
        return None
    dst = s["LV_TWh"].diff() * 1000.0
    return pd.DataFrame({"LV_withdrawal_GWhd": (-dst).clip(lower=0), "LV_injection_GWhd": dst.clip(lower=0)})


def it_storage_from_stock(data_dir):
    """Italy's storage flows in GWh/d from the day-to-day change in AGSI+'s stock level: AGSI's injection/withdrawal columns imply 28.6 TWh more net
    injection than the stock rose in 2025 (3.0 and 3.5 TWh in 2023-24), which made the 2025 Italian balance 10 TWh short against +2-3% in other years."""
    s = _sheet_or_empty(os.path.join(data_dir, "eu_gas_storage_daily.xlsx"), "Daily", "date")
    if "IT_TWh" not in s:
        return None
    dst = s["IT_TWh"].diff() * 1000.0
    return pd.DataFrame({"IT_withdrawal_GWhd": (-dst).clip(lower=0), "IT_injection_GWhd": dst.clip(lower=0)})


def point_fix_args(data_dir, cc, tso, bio, gni, bal_nl=None):
    """Country-specific corrections from ENTSOG points the main pull's classification drops (ENTSOG_POINT_FIXES_DAILY.py), as
    (keyword arguments for gas_country_balance, consumption override or None, note text or None).
    GR: TAP's Nea Mesimvria entry (Azerbaijani gas) is added to pipeline imports. HU: the 'Exit for Blending' is taken off production
    (imported gas blended with high-CO2 domestic gas re-enters at the production entry). UK: the Moffat exit is added to exports - the
    Republic of Ireland's share is GNI's own Moffat import figure, the rest (Northern Ireland, Isle of Man) is UK consumption that the
    National Gas NTS offtake series lacks. FR: ODRE consumption is the GRTgaz/Teréga offtake (it equals ENTSOG's distribution plus
    industrial exits), which excludes biomethane injected straight into the distribution networks, so that biomethane is added to
    consumption (it is also a supply line)."""
    if cc == "IT" and it_storage_from_stock(data_dir) is not None:
        return {"storage": it_storage_from_stock(data_dir)}, None, (
            " Storage flows are the day-to-day change in AGSI+'s stock level: AGSI's own injection and withdrawal columns for Italy disagree with the stock by 28.6 TWh in 2025.")
    fx = _sheet_or_empty(os.path.join(data_dir, POINT_FIXES_FILE), "Daily", "date")
    if cc == "NL":
        nl = nl_cbs_gas(data_dir, bal_nl)
        if nl is not None:
            em_raw = emden_entsog(data_dir)
            nl_extra = nl["Via Denmark"] if "Via Denmark" in nl else None
            if em_raw:
                nl_extra = em_raw["NL"].add(nl_extra.reindex(em_raw["NL"].index), fill_value=0) if nl_extra is not None else em_raw["NL"]
            return ({"prod_override": nl["Production"]} | ({"extra_imports": nl_extra} if nl_extra is not None else {})), nl["Consumption"], (
                " Production and consumption are Statistics Netherlands' (CBS 86103NED) gas balance, spread over the days of each month, in place of ENTSOG's "
                "production entries (15-17 TWh a year above CBS) and distribution + final-consumer exits (about 6 TWh below CBS total consumption); "
                "months CBS has not yet published use ENTSOG scaled by the last twelve months' CBS/ENTSOG ratio. Pipeline imports add CBS's 'via Denmark' gas "
                "(Danish North Sea gas landed at Den Helder, 8-14 TWh a year; ENTSOG has no border row for it). The balance still runs a few percent short on the "
                "Dutch side because ENTSOG reports physical net flows (NL>DE about 10 TWh a year more net export than CBS's commercial figures; the mirror is Germany's surplus), "
                "ALSI LNG send-out is 4-8 TWh a year below CBS's net LNG imports, and CBS bunkering (5-6 TWh) is not in its consumption.")
    if cc == "FR" and tso is not None and "FR" in tso and len(bio) and "FR" in bio:
        return {}, tso["FR"].add(bio["FR"].reindex(tso.index).fillna(0)), (
            " Consumption is ODRE's GRTgaz/Teréga offtake plus the biomethane injected into the distribution networks (ODRE's offtake equals "
            "ENTSOG's distribution + industrial exits and so excludes it). The balance still runs about 3% long: ENTSOG misses about 15 TWh of "
            "French exports against Eurostat, and network own use and losses are not in the offtake.")
    if cc == "LV":
        el = _sheet_or_empty(os.path.join(data_dir, ELERING_FILE), "Daily", "date")
        kw, note = {}, ""
        sto = lv_storage_from_stock(data_dir)
        if sto is not None:
            kw["storage"] = sto
            note += (" Storage flows are the day-to-day change in AGSI+'s Inčukalns stock level: AGSI's injection/withdrawal columns disagree with the stock by 8.5 TWh (2023) "
                     "and 6.3 TWh (2024).")
        if "EE_karksi" in el and bal_nl is not None:
            k = el["EE_karksi"]
            try:
                ent = _sheet_or_empty(os.path.join(data_dir, GAS_FLOWS_FILE), "Border flows", "date")["LV>EE"].reindex(k.index).fillna(0)
            except Exception:  # noqa: BLE001
                ent = pd.Series(0.0, index=k.index)
            kw["extra_exports"] = (k.clip(lower=0) - ent).clip(lower=0)      # Karksi is + from Latvia; ENTSOG's own LV>EE share is not counted twice
            kw["extra_imports"] = (-k).clip(lower=0)
            kw["until"] = k.dropna().index.max()
            note += (" Latvia's gas to Estonia (Karksi, Inčukalns gas for Estonian consumers and the Balticconnector) is Elering's own flow, which ENTSOG's Latvian exit list lacks; "
                     "days after Elering's last Karksi figure (Nov 2025) are left out.")
        return kw, None, note or None
    if cc in ("LU", "BE") and bal_nl is not None and "LU_imports_GWhd" in bal_nl:
        lu = lu_imports(bal_nl)
        if cc == "LU":
            return {"import_floor": lu}, None, (
                " Luxembourg's only ENTSOG entry is Bras-Petange from Belgium, which stops reporting flows in 2024 (0.0 TWh against 6.7 TWh of "
                "metered consumption), so pipeline imports are the larger of that entry and Creos's own delivered volumes (ENTSOG distribution + industrial exits); "
                "with no production or storage the balance then closes by construction.")
        return {"extra_exports": lu}, None, (
            " Pipeline exports add the gas delivered to Luxembourg at Bras-Petange (Luxembourg's imports, as above), which ENTSOG's Belgian exit list lacks "
            "(it was the whole of Belgium's 2024-25 surplus, 6-7 TWh).")
    if cc == "CZ":
        cee = _sheet_or_empty(os.path.join(data_dir, "europe_tso_gas_demand_cee_daily.xlsx"), "Daily", "date")
        if "CZ_border_entry" in cee:
            return {"import_floor": cee["CZ_border_entry"], "export_floor": cee["CZ_border_exit"]}, None, (
                " Pipeline imports and exports are the larger of ENTSOG's Czech-side figures and NET4GAS's own allocated border entries/exits (Brandov, Waidhaus, Lanzhot, Cesky Tesin): "
                "in 2023 ENTSOG's VIP Brandov was reported at 14 TWh and its physical points (EUGAL, OPAL, Hora Svate Katerina, Olbernhau) sum to only 63 TWh "
                "against NET4GAS's 79 TWh, which left the balance 23% short. Consumption is NET4GAS's own system balance, so the balance closes largely by construction.")
    if not len(fx):
        return {}, None, None
    if cc == "DE" and nord_stream(data_dir) is not None:
        em_raw = emden_entsog(data_dir)
        ns = nord_stream(data_dir)
        return {"extra_imports": ns.add(em_raw["DE"].reindex(ns.index).fillna(0), fill_value=0).combine_first(em_raw["DE"]) if em_raw else ns}, None, (
            (" Pipeline imports include Norwegian gas at Emden (EPT1: OGE and GUD entries; the Dutch share, GTS, is in the Netherlands balance), which ENTSOG's country classification drops." if em_raw else "") +
            " Pipeline imports include the Nord Stream 1 gas entering at Greifswald (NEL and OPAL entries, Russian origin, to Sept 2022), which "
            "ENTSOG's country classification drops because it lists no far side for those points (Germany + Netherlands was about 100 TWh a quarter short before).")
    # Greece: no TAP point fix any more - the rebuilt border-flow pull books the Nea Mesimvria entry as AL>GR (11.1 TWh in 2025), so adding
    # GR_tap_imports on top counted it twice (balance +16%); the pull's own AL>GR is in GR_imports.
    if cc == "HU" and "HU_production_exit" in fx:
        return {"prod_adjust": fx["HU_production_exit"]}, None, (
            " Production is ENTSOG's 'Aggregated Single Production' entry less the 'Exit for Blending': imported gas leaves the grid, is blended "
            "with high-CO2 domestic gas and re-enters at the production entry, so that entry double-counts about 14 TWh a year of imports.")
    if cc == "UK" and "UK_moffat_exit" in fx and tso is not None and "UK" in tso:
        mof = fx["UK_moffat_exit"]
        roi = gni["Moffat"].reindex(mof.index) if gni is not None and len(gni) else pd.Series(float("nan"), index=mof.index)
        to_roi = roi.where(roi.notna(), mof).clip(upper=mof)
        try:    # ENTSOG's UK exports already hold the Carrickfergus exit (Northern Ireland -> Ireland, 7-9 TWh a year); it is part of the Twynholm take-off
            ukie = _sheet_or_empty(os.path.join(data_dir, GAS_FLOWS_FILE), "Border flows", "date")["UK>IE"].reindex(mof.index).fillna(0)
        except Exception:  # noqa: BLE001
            ukie = pd.Series(0.0, index=mof.index)
        try:    # National Gas NTS storage flows (the operator's own; ENTSOG lacks the Stublach, Holford and Hill Top entries)
            nts = add_charts._sheet(os.path.join(data_dir, "gb_gas_nts_daily.xlsx"), "Daily", "date")
            sto = pd.DataFrame({"GB_withdrawal_GWhd": nts["storage_withdrawal"], "GB_injection_GWhd": nts["storage_injection"]})
            sto = gb_site_storage(data_dir, sto)
        except Exception:  # noqa: BLE001
            sto = None
        return ({"extra_exports": to_roi} | ({"storage": sto} if sto is not None else {})), tso["UK"].add((mof - to_roi - ukie).clip(lower=0).fillna(0)), (
            " Exports include the Moffat exit to Ireland (the Republic's share is Gas Networks Ireland's Moffat import figure); the remainder of "
            "the Moffat flow (Northern Ireland and the Isle of Man, about 19 TWh a year, less the 7-9 TWh a year that Northern Ireland sends on to the Republic at Carrickfergus, which ENTSOG already counts in UK exports) is added to UK consumption because the National Gas NTS "
            "offtake covers Great Britain only. Storage withdrawals and injections are the nine storage sites' own daily flows from the National Gas Data Portal (the NTS aggregate overstates net withdrawals by about 5 TWh a year); before Oct 2024 they are the day-to-day change in the portal's total stock (net only). ENTSOG omits Moffat from its UK exports (the point's far side is listed as country UK).")
    return {}, None, None


BIO_FILE = "europe_biomethane_operators.xlsx"


BIO_STATS_FILE = "europe_biomethane_statistics.xlsx"
BIO_OPERATOR_COUNTRIES = ("FR", "DK", "NL", "AT")   # operator series; every other EU27 country comes from Eurostat's annual biomethane figures


def biomethane_daily(data_dir):
    """Biomethane injected into the gas grids, GWh/d per country. Operator series: France (ODRE) and Denmark (Energinet) daily; the
    Netherlands (CBS) and Austria (AGGM) monthly, spread evenly over the days. Every other EU27 country comes from Eurostat's annual
    'biogases blended into natural gas' figures (Germany 11.5 TWh in 2024, Italy 3.2, ...), spread evenly over the year and held at the
    last published year afterwards (labelled on the chart notes). Great Britain's biomethane is not added: the NTS offtake that is our UK
    consumption already excludes gas embedded in the distribution networks. ENTSOG's transmission-level production does not contain
    biomethane, but national consumption does, so it is a separate supply line (Denmark's is already inside the Energinet delivery
    figure, not added to demand)."""
    path = os.path.join(data_dir, BIO_FILE)
    dly, mon = _sheet_or_empty(path, "Daily", "date"), _sheet_or_empty(path, "Monthly", "month")
    stats = os.path.join(data_dir, BIO_STATS_FILE)
    smon, ann = _sheet_or_empty(stats, "Monthly", "month"), _sheet_or_empty(stats, "Annual", "year")
    out = pd.DataFrame()
    for cc in ("FR", "DK"):
        if f"{cc}_biomethane" in dly:
            out[cc] = dly[f"{cc}_biomethane"]
    monthly = {}
    if "NL_biomethane" in mon:
        monthly["NL"] = mon["NL_biomethane"]
    if "AT_biomethane_GWh" in smon:
        monthly["AT"] = smon["AT_biomethane_GWh"].dropna()
    end = pd.Timestamp.today().normalize()
    for cc, ser in monthly.items():
        ser = ser.dropna()
        days = pd.date_range(ser.index.min(), ser.index.max() + pd.offsets.MonthEnd(0), freq="D")
        out = out.reindex(out.index.union(days))
        out[cc] = (ser / ser.index.days_in_month).reindex(days, method="ffill")
    if len(ann):
        yrs = ann.index
        for col in ann.columns:
            cc = col.split("_")[0]
            if not col.endswith("_biomethane_TWh") or cc in BIO_OPERATOR_COUNTRIES or cc in ("EU27", "UK", "NO") or ann[col].fillna(0).sum() <= 0:
                continue
            ser = ann[col].dropna()
            ser.index = [d.year if hasattr(d, "year") else int(d) for d in ser.index]     # the sheet loader turns the year column into dates
            years = range(int(ser.index.min()), end.year + 1)
            vals = {y: (ser[y] if y in ser.index else ser.iloc[-1]) for y in years}     # held at the last published year
            days = pd.date_range(f"{min(years)}-01-01", end, freq="D")
            out = out.reindex(out.index.union(days))
            out[cc] = pd.Series([vals[d.year] * 1000.0 / (366 if d.is_leap_year else 365) for d in days], index=days)
    return out.sort_index()


DK_FILE = "denmark_energinet_gasflow_daily.xlsx"


def denmark_gas_balance(data_dir):
    """Denmark's gas balance from Energinet's own Gasflow dataset (monthly TWh). Supply: the North Sea entry (zero until Baltic Pipe started in
    Oct 2022, then mainly Norwegian gas for Baltic Pipe), Tyra (Danish fields, small until 2024), biomethane, storage withdrawals and
    imports from Germany. Uses: exports to Poland (Baltic Pipe), Sweden and Germany, storage injections. Consumption is gas delivered to
    Danish consumers (it already includes the biomethane, which is why biomethane is also a supply line). ENTSOG shows only about 15 of
    the roughly 250 GWh/d that pass through the North Sea entries."""
    d = add_charts._sheet(os.path.join(data_dir, DK_FILE), "Daily", "date")
    g = _col(d, "DK_germany")
    day = pd.DataFrame(index=d.index)
    day["North Sea entry (mainly Norwegian gas for Baltic Pipe)"] = _col(d, "DK_from_north_sea")
    day["Tyra (Danish fields)"] = _col(d, "DK_from_tyra")
    day["Biomethane"] = _col(d, "DK_biogas")
    day["Storage withdrawals"] = _col(d, "DK_storage").clip(lower=0)
    day["Imports from Germany"] = g.clip(lower=0)
    day["Exports to Poland"] = _col(d, "DK_to_poland").fillna(0)
    day["Exports to Sweden"] = _col(d, "DK_to_sweden").fillna(0)
    day["Exports to Germany"] = g.clip(upper=0)
    day["Storage injections"] = _col(d, "DK_storage").clip(upper=0)
    day["Consumption"] = -_col(d, "DK_to_denmark")
    day = day.dropna(subset=["Consumption", "Biomethane"])
    day = day[day.index >= "2021-10-01"]
    return _monthly_twh(day.fillna(0.0)) if len(day) else pd.DataFrame()


def austria_gas_balance(data_dir):
    """Austria's gas balance (monthly TWh) from AGGM's own market-area series (GAS_TSO_CEE_DAILY.py): domestic production (OMV, RAG), net border
    entry and exit, storage withdrawal and injection against the end-customer consumption AGGM determines from metering and allocations (the
    sum of its customer classes). ENTSOG gave 17-22% too little supply in 2023-24: it has no Austrian production (5 TWh a year), its Austrian-side
    border rows lack Baumgarten, and AGSI's Austrian storage flows (which include Haidach, fed from the German grid) differ from AGGM's by 5-13 TWh a year.
    The identity entry - exit + storage + production = consumption closes within 0.4% in 2023-26, which checks the AGGM series against each other."""
    d = _sheet_or_empty(os.path.join(data_dir, TSO_CEE_FILE), "Daily", "date")
    if not len(d) or "AT_net_entry" not in d:
        return pd.DataFrame()
    day = pd.DataFrame(index=d.index)
    day["Production"] = _col(d, "AT_production")
    day["Pipeline imports"] = _col(d, "AT_net_entry")
    day["LNG send-out"] = 0.0
    day["Storage withdrawals"] = _col(d, "AT_storage_withdrawal")
    day["Pipeline exports"] = -_col(d, "AT_net_exit")
    day["Storage injections"] = -_col(d, "AT_storage_injection")
    day["Consumption"] = _col(d, "AT_total")
    day = day.dropna()
    day = day[day.index >= "2021-10-01"]
    return _monthly_twh(day[GAS_BAL_COLS]) if len(day) else pd.DataFrame()


NORWAY_ENTRIES_FILE = "entsog_norway_entries_daily.xlsx"


def emden_gap(data_dir, nor):
    """Norwegian gas that reaches Germany through Emden, GWh/d: Gassco's flow to Germany minus the part ENTSOG reports at Dornum and minus the Danish North Sea entry (Baltic Pipe gas, which Gassco books under Germany).
    ENTSOG publishes nothing at Emden (EPT1, EPT2, NPT) under any indicator, so Germany's pipeline imports in the ENTSOG pull miss it.
    Gassco days that are missing take the month's mean."""
    ent = add_charts._sheet(os.path.join(data_dir, NORWAY_ENTRIES_FILE), "Daily", "date")
    de_cols = [c for c in ent.columns if c.startswith("DE_")]
    if not de_cols or "NO_to_DE" not in nor:
        return None
    gass = nor["NO_to_DE"]
    gass = gass.fillna(gass.groupby([gass.index.year, gass.index.month]).transform("mean"))
    seen = ent[de_cols].sum(axis=1, min_count=1).reindex(gass.index).fillna(0)
    # Since Baltic Pipe (Oct 2022) Gassco books the Norwegian gas bound for Denmark/Poland/Sweden under "Germany" too: a daily regression of
    # Gassco's Germany on Dornum and Energinet's North Sea entry gives a coefficient of 1.0 on the Danish entry (R2 0.66), none on "Other".
    # That gas never reaches Germany (Denmark's own balance counts it), so it is taken out of the Emden add.
    try:
        dk = add_charts._sheet(os.path.join(data_dir, DK_FILE), "Daily", "date")["DK_from_north_sea"]
    except Exception:  # noqa: BLE001
        dk = ent[[c for c in ent.columns if c.startswith("DK_")]].sum(axis=1, min_count=1) if any(c.startswith("DK_") for c in ent.columns) else None
    dk = dk.reindex(gass.index).clip(lower=0).fillna(0) if dk is not None else 0
    return (gass - seen - dk).clip(lower=0).dropna()


NORWAY_FILE = "norway_gassco_gas_flows_daily.xlsx"
NORWAY_SRC = ("Gassco (Norwegian gas flows by delivery destination, daily; mcm/d x 11.2 GWh per mcm)", "https://gassco.eu/")
NORWAY_COLS = {"NO_to_GB": "Great Britain", "NO_to_DE": "Germany", "NO_to_FR": "France", "NO_to_BE": "Belgium", "NO_other": "Other"}


def norway_monthly(n):
    """Gassco daily GWh by destination -> monthly TWh (months with >= 60% of the days, scaled to the full month)."""
    day = n[[c for c in NORWAY_COLS if c in n]].rename(columns=NORWAY_COLS)
    m = monthly_cover(day, 0.6).dropna(how="all") / 1000.0
    last = day.dropna(how="all").index.max()
    if len(day) and last < last + pd.offsets.MonthEnd(0):   # drop the month in progress
        m = m[m.index < last.to_period("M").to_timestamp()]
    return m.dropna(how="all").fillna(0.0)


TSO_GAS_FILE = "europe_tso_gas_demand_daily.xlsx"
TSO_CEE_FILE = "europe_tso_gas_demand_cee_daily.xlsx"
TSO_EXTRA_FILE = "europe_tso_gas_demand_extra_daily.xlsx"


def _sheet_or_empty(path, sheet, index):
    try:
        return add_charts._sheet(path, sheet, index)
    except Exception:  # noqa: BLE001
        return pd.DataFrame()


def tso_consumption(data_dir):
    """National gas consumption from the operators' own series, GWh/d: Germany (THE), France (ODRE), Spain (Enagas daily from 2023,
    Enagas monthly bulletin spread evenly over the days before), Denmark (Energinet), Portugal (REN), Austria (AGGM), Czechia
    (NET4GAS system balance: border + storage + production flows), Lithuania (Amber Grid), Finland (Gasgrid), Great Britain
    (National Gas NTS). ENTSOG's country totals capture only part of these (Germany reports final consumers as one aggregate, Spain
    has few demand points, AT/LT/FI/DK have no consumption points). Poland, Romania and Croatia are left on ENTSOG: the operators'
    own exits equal ENTSOG's, and the remaining gap to Eurostat is domestic production that never enters the grid (missing from
    ENTSOG production too, so their balances already close)."""
    d = add_charts._sheet(os.path.join(data_dir, TSO_GAS_FILE), "Daily", "date")
    cee = _sheet_or_empty(os.path.join(data_dir, TSO_CEE_FILE), "Daily", "date")
    ext = _sheet_or_empty(os.path.join(data_dir, TSO_EXTRA_FILE), "Daily", "date")
    out = pd.DataFrame({"DE": d.get("DE_total"), "FR": d.get("FR_total"), "ES": d.get("ES_total"), "DK": d.get("DK_total"),
                        "PT": d.get("PT_total")})
    for cc, src, col in (("AT", cee, "AT_total"), ("CZ", cee, "CZ_total"), ("LT", cee, "LT_total"), ("FI", ext, "FI_total")):
        if len(src) and col in src:
            out[cc] = src[col].reindex(out.index.union(src.index))
    mon = _sheet_or_empty(os.path.join(data_dir, TSO_EXTRA_FILE), "Monthly", "month")
    if len(mon) and "ES_national" in mon:   # Spain before the daily series starts: monthly bulletin spread evenly over the days
        days = pd.date_range(mon.index.min(), mon.index.max() + pd.offsets.MonthEnd(0), freq="D")
        per_day = (mon["ES_national"] / mon.index.days_in_month).reindex(days, method="ffill")
        out["ES"] = out["ES"].reindex(out.index.union(days)).combine_first(per_day)
    out = out.dropna(axis=1, how="all")
    try:   # Great Britain: National Gas NTS offtake (LDZ + power stations + industrial)
        g = add_charts._sheet(os.path.join(data_dir, "gb_gas_nts_daily.xlsx"), "Daily", "date")
        out["UK"] = g[["ldz_offtake", "powerstations", "industrial_offtake"]].sum(axis=1, min_count=3)
    except Exception:  # noqa: BLE001
        pass
    return out.sort_index()


def gni_daily(data_dir):
    """Gas Networks Ireland daily series (GWh/d): Corrib + Inch production, Moffat imports from Great Britain (Republic of
    Ireland share) and Republic of Ireland demand. ENTSOG's Irish country totals miss most of Moffat and the Irish demand."""
    path = os.path.join(data_dir, "ireland_gas_combined_daily.xlsx")
    sup, dem = add_charts._sheet(path, "Supply", "date"), add_charts._sheet(path, "Demand", "date")
    out = pd.DataFrame(index=sup.index)
    out["Production"] = sup[[c for c in ("Corrib_Production_GWh", "Inch_Production_GWh") if c in sup]].sum(axis=1, min_count=1)
    out["Moffat"] = sup["Moffat_Imports_GWh"]
    out["Consumption"] = dem["Total_ROI_GWh"].reindex(out.index)
    return out


def _eu_lines(f):
    """One country's monthly balance frame -> the standard EU lines (Denmark's Energinet frame has its own column names)."""
    out = pd.DataFrame(index=f.index)
    pick = lambda pre: [c for c in f.columns if str(c).startswith(pre)]   # noqa: E731
    prod = [c for c in ["Production"] + pick("North Sea") + pick("Tyra") if c in f]
    out["Production"] = f[prod].sum(axis=1, min_count=1) if prod else 0.0
    for k, cols in (("Pipeline imports", ["Pipeline imports"] + pick("Imports from")), ("Pipeline exports", ["Pipeline exports"] + pick("Exports to")),
                    ("LNG send-out", ["LNG send-out"]), ("Storage withdrawals", ["Storage withdrawals"]), ("Storage injections", ["Storage injections"]),
                    ("Biomethane", ["Biomethane"]), ("Consumption", ["Consumption"])):
        have = [c for c in cols if c in f]
        out[k] = f[have].sum(axis=1, min_count=1) if have else 0.0
    return out.fillna(0.0)


def implied_member_balance(gbal, cc, storage, sweden_from_dk=None):
    """Balance for an EU27 member with no consumption series of its own (Slovakia: ENTSOG has only its border rows; Sweden: none, the gas
    arrives from Denmark), so the EU27 sum still carries the neighbours' exports to it. Consumption is the net gas the country takes
    (imports less exports plus storage), so the balance closes by construction and adds nothing to the EU error; it understates the true consumption
    (Slovakia about 49 TWh against the 26 TWh ENTSOG shows) and is labelled as implied. Monthly TWh, or an empty frame."""
    if cc == "SE":
        if sweden_from_dk is None or not sweden_from_dk.notna().any():
            return pd.DataFrame()
        day = pd.DataFrame({"Pipeline imports": -sweden_from_dk})
    else:
        if f"{cc}_imports_GWhd" not in gbal:
            return pd.DataFrame()
        day = pd.DataFrame(index=gbal.index)
        day["Pipeline imports"] = _col(gbal, f"{cc}_imports_GWhd")
        day["Pipeline exports"] = -_col(gbal, f"{cc}_exports_GWhd").fillna(0)
        day["Storage withdrawals"] = _col(storage, f"{cc}_withdrawal_GWhd").reindex(gbal.index).fillna(0)
        day["Storage injections"] = -_col(storage, f"{cc}_injection_GWhd").reindex(gbal.index).fillna(0)
    day = day.dropna(subset=["Pipeline imports"]).fillna(0.0)
    day = day[day.index >= "2021-10-01"]
    day["Consumption"] = day.sum(axis=1)
    return _monthly_twh(day) if len(day) else pd.DataFrame()


def eu_gas_balance(frames, border=None, fallback=None, min_share=0.8):
    """EU27 gas balance as the SUM of the corrected country balances (`frames`: monthly TWh frames as charted per country, so every
    per-country fix - Emden / Nord Stream / Greifswald, Gassco Norway, CBS Netherlands, AGGM Austria, NET4GAS floors, Energinet Denmark, GNI Ireland,
    biomethane - is in the EU total). Intra-EU pipeline flows are taken out of the imports and exports lines (`border`: the larger-of-both-sides
    flow per border from the Border flows sheet, both ends in the EU27, capped at the smaller of the summed imports and exports); the net
    is untouched, so a border one side reports and the other does not shows up as the residual instead of being hidden. Countries with under
    `min_share` of the months are left out (returned as the second value), and only months every kept country has are shown."""
    lines = {}
    for c, f in frames.items():
        if f is not None and len(f):
            fb = (fallback or {}).get(c)   # months an operator series lacks (Austria Jan-Apr 2022) take the ENTSOG-based balance of that country
            if fb is not None and len(fb):
                f = pd.concat([_eu_lines(f), _eu_lines(fb[~fb.index.isin(f.index)])]).sort_index()
                lines[c] = f
            else:
                lines[c] = _eu_lines(f)
    span = pd.date_range("2021-10-01", max(f.index.max() for f in lines.values()), freq="MS")
    keep = {c: f for c, f in lines.items() if f.reindex(span).notna().any(axis=1).mean() >= min_share}
    left = [c for c in lines if c not in keep]
    for c, f in keep.items():   # a one- or two-month hole in a small country's feed (Greece Apr 2022, Luxembourg Sep-Oct 2023) is interpolated rather than dropping the EU month
        g = f.reindex(span)
        g = g.interpolate(limit=2, limit_area="inside")
        g = g.where(g.notna().any(axis=1), g.shift(12))   # a feed that stops (Estonia, Latvia after Oct 2025): same month a year earlier, counted as a gap-fill
        keep[c] = g.dropna(how="all")
    months = span[[all(m in f.index for f in keep.values()) for m in span]]
    tot = sum(f.reindex(months).fillna(0.0) for f in keep.values())
    if border is not None and len(border):
        pairs = [c for c in border.columns if ">" in c and all(x in EU27_GAS for x in c.split(">"))]
        intra = _monthly_twh(border[pairs].sum(axis=1, min_count=1).to_frame("x"))["x"].reindex(months).fillna(0.0) if pairs else 0.0
        intra = pd.concat([intra, tot["Pipeline imports"], -tot["Pipeline exports"]], axis=1).min(axis=1).clip(lower=0)
        tot["Pipeline imports"] = tot["Pipeline imports"] - intra
        tot["Pipeline exports"] = tot["Pipeline exports"] + intra
    cols = [c for c in GAS_BAL_COLS[:-1] + ["Biomethane", "Consumption"]]
    return tot[cols], left


def ireland_gas_balance(data_dir):
    """Republic of Ireland gas balance from Gas Networks Ireland's own supply and demand series (monthly TWh): Corrib and Inch
    production, Moffat imports from Great Britain (ROI share - gas in transit to Northern Ireland is excluded) against ROI
    demand. ENTSOG's Irish country totals miss most of Moffat (about 7 TWh a year against about 42 in GNI's figures), so the
    Irish balance uses GNI. Ireland has no LNG terminal in service and no storage."""
    path = os.path.join(data_dir, "ireland_gas_combined_daily.xlsx")
    sup = add_charts._sheet(path, "Supply", "date")
    dem = add_charts._sheet(path, "Demand", "date")
    day = pd.DataFrame(index=sup.index)
    day["Production"] = sup[[c for c in ("Corrib_Production_GWh", "Inch_Production_GWh") if c in sup]].sum(axis=1, min_count=1)
    day["Pipeline imports"] = sup["Moffat_Imports_GWh"]
    day["LNG send-out"] = 0.0
    day["Storage withdrawals"] = 0.0
    day["Pipeline exports"] = 0.0
    day["Storage injections"] = 0.0
    day["Consumption"] = dem["Total_ROI_GWh"].reindex(day.index)
    day = day.dropna(subset=["Pipeline imports", "Consumption"])
    day = day[day.index >= "2021-01-01"]
    return _monthly_twh(day[GAS_BAL_COLS]) if len(day) else pd.DataFrame()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(DATA_DIR, "europe_master.xlsx"))
    ap.add_argument("--data-dir", default=DATA_DIR)
    args = ap.parse_args()
    cfg = sys.modules[__name__]

    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    dash2 = wb.create_sheet("Dashboard - Power")
    used = {"Dashboard", "Dashboard - Power", "Dashboard - Capacity factors", "Sources"}
    sources = []

    gas = sam.collect(wb, DATASETS, args.data_dir, used, sources, cfg=cfg)
    parts = [sam.collect(wb, RAW_POWER_DATASETS, args.data_dir, used, sources, cfg=cfg),
             sam.collect(wb, PRICE_DATASETS, args.data_dir, used, sources, cfg=cfg),
             sam.collect(wb, CAPACITY_DATASETS, args.data_dir, used, sources, cfg=cfg),
             sam.collect(wb, HYDRO_DATASETS, args.data_dir, used, sources, cfg=cfg)]
    power = tuple(sum((p[i] for p in parts), []) for i in range(3))

    pos = 0
    gen_frames = {}
    total, notes = europe_generation(args.data_dir, gen_frames)
    if not total.empty:
        total_chart(wb, used, power, pos, total, notes, "Europe generation total data",
                    "Europe power generation by source (ENTSO-E countries)", "GWh per month", ENTSOE,
                    "Countries summed (only months all of them have):")
        pos += 1
        print("Europe generation total:", "; ".join(notes))
    else:
        power[2].append("Europe generation total: " + "; ".join(notes))
    cap, cap_notes = europe_capacity(args.data_dir)
    if not cap.empty:
        total_chart(wb, used, power, pos, cap, cap_notes, "Europe capacity total data",
                    "Europe installed generating capacity (ENTSO-E countries)", "GW installed",
                    ("ENTSO-E Transparency Platform: installed capacity per production type", "https://transparency.entsoe.eu/"),
                    "Countries summed (annual figures, 1 January; storage excluded):", date_format="%Y")
        pos += 1
        print("Europe capacity total:", "; ".join(cap_notes))

    cf_chart, cf_row, cf_missing, cf_countries = capacity_factors.add_capacity_factor_sheets(
        wb, used, sam.sheet_name, gen_frames,
        {name: os.path.join(args.data_dir, f"{slug}_power_capacity.xlsx") for _, (name, slug, _) in COUNTRIES.items()},
        sam.CHART_W, sam.CHART_H,
        "Generation (ENTSO-E) / installed capacity (ENTSO-E, annual)", "Europe")
    if cf_chart:
        power[0].insert(pos, cf_chart)
        power[1].insert(pos, cf_row)
    power[2].extend(cf_missing)
    # gas balance: EU27 first, then each country (ENTSOG flows + ALSI LNG + AGSI+ storage)
    try:
        gbal = add_charts.by_date(add_charts.read(os.path.join(args.data_dir, GAS_FLOWS_FILE), "Country balance"), "date")
        gorg = add_charts._sheet(os.path.join(args.data_dir, GAS_FLOWS_FILE), "Imports by origin", "date")
        gdst = add_charts._sheet(os.path.join(args.data_dir, GAS_FLOWS_FILE), "Exports by destination", "date")
        gsto = add_charts._sheet(os.path.join(args.data_dir, "eu_gas_storage_daily.xlsx"), "Daily", "date")
        glng = add_charts._sheet(os.path.join(args.data_dir, "eu_lng_terminals_daily.xlsx"), "Daily", "date")
    except Exception as e:  # noqa: BLE001
        gbal = pd.DataFrame()
        gas[2].append(f"gas balance charts need {GAS_FLOWS_FILE} ({type(e).__name__}: {e})")
    if len(gbal):
        try:
            gni = gni_daily(args.data_dir)
        except Exception as e:  # noqa: BLE001
            gni = None
            gas[2].append(f"EU gas balance without GNI Ireland ({type(e).__name__}: {e})")
        try:
            tso = tso_consumption(args.data_dir)
        except Exception as e:  # noqa: BLE001
            tso = None
            gas[2].append(f"EU gas balance without TSO consumption ({type(e).__name__}: {e})")
        try:
            nor = add_charts._sheet(os.path.join(args.data_dir, NORWAY_FILE), "Daily", "date")
        except Exception as e:  # noqa: BLE001
            nor = pd.DataFrame()
            gas[2].append(f"Norway/Gassco flows unavailable ({type(e).__name__}: {e})")
        try:
            bio = biomethane_daily(args.data_dir)
        except Exception as e:  # noqa: BLE001
            bio = pd.DataFrame()
            gas[2].append(f"gas balances without biomethane ({type(e).__name__}: {e})")
        try:
            emden = emden_gap(args.data_dir, nor) if len(nor) else None
        except Exception as e:  # noqa: BLE001
            emden = None
            gas[2].append(f"Germany gas balance without the Emden correction ({type(e).__name__}: {e})")
        eu_frames, eu_fallback = {}, {}
        for cc in [c for c in GAS_NAMES if any(col.startswith(f"{c}_") for col in gbal.columns)]:
            fix_kw, fix_cons, fix_note = point_fix_args(args.data_dir, cc, tso, bio, gni, gbal)
            cons_in = fix_cons if fix_cons is not None else (tso[cc] if (tso is not None and cc in tso) else None)
            sto_in = fix_kw.pop("storage", gsto)
            b = gas_country_balance(gbal, cc, sto_in, glng, cons_in,
                                    nor["NO_to_GB"] if (cc == "UK" and len(nor) and "NO_to_GB" in nor) else None,
                                    bio[cc] if (len(bio) and cc in bio) else None, **fix_kw)
            b_entsog = b
            note = None
            if tso is not None and cc in tso:
                note = ("Consumption from the TSO's own series (" + {"DE": "Trading Hub Europe", "FR": "ODRE / GRTgaz-Teréga", "ES": "Enagás", "UK": "National Gas NTS", "DK": "Energinet", "PT": "REN", "AT": "AGGM", "CZ": "NET4GAS (system balance)", "LT": "Amber Grid", "FI": "Gasgrid"}[cc]
                        + ") in place of ENTSOG's partial country total. Supply (production, pipeline imports, LNG, storage withdrawals) less exports "
                        "and storage injections against that consumption.")
            if len(bio) and cc in bio:
                note = (note or "Supply less exports and storage injections against consumption.") + (
                    " Biomethane injected into the grids (separate supply line) is added to supply: it is part of national consumption but "
                    "not of ENTSOG's transmission-level production" + ("; Denmark's consumption figure already includes it, so it is not added to demand"
                                                                      if cc == "DK" else "") + ".")
            if cc in ("DE", "NL") and emden is not None and emden_entsog(args.data_dir) is None:
                note = (note or "Supply less exports and storage injections against consumption.") + (
                    " Excludes the Norwegian gas that arrives at Emden: ENTSOG publishes nothing at Emden and the gas feeds both the German and "
                    "the Dutch grids, so this balance is short on its own; see the Germany + Netherlands balance.")
            if cc == "UK" and len(nor) and "NO_to_GB" in nor:
                note += (" Pipeline imports are Gassco's Norway-to-Great-Britain flows; the rest of the St Fergus and Easington entry points "
                         "(UK North Sea gas) is counted in production.")
            if fix_note:
                note = (note or "Supply less exports and storage injections against consumption.") + fix_note
            if cc == "EE":
                eb = estonia_gas_balance(args.data_dir)
                if len(eb):
                    b, note = eb, ("Estonia from Elering's own flows: imports and exports at Balticconnector (Finland) and Karksi (Latvia) against the gas delivered to "
                                   "Estonian consumers. ENTSOG has no Karksi row, so it missed the Latvian gas that supplies Estonia and the Balticconnector exports to Finland "
                                   "(balance -267% to +52%). Elering stopped publishing Karksi after Nov 2025, so later months are left out.")
                else:
                    gas[2].append("Estonia gas balance from Elering unavailable; ENTSOG used")
            if cc == "DK":
                try:
                    b, note = denmark_gas_balance(args.data_dir), (
                        "Denmark from Energinet's own Gasflow dataset: the North Sea entry (zero until Baltic Pipe started in Oct 2022, then "
                        "mainly Norwegian gas for Baltic Pipe) and Tyra (Danish fields, small until 2024), biomethane, storage and the German border against exports to Poland and "
                        "Sweden; consumption is gas delivered to Danish consumers (including the biomethane). ENTSOG captures only a small part "
                        "of these flows.")
                except Exception as e:  # noqa: BLE001
                    gas[2].append(f"Denmark gas balance from Energinet failed ({type(e).__name__}: {e}); ENTSOG used")
            if cc == "AT":
                ab = austria_gas_balance(args.data_dir)
                if len(ab):
                    b, note = ab, (
                        "Austria from AGGM's own market-area series: domestic production (OMV, RAG), net border entry and exit (all border points, "
                        "including Baumgarten), storage withdrawals and injections against the end-customer consumption AGGM determines from metering and "
                        "allocations. Biomethane is inside these series and not added. ENTSOG alone left the balance 12-22% short in 2022-24 (no Austrian "
                        "production, no Austrian-side Baumgarten row, and AGSI's storage flows differ from AGGM's); AGGM's identity closes within 0.4%. "
                        "Months with fewer than 90% of days are left out.")
                else:
                    gas[2].append("Austria gas balance from AGGM unavailable; ENTSOG used")
            if cc == "IE":
                try:
                    b, note = ireland_gas_balance(args.data_dir), (
                        "Republic of Ireland from Gas Networks Ireland's own series: Corrib/Inch production and Moffat imports "
                        "(from Great Britain, ROI share) against ROI demand. ENTSOG's Irish totals miss most of Moffat.")
                except Exception as e:  # noqa: BLE001
                    gas[2].append(f"Ireland gas balance from GNI failed ({type(e).__name__}: {e}); ENTSOG used")
            if cc in EU27_GAS:
                eu_frames[cc] = b
                eu_fallback[cc] = b_entsog
            if b.empty:
                gas[2].append(f"{GAS_NAMES[cc]} gas balance: too little data")
                continue
            total_chart(wb, used, gas, None, b, [note or ("Supply (production, pipeline imports, LNG, storage withdrawals) less exports and "
                                                 "storage injections should land near consumption (distribution + final consumers); "
                                                 "the gap is unreported or unclassified flow")],
                        f"{GAS_NAMES[cc]} gas balance data", f"{GAS_NAMES[cc]} gas balance: supply and storage vs consumption",
                        "TWh per month", GAS_BALANCE_SRC, "Notes:", label=GAS_NAMES[cc], line_cols=("Consumption",))
    if len(gbal):
        try:
            gbord = add_charts._sheet(os.path.join(args.data_dir, GAS_FLOWS_FILE), "Border flows", "date")
        except Exception:  # noqa: BLE001
            gbord = None
        try:   # Months a small country's feed lacks (Estonia and Latvia after Oct 2025) repeat the same month of the year before. Slovakia and Sweden have no consumption series: implied from the net gas they take (see implied_member_balance)
            eu_frames["SK"] = implied_member_balance(gbal, "SK", gsto)
            dk_raw = add_charts._sheet(os.path.join(args.data_dir, DK_FILE), "Daily", "date")
            eu_frames["SE"] = implied_member_balance(gbal, "SE", gsto, _col(dk_raw, "DK_to_sweden").reindex(gbal.index))
        except Exception as e:  # noqa: BLE001
            gas[2].append(f"EU gas balance without Slovakia/Sweden ({type(e).__name__}: {e})")
        eu, eu_left = eu_gas_balance(eu_frames, gbord, eu_fallback)
        if not eu.empty:
            total_chart(wb, used, gas, 0, eu, [
                "EU27: the sum of the country balances charted below, each with its own corrections (national consumption series from the TSOs and statistics offices, "
                "Gas Networks Ireland, Energinet, AGGM and CBS balances, NET4GAS floors, Norwegian gas at Emden / Greifswald and Hungary, Greece, Great Britain fixes, biomethane as a separate "
                "supply line). Pipeline imports and exports are those from/to outside the EU: flows between two EU countries (larger-of-both-sides border flows) are taken out of both lines, "
                "which leaves the net unchanged, so a border only one side reports remains in the residual. LNG and storage are the countries' own GIE ALSI / AGSI+ figures. "
                "Months a small country's feed lacks (Estonia and Latvia after Oct 2025) repeat the same month of the year before. Slovakia and Sweden have no consumption series: their consumption is the net gas ENTSOG / Energinet show them taking (implied, so they add nothing to the error and "
                "understate the true figure, Slovakia by about 20 TWh a year)."
                + (f" Left out for lack of data: {', '.join(GAS_NAMES.get(c, c) for c in eu_left)}." if eu_left else "")],
                        "EU gas balance data", "EU gas balance: supply and storage vs consumption (TWh per month)",
                        "TWh per month", GAS_BALANCE_SRC, "Notes:", label="Europe", line_cols=("Consumption",))
    if len(gbal) and emden is not None:
        try:   # Germany + Netherlands: Gassco's Emden gas is split between them in a way the raw data cannot show, so they are combined
            ns, em_raw = nord_stream(args.data_dir), emden_entsog(args.data_dir)
            em_de = em_raw["DE"] if em_raw else emden           # ENTSOG's own Emden entries where pulled, else the Gassco-based estimate
            de_b = gas_country_balance(gbal, "DE", gsto, glng, tso["DE"] if "DE" in tso else None, None,
                                       bio["DE"] if (len(bio) and "DE" in bio) else None,
                                       em_de.add(ns.reindex(em_de.index).fillna(0), fill_value=0).combine_first(ns) if ns is not None else em_de)
            nl_kw, nl_cons, _nl_note = point_fix_args(args.data_dir, "NL", tso, bio, gni, gbal)
            nl_b = gas_country_balance(gbal, "NL", gsto, glng, nl_cons, None, bio["NL"] if (len(bio) and "NL" in bio) else None, **nl_kw)
            colsb = [c for c in de_b.columns if c in nl_b.columns]
            both = de_b[colsb].add(nl_b[colsb], fill_value=0)
            if len(both) >= 12:
                total_chart(wb, used, gas, None, both, [
                    "Germany + Netherlands combined. Includes the Norwegian gas that arrives at Emden (Gassco's flow to Germany minus the Dornum "
                    "volume ENTSOG reports and minus the Baltic Pipe gas for Denmark/Poland that Gassco books under Germany); ENTSOG publishes nothing at Emden, and the gas feeds both grids, so the two countries are shown together. "
                    "Flows between the two countries are counted on both sides (ENTSOG's own-side figures differ by under 10 TWh a year). Dutch production and consumption are CBS StatLine 86103NED (national statistics) in place of ENTSOG's."],
                            "Germany Netherlands gas balance data", "Germany + Netherlands gas balance: supply and storage vs consumption",
                            "TWh per month", GAS_BALANCE_SRC, "Notes:", label="Germany + Netherlands", line_cols=("Consumption",))
        except Exception as e:  # noqa: BLE001
            gas[2].append(f"Germany + Netherlands balance failed ({type(e).__name__}: {e})")
    if len(gbal) and len(nor):
        nm = norway_monthly(nor)
        if len(nm) >= 12:
            total_chart(wb, used, gas, None, nm, ["Norwegian pipeline gas by delivery destination (Gassco). Great Britain (Langeled to Easington, "
                                                  "Vesterled/FLAGS to St Fergus) is not in ENTSOG's Norway origin flows. Daily values are Gassco's "
                                                  "published figures; months with under 60% of days are left out and gaps are scaled to the full month"],
                        "Norway gas exports data", "Norway gas exports by destination", "TWh per month", NORWAY_SRC, "Notes:", label="Norway")
    # supply/demand balance per country (needs the flows workbook and each country's load)
    flows_path = os.path.join(args.data_dir, FLOWS_FILE)
    try:
        net_all = add_charts.by_date(add_charts.read(flows_path, "Net imports"), "date")
    except Exception as e:  # noqa: BLE001
        net_all = pd.DataFrame()
        power[2].append(f"balance charts need {FLOWS_FILE} ({type(e).__name__}: {e})")
    bal_frames, bal_src = {}, {}
    inputs = [(name, os.path.join(args.data_dir, GEN_OVERRIDE.get(name, f"{slug}_power_generation_daily.xlsx")), net_all[name] if name in net_all else None)
              for _, (name, slug, _z) in COUNTRIES.items()]
    gb_path = os.path.join(args.data_dir, GB_FILE)
    try:   # Great Britain is not in ENTSO-E generation: Elexon/NESO workbook carries its own interconnector net imports (MWh -> GWh)
        gb_net = add_charts.by_date(add_charts.read(gb_path, "Daily"), "date")["NetImports_MWh"].apply(pd.to_numeric, errors="coerce") / 1000.0
        inputs.append(("Great Britain", gb_path, gb_net))
        bal_src["Great Britain"] = GB_BALANCE_SRC
    except Exception as e:  # noqa: BLE001
        power[2].append(f"Great Britain balance ({type(e).__name__}: {e})")
    bal_src["Switzerland"] = CH_BALANCE_SRC
    bal_src["Netherlands"] = NL_BALANCE_SRC
    try:   # Switzerland: BFE's physical imports less exports (monthly, spread over days) replace the ENTSO-E flow sum
        ch = add_charts.by_date(add_charts.read(os.path.join(args.data_dir, GEN_OVERRIDE["Switzerland"]), "Daily"), "date")["NetImports_MWh"]
        ch = pd.to_numeric(ch, errors="coerce") / 1000.0
        inputs = [(n, g, ch if n == "Switzerland" else x) for n, g, x in inputs]
    except Exception as e:  # noqa: BLE001
        power[2].append(f"Switzerland net imports from BFE not available, ENTSO-E flows used ({type(e).__name__}: {e})")
    for n_, (f_, c_) in NETIMP_EXTRA.items():   # e.g. Spain: Morocco + Andorra from REE (ENTSO-E has no zone for them)
        try:
            ex = pd.to_numeric(add_charts.by_date(add_charts.read(os.path.join(args.data_dir, f_), "Daily"), "date")[c_], errors="coerce") / 1000.0
            inputs = [(n, g, (x + ex.reindex(x.index)) if (n == n_ and x is not None) else x) for n, g, x in inputs]
            bal_src[n_] = ES_BALANCE_SRC
        except Exception as e:  # noqa: BLE001
            power[2].append(f"{n_} extra net imports not available, ENTSO-E flows only ({type(e).__name__}: {e})")
    pump_ov = {}
    for n_, f_ in PUMP_CONS_OVERRIDE.items():
        try:
            pump_ov[n_] = pd.to_numeric(add_charts.by_date(add_charts.read(os.path.join(args.data_dir, f_), "Daily"), "date")["PumpedStorageConsumption_MWh"], errors="coerce") / 1000.0
            bal_src[n_] = SI_BALANCE_SRC
        except Exception as e:  # noqa: BLE001
            power[2].append(f"{n_} pumping consumption not available, balance left without it ({type(e).__name__}: {e})")
    load_ov = {}
    for n_, f_ in LOAD_OVERRIDE.items():
        try:
            load_ov[n_] = pd.to_numeric(add_charts.by_date(add_charts.read(os.path.join(args.data_dir, f_), "Daily"), "date")["Load_MWh"], errors="coerce") / 1000.0
            bal_src[n_] = DK_BALANCE_SRC
        except Exception as e:  # noqa: BLE001
            power[2].append(f"{n_} load override not available, ENTSO-E load used ({type(e).__name__}: {e})")
    for name, gen_path, net in inputs:
        if net is None or name == "Ireland (all-island SEM)":   # Ireland comes from EirGrid + Ember below
            continue
        try:
            b = country_balance(gen_path, net, load_ov.get(name))
            if name in pump_ov and len(b):   # pumping consumption ENTSO-E lacks (GWh per month = sum of the days)
                b["Pumped & battery (net)"] = b["Pumped & battery (net)"] - pump_ov[name].resample("MS").sum().reindex(b.index).fillna(0.0)
        except Exception as e:  # noqa: BLE001
            power[2].append(f"{name} balance ({type(e).__name__}: {e})")
            continue
        if len(b) >= 3:
            bal_frames[name] = b
        else:
            power[2].append(f"{name} balance: fewer than 3 months with load and flows")
    try:   # Republic of Ireland: monthly fuel mix (Ember, then EirGrid) vs EirGrid demand; net imports = demand - generation
        mix, dem, _note = ireland_monthly(args.data_dir)
        b = mix.reindex(columns=FUELS).fillna(0.0).copy()
        b["Net imports"] = dem - b.sum(axis=1)
        b["Pumped & battery (net)"] = 0.0
        b["Load"] = dem
        b = b.dropna()[BALANCE_COLS].astype(float)
        b = b[b.index >= START]
        if len(b) >= 3:
            bal_frames["Ireland"] = b
            bal_src["Ireland"] = ("EirGrid system data (demand, wind, solar) and Ember (fuel split); net imports = EirGrid demand less generation",
                                  "https://www.eirgrid.ie/grid/system-and-renewable-data-reports")
    except Exception as e:  # noqa: BLE001
        power[2].append(f"Ireland balance ({type(e).__name__}: {e})")
    # Pseudo balancing item = load minus supply: positive = embedded/self-consumed generation the feed does not carry, negative = losses or
    # demand the supply side does not cover. It is implied, not measured, and makes the bars add up to load exactly.
    for name, b in list(bal_frames.items()):
        b2 = b.copy()
        supply = b2[["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Other", "Net imports", "Pumped & battery (net)"]].sum(axis=1)
        b2.insert(b2.columns.get_loc("Pumped & battery (net)"), BAL_ITEM, b2["Load"] - supply)
        bal_frames[name] = b2
    for name, b in bal_frames.items():
        src_label = {"Switzerland": "Swissgrid/BFE", "Netherlands": "CBS (flows ENTSO-E)", "Great Britain": "Elexon BMRS + NESO", "Denmark": "ENTSO-E + Energinet load", "Ireland": "EirGrid + Ember (net imports = demand - generation)"}.get(name, "ENTSO-E")
        total_chart(wb, used, power, None, b, [f"{name}: generation + net imports + pumped storage/batteries net vs load; "
                                               f"months with >= 75% of days (scaled to the month); {src_label}"]
                    + coverage_notes(name, b),
                    f"{name} balance data", f"{name} power balance: supply by source and net imports vs load",
                    "GWh per month", bal_src.get(name, BALANCE_SRC), "Notes:", label=name,
                    line_cols=("Pumped & battery (net)", "Load"))
    if len(bal_frames) >= 10:
        last_m = max(b.index.max() for b in bal_frames.values())
        core, skipped, skipped_notes = {}, [], []
        for n, b in bal_frames.items():
            span = pd.date_range(b.index.min(), last_m, freq="MS")
            if b.index.min() <= LATE_START_OK and len(b.index.intersection(span)) / len(span) >= MIN_SHARE:
                core[n] = b
            else:
                skipped.append(n)
        first_m = max([pd.Timestamp(START)] + [b.index.min() for b in core.values()])
        window = pd.date_range(first_m, last_m, freq="MS")
        late = [f"{n} from {b.index.min():%b/%y}" for n, b in core.items() if b.index.min() > pd.Timestamp(START) and b.index.min() == first_m]
        if late:
            skipped_notes.append(f"Totals start {first_m:%b/%y} because " + ", ".join(late)
                                 + f" (earlier months of the feed are missing or under {MIN_COVERAGE:.0%} of the country's typical share of load)")
        # a single month a small country lacks (e.g. Bosnia, Sep 2023: 13 days of load) is interpolated from its neighbours
        filled = {}
        for n, b in core.items():
            f = b.reindex(window).interpolate(limit=1, limit_area="inside")
            gaps = [f"{d:%b/%y}" for d in window if d not in b.index and f.loc[d].notna().all()]
            if gaps:
                skipped_notes.append(f"{n}: {', '.join(gaps)} interpolated from neighbouring months (too few days of data)")
            filled[n] = f
        core = filled
        months = [d for d in window if all(f.loc[d].notna().all() for f in core.values())] if core else []
        if months:
            eu = sum(b.loc[months] for b in core.values())
            eu.index.name = "date"
            total_chart(wb, used, power, pos, eu, [f"{len(core)} countries with load and flows: " + ", ".join(core)]
                        + [f"NOT INCLUDED: {n} (balance months since {pd.Timestamp(START):%b/%y}: too few)" for n in skipped] + skipped_notes,
                        "Europe balance data", "Europe power balance: supply by source and net imports vs load",
                        "GWh per month", BALANCE_SRC, "Countries summed (only months all of them have):",
                        line_cols=("Pumped & battery (net)", "Load"))
            pos += 1
    sam.draw_dashboard(dash, "Europe energy - gas dashboard", *gas)
    sam.draw_dashboard(dash2, "Europe energy - power generation, capacity and prices", *power)
    cf_dash = wb.create_sheet("Dashboard - Capacity factors", 2)
    sam.draw_dashboard(cf_dash, "Europe - capacity factor by generation type (generation / capacity x hours)",
                       [c for c, _ in cf_countries], [r for _, r in cf_countries], cf_missing)

    # long-term fundamentals: annual history (Ember yearly, EI Statistical Review, World Bank, IMF, degree days)
    fundamentals.add_long_term_dashboard(wb, used, sam.sheet_name, "Europe", args.data_dir, sam.CHART_W, sam.CHART_H,
                                         rows_per_chart=sam.ROWS_PER_CHART, degree_days=("HDD_18", "CDD_18"),
                                         index=sum(s.startswith("Dashboard") for s in wb.sheetnames))

    src = wb.create_sheet("Sources")
    src.append(["Country", "Dataset", "Workbook", "Publisher", "Link", "Units and notes (from the source workbook)"])
    for cell in src[1]:
        cell.font = Font(bold=True)
    for s in sources:
        src.append(list(s))
        if s[4]:
            src.cell(row=src.max_row, column=5).hyperlink = s[4]
    for col, w in (("A", 24), ("B", 16), ("C", 44), ("D", 60), ("E", 50), ("F", 160)):
        src.column_dimensions[col].width = w

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_charts.save_atomic(wb, args.out)
    print(f"Saved {args.out}: {len(gas[0])} gas charts, {len(power[0])} power charts; {len(wb.sheetnames)} tabs")
    for label, m in (("gas", gas[2]), ("power", power[2])):
        if m:
            print(f"missing ({label}):", m)


if __name__ == "__main__":
    main()
