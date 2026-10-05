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
    "Italy": "Net imports include the Sicily-Malta cable (exports of 0.5-1.0 TWh a year, ENTSO-E zone MT), added in Oct 2026 (it lowered supply/load by 0.2-0.3 points). ENTSO-E generation (about 215-220 TWh in 2024) omits embedded and self-consumed generation (Eurostat/Terna count 263.5 TWh) and ENTSO-E load (273 TWh) omits the matching demand (Terna 312 TWh); supply is 95-98% of load. Eurostat checks are in italy_eurostat_power_daily.xlsx but are not used here.",
    "Great Britain": (
        "Load now includes station load (power stations' own use, 4.7-4.8 TWh a year = NESO TSD - ND - pumping - interconnector exports, half-hourly): Elexon FUELHH metered output is gross of it while "
        "NESO national demand excludes it, which had left supply 2-3% above load (8 TWh in 2025). Supply is now 100.2-101.2% of load in 2021-25. "
        "Interconnector audit (discovery_archive/europe/POWER_INTERCONNECTORS_AUDIT.md): every GB cable (IFA, IFA2, ElecLink, Nemo, BritNed, Viking, NSL, Moyle, EWIC, Greenlink) is in Elexon's net imports, "
        "and ENTSO-E's mirror flows (France, Netherlands, Belgium, Denmark, Norway, Ireland) agree with Elexon within 0.5 TWh a year (Elexon is metered at the GB end, ENTSO-E at the far end: the 2-6% difference on Nemo, BritNed, NSL is cable losses), "
        "except Greenlink, which ENTSO-E does not report until June 2025 (about 0.9 TWh of Feb-May 2025 flow is missing from the SEM-GB border). Moyle's flow to Northern Ireland is in no covered country's balance."),
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
    "Slovenia": (
        "Supply was 103.5-103.9% of load because ENTSO-E publishes the Avce pumped-storage plant's generation (0.30 / 0.28 / 0.32 TWh in 2023 / 2024 / 2025) but no pumping consumption (zero in every year), "
        "while its load excludes pumping. The Statistical Office of Slovenia (SiStat 1817602S) gives 0.405 / 0.381 / 0.437 TWh used for pumped storage (a constant 1.346 x the pumped output, i.e. an assumed 74% round trip, "
        "close to the 1.35-1.37 ENTSO-E measures at Kruonis and Coo); it is subtracted from 'Pumped & battery (net)' (SLOVENIA_SISTAT_PUMPING.py), which leaves supply at 100.3-100.5% of load. "
        "Krsko is not a cause: ENTSO-E counts the whole plant (5.3-5.6 TWh, SiStat net nuclear 5.33 / 5.55 TWh) in Slovenia and the physical flow to Croatia already carries Croatia's half. "
        "Generation (14.2 TWh in 2023) and load (12.2) are both about 1 TWh below SiStat's net production (15.1) and final consumption plus network losses (13.2): solar behind the meter (SiStat 0.98 TWh, ENTSO-E 0.31) is in neither."),
    "Serbia": (
        "Supply is 103.5-104.0% of load. ENTSO-E publishes the pumped-storage generation of Bajina Basta (0.64 / 0.36 / 0.46 TWh in 2023 / 2024 / 2025) but no pumping consumption (zero in every year), while its load excludes pumping. "
        "At the 1.35 consumption-to-output ratio ENTSO-E reports for Kruonis and Coo this is 0.86 / 0.49 / 0.62 TWh and would bring supply to 100.9 / 102.6 / 102.2% of load. The remaining 0.3-0.9 TWh (1-2.6%) is not explained: "
        "EMS publishes no consumption or balance data (ems.rs carries market rules only; checked), so there is no raw pumping figure or national load to put in its place and the ENTSO-E values are kept. Treat Serbia as indicative."),
    "Lithuania": (
        "Supply is 103.9 / 104.0 / 103.1% of load in 2023 / 2024 / 2025, a near-constant 0.36-0.49 TWh a year (1.0-1.3 GWh a day, 0.07-0.18 TWh in every quarter 2023-26, unrelated to load, wind or solar), so it is not caused by the Baltic "
        "synchronisation of February 2025 (0.46 TWh in 2023, 0.36 in 2025). Nor is it Kruonis (pumping consumption is reported: 0.73 / 0.78 / 0.64 TWh, 1.33-1.38 x its output) or the Belarus flows (1.3 TWh imported in 2023, 0.5 in 2024, none in 2025, "
        "against a gap that does not follow them). Litgrid publishes no machine-readable balance or loss data (its data pages are viewer pages; Statistics Lithuania blocks GitHub), so the cause is not verified; the size is of the order of "
        "transmission losses, which an ENTSO-E load definition may leave out. Before 2022 imports from Belarus/Russia are not in the ENTSO-E flow data used here."),
    "Belgium": (
        "Generation, load and flows are Elia's own open data (ods201, ods001, ods026): 2025 generation 68.81 TWh, total load 80.20, net import 14.09, identical to ENTSO-E, so the surplus sits inside Elia's data. "
        "Generation plus net imports exceeds Elia's total load by 3.7 / 2.2 / 2.8 TWh in 2023 / 2024 / 2025 before pumping; ENTSO-E's Coo pumping consumption (1.66 / 1.41 / 1.18 TWh) takes it to 2.1 / 0.8 / 1.5 TWh "
        "(102.7 / 101.0 / 101.9% of load). Of the 2025 figure about 0.3 TWh is battery charging: Elia counts the batteries' discharge (0.29 TWh) in generation but publishes no charging. April-September 2023 carries an extra 2 TWh "
        "(0.5-0.7 TWh a month in April, May and August against 0.15 normally) in all three Elia datasets together, which neither Elia's documentation nor the flows explain. Not patched."),
    "Spain": (
        "ENTSO-E has no bidding zones for Morocco or Andorra, so Spain's ENTSO-E net imports (France and Portugal only) missed the exports Red Electrica reports to Morocco (1.86 / 2.54 / 3.75 TWh in 2023 / 2024 / 2025) and Andorra (0.24 / 0.24 / 0.21 TWh). "
        "REE's four border balances sum exactly to its published cross-border balance (13.96 / 10.23 / 12.80 TWh net export), and its France and Portugal figures match ENTSO-E's within 0.1 TWh. Adding Morocco and Andorra "
        "(SPAIN_REE_EXCHANGES.py, REData) took supply from 101.0 / 100.7 / 101.0% to 100.1 / 99.5 / 99.4% of load; the remaining -0.5% (1-1.6 TWh) is not traced."),
    "Denmark": "Load is Energinet's settlement gross consumption (incl. grid losses and 2.7 TWh of power-to-heat in 2025). ENTSO-E's Danish load is 4-7% lower, which made supply look 4-7% too high; "
               "ENTSO-E net imports match Energinet's exchanges (7.4 TWh in 2025). Remaining gap: ENTSO-E generation is about 1 TWh above Energinet's production.",
    "Poland": ("ENTSO-E's Polish load is PSE's national demand (KSE): PSE's open-data API (api.raporty.pse.pl, kse-load / his-wlk-cal 'demand') equals the ENTSO-E figure to within 0.2% every month from "
               "Jul 2024 to Dec 2025, and PSE's generation by fuel (his-gen-pal) equals ENTSO-E's (15.76 TWh in Jan 2025 both). The series steps down on 14 June 2024, the day PSE's new reporting platform starts: the daily gap "
               "load - (generation + net imports - pumping) falls from +22-28 GWh a day (about 1 GW, 8.8 TWh a year, 5.5% of load) to -1..-4 GWh, i.e. supply was 93.8% / 94.6% / 97.9% of load in 2022 / 2023 / 2024 and 100.1% in 2025. "
               "Generation is not the cause (ENTSO-E is within 0.8-2 TWh of Eurostat in 2022-23 while load was 11.6-12.7 TWh above Eurostat consumption until June 2024 and level with it afterwards), so PSE changed what its demand "
               "figure includes (about 1 GW of own use / station load or similar) on that date; the PSE API holds nothing before 14 June 2024, so 2022-23 cannot be restated from a raw PSE series and the balancing item is left as the measure of it."),    "Slovakia": "Net imports exclude double-counted Ukraine flows (ENTSO-E reports the same tie-lines under three Ukraine zones); supply now matches load within 1%.",
    "Finland": "2021-22 imports from Russia are not in the ENTSO-E flow data used here.",
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
                      extra_exports=None, prod_adjust=None, prod_override=None, import_floor=None, export_floor=None, until=None,
                      lng_extra=None, prod_extra=None):
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
    if cons_override is not None:
        # ENTSOG's exits (a partial series for most countries) fill only days BEFORE the national series ends; after its last day the days are left blank
        # (not backfilled with a partial total), so a month the national feed has not finished is dropped or scaled from its own days, never mixed.
        last_ok = cons_override.last_valid_index()
        day["Consumption"] = cons_override.reindex(bal.index).combine_first(own.where(own.index <= last_ok) if last_ok is not None else own)
    else:
        day["Consumption"] = own
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
    if prod_extra is not None:          # system entries ENTSOG lacks (Great Britain's Rough sub-terminal and small onshore entries): added to production
        day["Production"] = day["Production"] + prod_extra.reindex(bal.index).fillna(0)
    if lng_extra is not None:           # LNG that leaves the terminals without passing the send-out into the grid (Spain's LNG truck loadings): added to the LNG line
        day["LNG send-out"] = day["LNG send-out"] + lng_extra.reindex(bal.index).fillna(0)
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


def haidach_flows(data_dir):
    """Gas moved between the German grid and the Haidach storage (GWh/d), bayernets' physical flow at UGS-00274 'Haidach (AT) / Haidach USP (DE)' from ENTSOG
    (ENTSOG_POINT_FIXES_DAILY.py): {"in": withdrawals from Haidach into the German grid, "out": injections into Haidach} or None. Haidach lies in Austria and is in AGSI+'s
    Austrian stock, but it is fed from and delivers to the German grid, so neither Germany's AGSI+ storage nor its ENTSOG border rows hold it: Germany + Netherlands
    ran +5 TWh a month in summer (gas sent to Haidach) and -6 to -9 in winter (gas it delivered) until it was added to Germany's storage lines.
    RAG's Haiming 2 storage (bayernets ITP-00308) is the same case and is added to the same flows: Austrian AGSI+ stock change minus AGGM's market-area East storage
    equals Haidach + Haiming within about 0.1 TWh a month (AGGM publishes no Austrian-fed share: SSO Haidach THE/VTP and Ueberackern 7Fields allocations are empty/zero)."""
    fx = _sheet_or_empty(os.path.join(data_dir, POINT_FIXES_FILE), "Daily", "date")
    if not {"DE_haidach_in", "DE_haidach_out"} <= set(fx.columns):
        return None
    hin, hout = fx["DE_haidach_in"], fx["DE_haidach_out"]
    if {"DE_haiming_in", "DE_haiming_out"} <= set(fx.columns):       # RAG's Haiming 2 storage (ITP-00308): same set-up, also in AGSI+'s Austrian stock
        hin = hin.add(fx["DE_haiming_in"], fill_value=0)
        hout = hout.add(fx["DE_haiming_out"], fill_value=0)
    return {"in": hin, "out": hout}


def de_storage_with_haidach(data_dir, sto):
    """AGSI+ storage frame with Germany's withdrawal / injection columns plus the Haidach flows (see haidach_flows); `sto` unchanged where those are missing."""
    h = haidach_flows(data_dir)
    if h is None or "DE_withdrawal_GWhd" not in sto:
        return sto
    out = sto.copy()
    ix = out.index
    out["DE_withdrawal_GWhd"] = _col(sto, "DE_withdrawal_GWhd").add(h["in"].reindex(ix).fillna(0), fill_value=0)
    out["DE_injection_GWhd"] = _col(sto, "DE_injection_GWhd").add(h["out"].reindex(ix).fillna(0), fill_value=0)
    return out


def at_storage_without_haidach(data_dir, sto):
    """AGSI+ storage frame with the Haidach flows that now sit in Germany's storage (see haidach_flows) taken out of Austria's stock change, so the EU sums count
    them once. Austria's own AGGM balance never had them."""
    h = haidach_flows(data_dir)
    if h is None or "AT_withdrawal_GWhd" not in sto:
        return sto
    out = sto.copy()
    ix = out.index
    net = (_col(sto, "AT_withdrawal_GWhd") - _col(sto, "AT_injection_GWhd")) - (h["in"].reindex(ix).fillna(0) - h["out"].reindex(ix).fillna(0))
    net = net.where(_col(sto, "AT_withdrawal_GWhd").notna())
    out["AT_withdrawal_GWhd"] = net.clip(lower=0)
    out["AT_injection_GWhd"] = (-net).clip(lower=0)
    return out


def de_at_exports(data_dir):
    """German exit flows to Austria (GWh/d) that the ENTSOG border rule leaves out: Ueberackern ABG/SUDAL and RC Lindau (ENTSOG_POINT_FIXES_DAILY.py).
    The rule keeps the larger of the VIP sum and the physical-point sum on a side; VIP Oberkappel (69.5 TWh in 2025) plus VIP Kiefersfelden-Pfronten
    win, but Ueberackern (24.2) and Lindau (2.7) are not in either VIP, so Germany's exports to Austria came out 73 TWh against Austria's 94.
    If the rule is ever changed to add complementary points, drop this correction (it would count them twice)."""
    fx = _sheet_or_empty(os.path.join(data_dir, POINT_FIXES_FILE), "Daily", "date")
    cols = [c for c in ("DE_at_ueberackern", "DE_at_ueberackern2", "DE_at_lindau") if c in fx]
    return fx[cols].sum(axis=1, min_count=1) if len(cols) == 3 else None


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


# Countries whose storage flows stay on their own operator series: Denmark (Energinet) and Great Britain (National Gas site data, AGSI+ stopped for GB).
STOCK_STORAGE_SKIP = ("EU", "DK", "GB", "UA")


def stock_flows(sto, skip=STOCK_STORAGE_SKIP):
    """AGSI+ daily frame with every country's injection/withdrawal columns REPLACED by the day-to-day change in its AGSI+ stock level (<CC>_TWh),
    the measured quantity: AGSI's own flow columns disagree with the stock (Italy +28.6 TWh in 2025, Latvia 8.5 TWh in 2023, Austria 6-9 TWh a year because
    its stock includes Haidach, fed through Germany). Data quality: missing days are carried forward (up to 3), and a one-day spike that reverses the next
    day (a stock misreport) is replaced by the mean of its neighbours. Countries without a stock column keep AGSI's flows."""
    out = sto.copy()
    full = pd.date_range(sto.index.min(), sto.index.max(), freq="D")
    for col in [c for c in sto.columns if str(c).endswith("_TWh")]:
        cc = col[:-4]
        if cc in skip:
            continue
        x = sto[col].reindex(full).ffill(limit=3)
        d = x.diff()
        md = d.abs().median()
        spike = (d.abs() > max(8 * md, 0.02 * x.max())) & (d.shift(-1) * d < 0) & ((d + d.shift(-1)).abs() < 0.5 * d.abs())
        x = x.where(~spike, (x.shift(1) + x.shift(-1)) / 2)
        dst = (x.diff() * 1000.0).reindex(sto.index)
        out[f"{cc}_withdrawal_GWhd"] = (-dst).clip(lower=0)
        out[f"{cc}_injection_GWhd"] = dst.clip(lower=0)
    return out


def storage_stock_vs_flows(sto_raw, sto_stock, years=(2022, 2023, 2024, 2025, 2026)):
    """Net withdrawal by year (TWh): AGSI+ flow columns less the stock change, per country (positive = the flow columns show more withdrawal than the stock did)."""
    rows = {}
    for col in [c for c in sto_raw.columns if str(c).endswith("_withdrawal_GWhd")]:
        cc = col.split("_")[0]
        if cc in STOCK_STORAGE_SKIP or f"{cc}_injection_GWhd" not in sto_raw:
            continue
        fl = (sto_raw[col].fillna(0) - sto_raw[f"{cc}_injection_GWhd"].fillna(0)) / 1000.0
        stk = (sto_stock[col].fillna(0) - sto_stock[f"{cc}_injection_GWhd"].fillna(0)) / 1000.0
        rows[cc] = (fl - stk).groupby(fl.index.year).sum().reindex(list(years))
    return pd.DataFrame(rows).T.round(1)


def agsi_storage_frame(f, sto, cc):
    """Monthly balance frame with its storage lines replaced by the AGSI+ stock change of `cc` (monthly net). Austria: AGGM's market-area storage series
    leaves out Haidach (a 2.7 bcm site in Austria fed from the German grid, in AGSI+'s Austrian stock), so the gas Germany sends to Haidach reached no
    balance (6.5 TWh in 2025, 9.1 TWh in Jan-Sep 2026 in the EU total); the EU sum takes Austria's storage from the AGSI+ stock instead."""
    if f is None or not len(f) or f"{cc}_withdrawal_GWhd" not in sto:
        return f
    d = (_col(sto, f"{cc}_withdrawal_GWhd") - _col(sto, f"{cc}_injection_GWhd")).dropna()
    m = d.resample("MS").sum() / 1000.0
    n = d.resample("MS").count()
    ok = (n / pd.Series(n.index.days_in_month, index=n.index)) >= 0.9
    m = m.where(ok).reindex(f.index)
    out = f.copy()
    out["Storage withdrawals"] = m.clip(lower=0).where(m.notna(), f["Storage withdrawals"])
    out["Storage injections"] = m.clip(upper=0).where(m.notna(), f["Storage injections"])
    return out


PL_GAZSYSTEM_FILE = "poland_gazsystem_points_daily.xlsx"


def pl_yamal_imports(data_dir):
    """Poland's Belarus / Yamal gas that ENTSOG has no rows for (GWh per day), from Gaz-System's own entries (GAZSYSTEM_ENTRIES_DAILY.py):
    the Wysokoje and Tietierowka entries, plus the Yamal-Europe gas that reached Gaz-System at the PWP interconnection (and ONTRAS) beyond ENTSOG's
    Germany>Poland entry (ENTSOG's Mallnow figure is the net physical flow, so forward Yamal flow is missing from it), plus ENTSOG's Poland>Germany
    exit flow (Yamal transit through EuRoPol Gaz that never entered Gaz-System). The Yamal terms stop after May 2022 (Russia cut off Poland on 27 April;
    later days differ only by noise). None when the workbook is missing."""
    gz = _sheet_or_empty(os.path.join(data_dir, PL_GAZSYSTEM_FILE), "Daily", "date")
    if not len(gz) or "BY_wysokoje" not in gz:
        return None
    brd = _sheet_or_empty(os.path.join(data_dir, GAS_FLOWS_FILE), "Border flows", "date")
    if not len(brd) or "DE>PL" not in brd:
        return None
    ix = gz.index
    by = gz["BY_wysokoje"].fillna(0) + gz["BY_tietierowka"].fillna(0)
    entries = gz["DE_pwp"].fillna(0) + gz["DE_ontras"].fillna(0)
    ent = brd["DE>PL"].reindex(ix).fillna(0)
    out = brd["PL>DE"].reindex(ix).fillna(0) if "PL>DE" in brd else 0.0
    yam = ((entries - ent).clip(lower=0) + out).where(ix <= "2022-05-31", 0.0)
    return by + yam


def es_lng_trucks(data_dir):
    """LNG loaded onto trucks at Spain's regasification plants, GWh/d: the monthly total from the Enagas statistical bulletin's 'Regasification plants activity' table
    (GAS_TSO_SOUTHEAST_DAILY.py, sheet Monthly of the extra TSO workbook), spread evenly over the days of the month. Months without a figure stay blank (not filled)."""
    mon = _sheet_or_empty(os.path.join(data_dir, TSO_EXTRA_FILE), "Monthly", "month")
    if not len(mon) or "ES_lng_trucks" not in mon or mon["ES_lng_trucks"].notna().sum() < 12:
        return None
    m = mon["ES_lng_trucks"].dropna()
    parts = [pd.Series(v / d.days_in_month, index=pd.date_range(d, d + pd.offsets.MonthEnd(0), freq="D")) for d, v in m.items()]
    return pd.concat(parts).sort_index()


def fr_bio_transmission(data_dir):
    """France: the part of the ODRE biomethane series injected straight into the transmission networks (GWh/d), None until BIOMETHANE_DAILY.py has pulled it."""
    d = _sheet_or_empty(os.path.join(data_dir, BIO_FILE), "Daily", "date")
    if "FR_biomethane_transmission" not in d or d["FR_biomethane_transmission"].notna().sum() < 365:
        return None
    return d["FR_biomethane_transmission"]


def point_fix_args(data_dir, cc, tso, bio, gni, bal_nl=None, sto=None):
    """Country-specific corrections from ENTSOG points the main pull's classification drops (ENTSOG_POINT_FIXES_DAILY.py), as
    (keyword arguments for gas_country_balance, consumption override or None, note text or None).
    GR: TAP's Nea Mesimvria entry (Azerbaijani gas) is added to pipeline imports. HU: the 'Exit for Blending' is taken off production
    (imported gas blended with high-CO2 domestic gas re-enters at the production entry). UK: the Moffat exit is added to exports - the
    Republic of Ireland's share is GNI's own Moffat import figure, the rest (Northern Ireland, Isle of Man) is UK consumption that the
    National Gas NTS offtake series lacks. FR: ODRE consumption is the GRTgaz/Teréga offtake (it equals ENTSOG's distribution plus
    industrial exits), which excludes biomethane injected straight into the distribution networks, so that biomethane is added to
    consumption (it is also a supply line)."""
    if cc == "IT" and it_storage_from_stock(data_dir) is not None:
        kw, cons, note = {"storage": it_storage_from_stock(data_dir)}, None, (
            " Storage flows are the day-to-day change in AGSI+'s stock level: AGSI's own injection and withdrawal columns for Italy disagree with the stock by 28.6 TWh in 2025.")
        fx_it = _sheet_or_empty(os.path.join(data_dir, POINT_FIXES_FILE), "Daily", "date")
        if bal_nl is not None and "IT_srg_other_tsos" in fx_it and "IT_distribution_GWhd" in bal_nl:
            own = bal_nl[["IT_distribution_GWhd", "IT_final_consumers_GWhd"]].sum(axis=1, min_count=1)
            cons = own + fx_it["IT_srg_other_tsos"].reindex(own.index).fillna(0)
            note += (" Consumption adds Snam Rete Gas's 'delivery to other transmission networks' (ENTSOG point ITP-00288, 14.6 TWh in 2025): gas handed to the smaller Italian "
                     "transmission systems, whose downstream exits ENTSOG does not publish, so it is missing from the distribution + industrial + thermal exits "
                     "(it was the whole +1.2 TWh a month surplus). What remains (about 0.3 TWh a month) is Snam's compressor fuel and losses, which no published series gives.")
        return kw, cons, note
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
    if cc == "ES":
        tr = es_lng_trucks(data_dir)
        if tr is not None:
            return {"lng_extra": tr}, None, (
                " LNG send-out includes the LNG loaded onto trucks at the seven regasification plants (about 11-12 TWh a year, Enagás statistical bulletin, spread evenly over the days of each month): "
                "Enagás's national market demand (conventional market) includes it, but it leaves the terminals without passing the send-out into the grid (ALSI send-out counts grid send-out only), "
                "and it was the whole -0.9 TWh a month shortfall. LNG ship reloads (international market demand) are not in national demand and are not added.")
    if cc == "FR" and bal_nl is not None and "FR_distribution_GWhd" in bal_nl and len(bio) and "FR" in bio:
        own = bal_nl[["FR_distribution_GWhd", "FR_final_consumers_GWhd"]].sum(axis=1, min_count=1)
        t = fr_bio_transmission(data_dir)
        if t is not None:
            dn = (bio["FR"] - t.reindex(bio.index).fillna(0)).clip(lower=0)      # biomethane injected into the distribution networks
            return {"prod_override": pd.Series(0.0, index=bal_nl.index)}, own + dn.reindex(own.index).fillna(0), (
                " Consumption is the gas that leaves the GRTgaz/Teréga transmission networks (ENTSOG distribution + final-consumer exits, metered) plus the biomethane injected into the distribution "
                "networks. ODRE's daily consumption series was used before: its distribution part equals ENTSOG's exits (-0.1 TWh a month) but its industrial part is 9.7 TWh a year (8%) below the metered "
                "industrial exit in 2025 (1.5 TWh a month in winter, 0.3 in summer), i.e. gas the networks deliver for uses ODRE's industrial list does not cover (network, storage and terminal own use among them); "
                "Eurostat's own French balance carries a 7.9 TWh statistical difference of the same size. ENTSOG's French production entry (GRTgaz and Teréga biomethane producers, from Apr 2025, 2.2 TWh) is "
                "biomethane already counted in the biomethane line, so it is not counted twice; the biomethane injected into the transmission networks (2.9 TWh in 2025) is inside the exits and is not added to consumption.")
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
    if cc == "PL":
        pl_in = pl_yamal_imports(data_dir)
        return ({"extra_imports": pl_in} if pl_in is not None else {}), None, (
            (" Pipeline imports add the Belarus / Yamal gas of Oct 2021 - May 2022 that ENTSOG has no Kondratki / Wysokoje rows for (Gaz-System's own Market Information Module: the Wysokoje and Tietierowka entries, "
             "12.8 + 0.4 TWh in 2022, plus the Yamal-Europe gas that entered Gaz-System at the PWP interconnection beyond ENTSOG's Mallnow entry, or left Poland at Mallnow towards Germany, 14.2 TWh in 2022 and about 27 in Q4 2021); "
             "together 27.5 TWh in 2022, exactly Eurostat's Belarus partner figure, which took the 2022 balance from -15% to +0.5%." if pl_in is not None else
             " Poland 2022 runs about 15% short because ENTSOG has no Kondratki / Wysokoje (Yamal, Belarus) entry rows: Eurostat's Belarus partner figure for 2022 is 27.5 TWh, the size of the gap.") +
            " Border audit (ENTSOG both sides, Eurostat as a benchmark only): every other Polish border agrees within 10% except Ukraine in 2025 (exit 22.9, Ukrainian entry 20.9 TWh; the larger is used). "
            "Baltic Pipe (Faxe) matches Energinet within 2%.")
    if cc == "BG":
        return {}, None, (
            " Border audit: ENTSOG's Bulgarian entries and exits at Negru Voda, Kireevo, Kulata, Kyustendil and Komotini agree with the neighbours' own sides within 1% "
            "(TurkStream enters at Strandzha 2 and Strandzha, 206 TWh in 2025; Serbia's onward flow to Hungary plus its consumption matches Kireevo), and no border is double counted. "
            "The remaining surplus of about 2-3 TWh a year on 190-215 TWh of transit (about 1%) is compressor fuel gas and transit losses that no published series gives.")
    if not len(fx):
        return {}, None, None
    if cc == "DE" and nord_stream(data_dir) is not None:
        em_raw = emden_entsog(data_dir)
        ns = nord_stream(data_dir)
        kw = {"extra_imports": ns.add(em_raw["DE"].reindex(ns.index).fillna(0), fill_value=0).combine_first(em_raw["DE"]) if em_raw else ns}
        if sto is not None and haidach_flows(data_dir) is not None:
            kw["storage"] = de_storage_with_haidach(data_dir, sto)
        de_at = de_at_exports(data_dir)
        if de_at is not None:
            kw["extra_exports"] = de_at
        return kw, None, (
            (" Pipeline exports add the German exit flows to Austria at Ueberackern and Lindau (24-27 TWh a year), which the border rule drops because the VIP Oberkappel it prefers does not contain them (Germany's exports to Austria were 73 TWh in 2025 against Austria's own 94)." if de_at is not None else "") +
            (" Storage withdrawals and injections are AGSI+'s stock change plus bayernets' physical flow at the Haidach storage (ENTSOG UGS-00274; Haidach lies in Austria and is in AGSI+'s Austrian stock but is fed from and delivers to the German grid, 10-24 TWh a year each way)." if (sto is not None and haidach_flows(data_dir) is not None) else "") +
            (" Pipeline imports include Norwegian gas at Emden (EPT1: OGE and GUD entries; the Dutch share, GTS, is in the Netherlands balance), which ENTSOG's country classification drops." if em_raw else "") +
            " Pipeline imports include the Nord Stream 1 gas entering at Greifswald (NEL and OPAL entries, Russian origin, to Sept 2022), which "
            "ENTSOG's country classification drops because it lists no far side for those points (Germany + Netherlands was about 100 TWh a quarter short before).")
    # Greece: no TAP point fix any more - the rebuilt border-flow pull books the Nea Mesimvria entry as AL>GR (11.1 TWh in 2025), so adding
    # GR_tap_imports on top counted it twice (balance +16%); the pull's own AL>GR is in GR_imports.
    if cc == "HU" and "HU_production_exit" in fx:
        return {"prod_adjust": fx["HU_production_exit"]}, None, (
            " Production is ENTSOG's 'Aggregated Single Production' entry less the 'Exit for Blending': imported gas leaves the grid, is blended "
            "with high-CO2 domestic gas and re-enters at the production entry, so that entry double-counts about 14 TWh a year of imports. "
            "ENTSOG's production entry only starts in Feb 2023 (about 2.2 TWh a month), so the Oct 2021 - Jan 2023 balance is about 13% short of consumption; the borders (Austria, Serbia, Romania, Croatia, Slovakia, Ukraine) agree with both neighbours' sides.")
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
        uk_entry, uk_rough, ent = None, None, None
        try:    # National Gas system entries ENTSOG lacks / double counts (GB_GAS_NTS_DAILY.py entry_* columns)
            nts_e = add_charts._sheet(os.path.join(data_dir, "gb_gas_nts_daily.xlsx"), "Daily", "date")
            ec = [c for c in ("entry_rough_subterminal", "entry_saltfleetby", "entry_murrow", "entry_glentham", "entry_burton_point") if c in nts_e]
            if "entry_rough_subterminal" in nts_e and nts_e["entry_rough_subterminal"].notna().sum() > 30:
                uk_entry = nts_e[ec].sum(axis=1, min_count=1)
                uk_rough = nts_e["entry_rough_storage"] if "entry_rough_storage" in nts_e else None
        except Exception:  # noqa: BLE001
            pass
        ukkw = ({"prod_extra": uk_entry} if uk_entry is not None else {}) | ({"prod_adjust": uk_rough} if uk_rough is not None else {})
        return ({"extra_exports": to_roi} | ukkw | ({"storage": sto} if sto is not None else {})), tso["UK"].add((mof - to_roi - ukie).clip(lower=0).fillna(0)), (
            " Exports include the Moffat exit to Ireland (the Republic's share is Gas Networks Ireland's Moffat import figure); the remainder of "
            "the Moffat flow (Northern Ireland and the Isle of Man, about 19 TWh a year, less the 7-9 TWh a year that Northern Ireland sends on to the Republic at Carrickfergus, which ENTSOG already counts in UK exports) is added to UK consumption because the National Gas NTS "
            "offtake covers Great Britain only. Storage withdrawals and injections are the nine storage sites' own daily flows from the National Gas Data Portal (the NTS aggregate overstates net withdrawals by about 5 TWh a year); before Oct 2024 they are the day-to-day change in the portal's total stock (net only). ENTSOG omits Moffat from its UK exports (the point's far side is listed as country UK)."
            + (" Production adds the National Gas system entries ENTSOG has no row for - the Rough sub-terminal at Easington (1.0-2.0 TWh a month, 17 TWh in 2025) and the small Saltfleetby, Murrow and Glentham biomethane entries - "
               "and takes off the Rough storage withdrawals that ENTSOG's Easington entry already contains until Sep 2025 (ENTSOG Easington less Langeled and Dimlington equals Rough storage to 0.01 TWh), which the storage line also counts. "
               "ENTSOG's other UK entries (St Fergus, Easington, Teesside, Bacton, Barrow) match National Gas's own entry allocations within 0.1%. NTS shrinkage (compressor fuel and unaccounted gas, 2-4 TWh a year) is not published as consumption and stays in the remainder."
               if uk_entry is not None else ""))
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


GWH_PER_MCM = 11.2           # GWh per million cubic metres (the Gassco conversion used for Norwegian flows; European gas varies roughly 10.5-11.5)
MCM_PER_BCF = 28.3168        # million cubic metres per billion cubic feet
GWH_PER_BCF = GWH_PER_MCM * MCM_PER_BCF


def gas_summary_table(eu):
    """EU27 gas balance in Bcf/d: latest complete month, the month before, the same month a year earlier, with month-on-month and year-on-year changes (absolute, no percentages)."""
    d = eu.copy()
    g = lambda c: d[c] if c in d else pd.Series(0.0, index=d.index)
    prod = g("Production") + g("Biomethane")
    rows = pd.DataFrame({
        "Demand (consumption)": g("Consumption"),
        "Production (incl. biomethane)": prod,
        "Net pipeline imports": g("Pipeline imports") + g("Pipeline exports"),   # exports are stored as negatives
        "LNG send-out": g("LNG send-out"),
        "Storage net withdrawal": g("Storage withdrawals") + g("Storage injections"),
    })
    rows["Supply less demand (residual)"] = rows.drop(columns="Demand (consumption)").sum(axis=1) - rows["Demand (consumption)"]
    days = pd.Series(d.index.days_in_month, index=d.index)
    bcfd = rows.mul(1000.0, axis=0).div(days, axis=0) / GWH_PER_BCF       # TWh per month -> GWh -> per day -> Bcf/d
    last = bcfd.index.max()
    prev, yago = last - pd.DateOffset(months=1), last - pd.DateOffset(years=1)
    out = pd.DataFrame(index=bcfd.columns)
    if prev in bcfd.index:
        out[f"{prev:%b %Y}"] = bcfd.loc[prev]
    out[f"{last:%b %Y}"] = bcfd.loc[last]
    if prev in bcfd.index:
        out["MoM"] = bcfd.loc[last] - bcfd.loc[prev]
    if yago in bcfd.index:
        out["YoY"] = bcfd.loc[last] - bcfd.loc[yago]       # change against the same month a year earlier (that month's level is not shown)
    return out, bcfd


def add_gas_summary_sheet(wb, used, eu, label="EU27"):
    out, bcfd = gas_summary_table(eu)
    ws = wb.create_sheet(sam.sheet_name("Summary - Gas Bcf per day", used))
    ws.append([f"{label} gas balance, billion cubic feet per day (Bcf/d)"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append([f"Latest complete month {bcfd.index.max():%b %Y} and the month before; MoM is the change on the month before, YoY the change on the same month a year earlier (Bcf/d). {label} = the sum of the corrected country balances (see the EU27 + UK gas balance tab)."])
    ws.append([])
    ws.append(["Bcf/d"] + list(out.columns))
    for c in ws[4]:
        c.font = Font(bold=True)
    for name, r in out.iterrows():
        ws.append([name] + [None if pd.isna(v) else round(float(v), 2) for c, v in r.items()])
    ws.append([])
    ws.append([f"Conversion: {GWH_PER_MCM} GWh per million m3 (Gassco's factor) and {MCM_PER_BCF} million m3 per Bcf, i.e. {GWH_PER_BCF:.1f} GWh per Bcf; "
               "energy balances are converted at one fixed factor, so Bcf/d figures move about 5% if a gross calorific value of 10.6 rather than 11.2 kWh/m3 is used."])
    ws.append(["Demand is consumption (national TSO and statistics series, ENTSOG exits where none); production is ENTSOG production plus biomethane; net pipeline imports are from/to outside the EU; "
               "storage is net withdrawal (positive) or injection (negative), the change in the GIE AGSI+ stock level; LNG is GIE ALSI send-out. The residual is the unexplained difference "
               "(network own use, losses and linepack that consumption series leave out, measurement differences); it is shown, not plugged."])
    ws.append(["Sources: ENTSOG, GIE AGSI+ / ALSI, national TSOs (see the Sources tab)."])
    ws.column_dimensions["A"].width = 34
    for col in "BCDEFGHI":
        ws.column_dimensions[col].width = 13
    return ws



def add_storage_check_sheet(wb, used, raw, stock):
    """Net withdrawal by country and year (TWh): AGSI+'s injection/withdrawal columns less the stock change (what the balances now use)."""
    t = storage_stock_vs_flows(raw, stock)
    ws = wb.create_sheet(sam.sheet_name("Storage stock vs flows", used))
    ws.append(["AGSI+ flow columns less AGSI+ stock change, net withdrawal in TWh per year (positive: the flow columns show more withdrawal than the stock fell)"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append(["The gas balances use the stock change (the measured quantity); Denmark (Energinet) and Great Britain (National Gas sites) keep their own series. 2026 is January to the latest day. Austria: AGGM's market-area storage series (used in the Austria tab) is 1.5 TWh below the AGSI+ stock change in Aug-Dec 2024, 6.5 above in 2025 and 9.1 above in Jan-Sep 2026 (Haidach, fed from the German grid, is in the AGSI+ stock only); the EU totals take Austria's storage from the AGSI+ stock."])
    ws.append([])
    ws.append(["Country"] + [str(c) for c in t.columns])
    for c in ws[4]:
        c.font = Font(bold=True)
    for cc, r in t.iterrows():
        ws.append([GAS_NAMES.get(cc, cc)] + [None if pd.isna(v) else float(v) for v in r])
    ws.column_dimensions["A"].width = 18
    return ws


EXTRA_EXPORT_DEST = {"DE": "AT", "BE": "LU", "UK": "IE", "LV": "EE"}   # point-fix exports (extra_exports) and the member that receives them
WHOLE_FRAME_MEMBERS = ("AT", "DK", "EE", "CZ", "IE", "LU")   # frames built from the operator's own border series (AGGM, Energinet, Elering, NET4GAS floors, GNI, Creos): every pipeline line is a flow with another member


def intra_block_flows(lines, block, sides, extra_out=None, whole=WHOLE_FRAME_MEMBERS, months=None):
    """Flows between two members of `block`, as the member frames in `lines` (monthly TWh, positive imports, negative exports) book them: (imports, exports),
    both positive monthly TWh. The receiving member's entry and the sending member's exit are each member's OWN side from the 'Border flows by side'
    sheet (ENTSOG pull), so taking the first out of the block's imports and the second out of its exports leaves one consistent figure per
    border: whatever the two operators measure differently (NL exit 258.7 vs DE entry 249.8 TWh in 2025) no longer sits in the residual. Frames built from an
    operator's own series (`whole`) are intra in full; exports a point fix adds (`extra_out`: Germany -> Austria, Belgium -> Luxembourg, Great Britain -> Ireland,
    Latvia -> Estonia) count when the receiving member is in the block."""
    idx = months if months is not None else sorted(set().union(*[f.index for f in lines.values()]))
    imp, exp = pd.Series(0.0, index=idx), pd.Series(0.0, index=idx)

    def side_sum(cols):
        cols = [c for c in cols if c in sides]
        if not cols:
            return pd.Series(0.0, index=idx)
        m = monthly_cover(sides[cols].sum(axis=1, min_count=1).to_frame("x")).dropna(how="all") / 1000.0
        return m["x"].reindex(idx).fillna(0.0)

    for c, f in lines.items():
        if c in whole:
            imp += f["Pipeline imports"].reindex(idx).fillna(0.0)
            exp += (-f["Pipeline exports"]).reindex(idx).fillna(0.0)
            continue
        imp += side_sum([f"{a}>{c} entry" for a in block if a != c])
        exp += side_sum([f"{c}>{b} exit" for b in block if b != c])
        dest = EXTRA_EXPORT_DEST.get(c)
        if extra_out and c in extra_out and dest in block:
            xo = extra_out[c]
            exp += (monthly_cover(xo.to_frame("x")).dropna(how="all")["x"] / 1000.0).reindex(idx).fillna(0.0)
    return imp, exp


def eu_gas_balance(frames, border=None, fallback=None, min_share=0.8, block=EU27_GAS, fill_max_share=0.03, sides=None, extra_out=None):
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
    lines = {c: f for c, f in lines.items() if c in block}
    span = pd.date_range("2021-10-01", max(f.index.max() for f in lines.values()), freq="MS")
    keep = {c: f for c, f in lines.items() if f.reindex(span).notna().any(axis=1).mean() >= min_share}
    left = [c for c in lines if c not in keep]
    for c, f in keep.items():   # a one- or two-month hole in a small country's feed (Greece Apr 2022, Luxembourg Sep-Oct 2023) is interpolated rather than dropping the EU month
        g = f.reindex(span)
        g = g.interpolate(limit=2, limit_area="inside")
        # a small country's feed that stops (Estonia, Latvia after Oct 2025; Finland, Lithuania in the latest month): same month a year earlier, counted as a gap-fill.
        # Only for countries under `fill_max_share` of the block's consumption: a large country's missing month (France's ODRE, Spain's Enagas) drops that month for the block.
        if f["Consumption"].mean() <= fill_max_share * sum(h["Consumption"].mean() for h in lines.values()):
            g = g.where(g.notna().any(axis=1), g.shift(12))
        keep[c] = g.dropna(how="all")
    months = span[[all(m in f.index for f in keep.values()) for m in span]]
    tot = sum(f.reindex(months).fillna(0.0) for f in keep.values())
    if sides is not None and len(sides):    # one consistent figure per intra-block border: each member's own side comes out of its own line (see intra_block_flows)
        imp, exp = intra_block_flows({c: f.reindex(months) for c, f in keep.items()}, set(keep), sides, extra_out, months=months)
        tot["Pipeline imports"] = tot["Pipeline imports"] - imp
        tot["Pipeline exports"] = tot["Pipeline exports"] + exp
    elif border is not None and len(border):
        pairs = [c for c in border.columns if ">" in c and all(x in block for x in c.split(">"))]
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
        gsto_raw = add_charts._sheet(os.path.join(args.data_dir, "eu_gas_storage_daily.xlsx"), "Daily", "date")
        gsto = stock_flows(gsto_raw)      # storage flows = day-to-day change of the AGSI+ stock level, not AGSI's flow columns
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
        eu_frames, eu_fallback, eu_extra_out = {}, {}, {}
        try:
            gsides = add_charts._sheet(os.path.join(args.data_dir, GAS_FLOWS_FILE), "Border flows by side", "date")
        except Exception:  # noqa: BLE001
            gsides = None
        for cc in [c for c in GAS_NAMES if any(col.startswith(f"{c}_") for col in gbal.columns)]:
            fix_kw, fix_cons, fix_note = point_fix_args(args.data_dir, cc, tso, bio, gni, gbal, gsto)
            cons_in = fix_cons if fix_cons is not None else (tso[cc] if (tso is not None and cc in tso) else None)
            sto_in = fix_kw.pop("storage", gsto)
            if cc in EXTRA_EXPORT_DEST and fix_kw.get("extra_exports") is not None:
                eu_extra_out[cc] = fix_kw["extra_exports"]
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
            if cc in EU27_GAS or cc == "UK":
                eu_frames[cc] = agsi_storage_frame(b, at_storage_without_haidach(args.data_dir, gsto), "AT") if (cc == "AT" and len(b)) else b
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
        eu, eu_left = eu_gas_balance(eu_frames, gbord, eu_fallback, sides=gsides, extra_out=eu_extra_out)
        eu_uk, _ = eu_gas_balance(eu_frames, gbord, eu_fallback, block=EU27_GAS + ["UK"], sides=gsides, extra_out=eu_extra_out)
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
            if not eu_uk.empty:
                total_chart(wb, used, gas, 1, eu_uk, [
                    "EU27 + United Kingdom: the EU27 sum of the country balances (see the EU gas balance tab) plus the Great Britain balance (National Gas NTS consumption and site-level storage, "
                    "Moffat / Carrickfergus point fixes, ALSI LNG, Norway split from Gassco). Flows between any two members, including Great Britain-Belgium (IUK), Great Britain-Netherlands (BBL) and "
                    "Great Britain-Ireland, are taken out of both pipeline lines (net unchanged); Norway, Algeria, Azerbaijan, Russia and the rest stay external. Storage in every member is the AGSI+ stock change "
                    "(Austria incl. Haidach; Great Britain the nine sites' own series; Denmark Energinet). A month is shown only when every large member's national consumption series covers it "
                    "(France ODRE, Spain Enagas, Germany THE ...); only members under 3% of the block's consumption (Estonia, Latvia, Finland, Lithuania ...) repeat the same month a year earlier when their feed stops."],
                            "EU27 + UK gas balance data", "EU27 + UK gas balance: supply and storage vs consumption (TWh per month)",
                            "TWh per month", GAS_BALANCE_SRC, "Notes:", label="EU27 + UK", line_cols=("Consumption",))
            try:
                add_gas_summary_sheet(wb, used, eu_uk, "EU27 + UK")
                add_storage_check_sheet(wb, used, gsto_raw, gsto)
            except Exception as e:  # noqa: BLE001
                gas[2].append(f"Gas summary table (Bcf/d) failed ({type(e).__name__}: {e})")
    if len(gbal) and emden is not None:
        try:   # Germany + Netherlands: Gassco's Emden gas is split between them in a way the raw data cannot show, so they are combined
            ns, em_raw = nord_stream(args.data_dir), emden_entsog(args.data_dir)
            em_de = em_raw["DE"] if em_raw else emden           # ENTSOG's own Emden entries where pulled, else the Gassco-based estimate
            de_b = gas_country_balance(gbal, "DE", de_storage_with_haidach(args.data_dir, gsto), glng, tso["DE"] if "DE" in tso else None, None,
                                       bio["DE"] if (len(bio) and "DE" in bio) else None,
                                       em_de.add(ns.reindex(em_de.index).fillna(0), fill_value=0).combine_first(ns) if ns is not None else em_de,
                                       extra_exports=de_at_exports(args.data_dir))
            nl_kw, nl_cons, _nl_note = point_fix_args(args.data_dir, "NL", tso, bio, gni, gbal)
            nl_b = gas_country_balance(gbal, "NL", gsto, glng, nl_cons, None, bio["NL"] if (len(bio) and "NL" in bio) else None, **nl_kw)
            colsb = [c for c in de_b.columns if c in nl_b.columns]
            both = de_b[colsb].add(nl_b[colsb], fill_value=0)
            if gsides is not None and len(gsides):    # NL>DE and DE>NL are inside the block: each side's own figure comes out of its own line, so the border cancels exactly
                d_imp, d_exp = intra_block_flows({"DE": de_b, "NL": nl_b}, {"DE", "NL"}, gsides, whole=(), months=both.index)
                both["Pipeline imports"] = both["Pipeline imports"] - d_imp
                both["Pipeline exports"] = both["Pipeline exports"] + d_exp
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
