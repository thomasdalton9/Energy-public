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
GAS_BALANCE_SRC = ("ENTSOG (production, pipeline flows, consumption), TSO consumption series (Germany THE, France ODRE, Spain Enagas, Denmark Energinet, Portugal REN, Austria AGGM, Czechia NET4GAS, Lithuania Amber Grid, Finland Gasgrid, Great Britain National Gas, Ireland GNI), Gassco (Norway to Great Britain), biomethane injection (France ODRE, Denmark Energinet, Netherlands CBS, Austria AGGM, other countries Eurostat annual), GIE ALSI (LNG send-out), GIE AGSI+ (storage)",
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


BALANCE_COLS = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Other", "Net imports",
                "Pumped & battery (net)", "Load"]


# Why reported supply can fall short of load for a country (ENTSO-E reports what TSOs meter; the rest is not in the feed).
KNOWN_GAPS = {
    "Switzerland": "Generation is Swissgrid's own production by carrier (storage hydro is gross of pumped-storage output); pumping consumption, "
                   "imports and exports are BFE's monthly electricity balance spread over the days. ENTSO-E's Swiss hydro was incomplete (supply/load ~70%) and is no longer used.",
    "Netherlands": "Generation is CBS monthly production by source (including rooftop solar) spread evenly over the days. Load is CBS consumption incl. distribution losses (ENTSO-E load is 10% lower in 2021-22, equal within 1% in 2024-25), so supply vs load closes by construction; flows are ENTSO-E and match CBS.",
    "Germany": "Industrial self-generation and small embedded plants are not in the feed; supply is typically 4-5% below load.",
    "Italy": "Embedded/self-consumed generation is not in the feed; supply is typically 2-5% below load.",
    "Great Britain": "Supply runs 2-3% above load: Elexon FUELHH metered output is gross of power-station own use, while NESO national demand (the load used) is net of it; "
                     "no raw station-load series is published, so the surplus is left visible.",
    "Bulgaria": "Supply is 3-7% below load in 2024-25: ENTSO-E generation misses part of the rapidly growing small solar fleet (not TSO-metered) and the load figure includes it; "
                "the ESO (TSO) site does not answer from GitHub, so there is no raw replacement yet.",
    "Montenegro": "The ENTSO-E physical flow on the Bosnia-Montenegro border (about 3 TWh a year) does not close either side's balance: Montenegro is oversupplied "
                  "(2025 +36%) and Bosnia undersupplied (-13%) by roughly the same volume. Montenegro's TSO (CGES) has no machine-readable feed reachable from GitHub (discovery_archive/europe/CGES_PROBE.py), "
                  "and 2025 generation is low because the Pljevlja coal plant was out Apr-Nov. Treat the Balkan (BA, ME, MK, XK, RS) balances as indicative.",
    "Bosnia and Herzegovina": "See Montenegro: the Bosnia-Montenegro physical flow looks overstated; supply is 13% below load in 2025 (ratio was 97-102% before).",
    "North Macedonia": "Small system with unreliable ENTSO-E load/flow reporting (supply 89-91% of load in 2023-24, 100% in 2025).",
    "Kosovo": "KOSTT generation is metered at the plant and load includes distribution losses and theft; supply is about 5% below load.",
    "Slovenia": "Supply is about 4% above load every year: ENTSO-E Slovenian load excludes some demand that generation and flows cover (grid losses/closed distribution systems).",
    "Serbia": "Supply is about 4% above load every year, consistent with a load definition that is net of transmission losses.",
    "Lithuania": "Supply is 3-4% above load since 2023, after the Baltic synchronisation changed the metered border flows.",
    "Denmark": "Load is Energinet's settlement gross consumption (incl. grid losses and 2.7 TWh of power-to-heat in 2025). ENTSO-E's Danish load is 4-7% lower, which made supply look 4-7% too high; "
               "ENTSO-E net imports match Energinet's exchanges (7.4 TWh in 2025). Remaining gap: ENTSO-E generation is about 1 TWh above Energinet's production.",
    "Poland": "Before 2024 supply is 5-6% below load: small embedded and industrial generation is not in the ENTSO-E feed.",
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
    out = [f"Supply (generation + net imports + pumped/battery net) is {ratio:.1%} of load over the last 12 months shown."]
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
                      extra_exports=None, prod_adjust=None):
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
    cols = GAS_BAL_COLS
    if biomethane is not None and biomethane.notna().any():
        day["Biomethane"] = biomethane.reindex(bal.index).fillna(0)
        cols = GAS_BAL_COLS[:-1] + ["Biomethane", "Consumption"]
    day = day.dropna(subset=["Pipeline imports", "Consumption"])
    day = day[day.index >= "2021-10-01"]
    return _monthly_twh(day[cols]) if len(day) else pd.DataFrame()


POINT_FIXES_FILE = "entsog_point_fixes_daily.xlsx"


def point_fix_args(data_dir, cc, tso, bio, gni):
    """Country-specific corrections from ENTSOG points the main pull's classification drops (ENTSOG_POINT_FIXES_DAILY.py), as
    (keyword arguments for gas_country_balance, consumption override or None, note text or None).
    GR: TAP's Nea Mesimvria entry (Azerbaijani gas) is added to pipeline imports. HU: the 'Exit for Blending' is taken off production
    (imported gas blended with high-CO2 domestic gas re-enters at the production entry). UK: the Moffat exit is added to exports - the
    Republic of Ireland's share is GNI's own Moffat import figure, the rest (Northern Ireland, Isle of Man) is UK consumption that the
    National Gas NTS offtake series lacks. FR: ODRE consumption is the GRTgaz/Teréga offtake (it equals ENTSOG's distribution plus
    industrial exits), which excludes biomethane injected straight into the distribution networks, so that biomethane is added to
    consumption (it is also a supply line)."""
    fx = _sheet_or_empty(os.path.join(data_dir, POINT_FIXES_FILE), "Daily", "date")
    if cc == "FR" and tso is not None and "FR" in tso and len(bio) and "FR" in bio:
        return {}, tso["FR"].add(bio["FR"].reindex(tso.index).fillna(0)), (
            " Consumption is ODRE's GRTgaz/Teréga offtake plus the biomethane injected into the distribution networks (ODRE's offtake equals "
            "ENTSOG's distribution + industrial exits and so excludes it). The balance still runs about 3% long: ENTSOG misses about 15 TWh of "
            "French exports against Eurostat, and network own use and losses are not in the offtake.")
    if not len(fx):
        return {}, None, None
    if cc == "GR" and "GR_tap_imports" in fx:
        return {"extra_imports": fx["GR_tap_imports"]}, None, (
            " Pipeline imports include the TAP entry at Nea Mesimvria (Azerbaijani gas), which ENTSOG's country classification drops "
            "because TAP's operator is listed with country GR.")
    if cc == "HU" and "HU_production_exit" in fx:
        return {"prod_adjust": fx["HU_production_exit"]}, None, (
            " Production is ENTSOG's 'Aggregated Single Production' entry less the 'Exit for Blending': imported gas leaves the grid, is blended "
            "with high-CO2 domestic gas and re-enters at the production entry, so that entry double-counts about 14 TWh a year of imports.")
    if cc == "UK" and "UK_moffat_exit" in fx and tso is not None and "UK" in tso:
        mof = fx["UK_moffat_exit"]
        roi = gni["Moffat"].reindex(mof.index) if gni is not None and len(gni) else pd.Series(float("nan"), index=mof.index)
        to_roi = roi.where(roi.notna(), mof).clip(upper=mof)
        try:    # National Gas NTS storage flows (the operator's own; ENTSOG lacks the Stublach, Holford and Hill Top entries)
            nts = add_charts._sheet(os.path.join(data_dir, "gb_gas_nts_daily.xlsx"), "Daily", "date")
            sto = pd.DataFrame({"GB_withdrawal_GWhd": nts["storage_withdrawal"], "GB_injection_GWhd": nts["storage_injection"]})
        except Exception:  # noqa: BLE001
            sto = None
        return ({"extra_exports": to_roi} | ({"storage": sto} if sto is not None else {})), tso["UK"].add((mof - to_roi).fillna(0)), (
            " Exports include the Moffat exit to Ireland (the Republic's share is Gas Networks Ireland's Moffat import figure); the remainder of "
            "the Moffat flow (Northern Ireland, Isle of Man, about 19 TWh a year) is added to UK consumption because the National Gas NTS "
            "offtake covers Great Britain only. Storage withdrawals and injections are National Gas NTS's own figures. ENTSOG omits Moffat from its UK exports (the point's far side is listed as country UK).")
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


NORWAY_ENTRIES_FILE = "entsog_norway_entries_daily.xlsx"


def emden_gap(data_dir, nor):
    """Norwegian gas that reaches Germany through Emden, GWh/d: Gassco's flow to Germany minus the part ENTSOG reports at Dornum.
    ENTSOG publishes nothing at Emden (EPT1, EPT2, NPT) under any indicator, so Germany's pipeline imports in the ENTSOG pull miss it.
    Gassco days that are missing take the month's mean."""
    ent = add_charts._sheet(os.path.join(data_dir, NORWAY_ENTRIES_FILE), "Daily", "date")
    de_cols = [c for c in ent.columns if c.startswith("DE_")]
    if not de_cols or "NO_to_DE" not in nor:
        return None
    gass = nor["NO_to_DE"]
    gass = gass.fillna(gass.groupby([gass.index.year, gass.index.month]).transform("mean"))
    seen = ent[de_cols].sum(axis=1, min_count=1).reindex(gass.index).fillna(0)
    return (gass - seen).clip(lower=0).dropna()


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


def eu_gas_balance(bal, org, dst, storage, lng, gni=None, tso=None, norway_eu=None, biomethane=None):
    """EU27 gas balance: production and consumption summed over the countries; extra-EU pipeline imports and exports
    from the origin / destination sheets; LNG and storage from GIE's EU aggregates. Where a TSO's own consumption series exists
    (`tso`: Germany, France, Spain) it replaces ENTSOG's country total (ENTSOG's own value is used on days the TSO series lacks);
    Ireland is taken from GNI (`gni`): its production, Moffat imports from Great Britain and demand replace ENTSOG's
    incomplete Irish figures. Norwegian pipeline imports (`norway_eu`, Gassco: Germany + France + Belgium + other) replace ENTSOG's
    Norway origin, which captures only about 60% of the flows (632 against 1,032 TWh in 2025)."""
    use_gni = gni is not None and len(gni)
    skip = {"IE"} if use_gni else set()
    tso_cc = [c for c in (tso.columns if tso is not None else []) if c in EU27_GAS]
    skip_cons = skip | set(tso_cc)
    cols = lambda cat, sk=skip: [f"{c}_{cat}_GWhd" for c in EU27_GAS if c not in sk and f"{c}_{cat}_GWhd" in bal]   # noqa: E731
    day = pd.DataFrame(index=bal.index)
    day["Production"] = bal[cols("production")].sum(axis=1, min_count=1).fillna(0)
    day["Pipeline imports"] = org.reindex(bal.index).sum(axis=1, min_count=1) if len(org) else float("nan")
    if norway_eu is not None and len(norway_eu) and len(org) and "NO" in org:
        nrw = norway_eu.reindex(bal.index)
        nrw = nrw.fillna(nrw.groupby([nrw.index.year, nrw.index.month]).transform("mean"))   # gaps take the month's mean
        ok = nrw.notna()
        day.loc[ok, "Pipeline imports"] = day["Pipeline imports"] - org["NO"].reindex(bal.index).fillna(0) + nrw
    cons = bal[cols("distribution", skip_cons) + cols("final_consumers", skip_cons)].sum(axis=1, min_count=1)
    for c in tso_cc:
        own = bal[[f"{c}_distribution_GWhd", f"{c}_final_consumers_GWhd"]].sum(axis=1, min_count=1) if f"{c}_distribution_GWhd" in bal else pd.Series(float("nan"), index=bal.index)
        cons = cons.add(tso[c].reindex(bal.index).combine_first(own), fill_value=0)
    if use_gni:
        g = gni.reindex(bal.index)
        day["Production"] = day["Production"] + g["Production"].fillna(0)
        # ENTSOG's own Irish imports (from Great Britain) are inside the origin sheet; swap them for GNI's Moffat figure
        day["Pipeline imports"] = day["Pipeline imports"] - _col(bal, "IE_imports_GWhd").fillna(0) + g["Moffat"]
        cons = cons.add(g["Consumption"], fill_value=0)
    day["LNG send-out"] = _col(lng, "EU_sendout_GWhd").reindex(bal.index).fillna(0)
    day["Storage withdrawals"] = _col(storage, "EU_withdrawal_GWhd").reindex(bal.index).fillna(0)
    day["Pipeline exports"] = -(dst.reindex(bal.index).sum(axis=1, min_count=1) if len(dst) else 0.0)
    day["Storage injections"] = -_col(storage, "EU_injection_GWhd").reindex(bal.index).fillna(0)
    day["Consumption"] = cons
    cols = GAS_BAL_COLS
    if biomethane is not None and len(biomethane.columns):
        eu_bio = biomethane[[c for c in biomethane.columns if c in EU27_GAS]].reindex(bal.index).sum(axis=1, min_count=1)
        day["Biomethane"] = eu_bio.fillna(0)
        cols = GAS_BAL_COLS[:-1] + ["Biomethane", "Consumption"]
    day = day.dropna(subset=["Pipeline imports", "Consumption"])
    day = day[day.index >= "2021-10-01"]
    return _monthly_twh(day[cols]) if len(day) else pd.DataFrame()


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
        nor_eu = None
        if len(nor) and all(c in nor for c in ("NO_to_DE", "NO_to_FR", "NO_to_BE")):
            nor_eu = nor[[c for c in ("NO_to_DE", "NO_to_FR", "NO_to_BE", "NO_other") if c in nor]].sum(axis=1, min_count=3)
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
        eu = eu_gas_balance(gbal, gorg, gdst, gsto, glng, gni, tso, nor_eu, bio)
        if not eu.empty:
            total_chart(wb, used, gas, 0, eu, ["EU27: production and consumption summed over the countries (ENTSOG; consumption for Germany (THE), France (ODRE) and Spain (Enagas) from the TSOs' own series; Ireland from Gas Networks Ireland, whose Moffat "
                                               "imports from Great Britain replace ENTSOG's incomplete Irish figures; Norwegian pipeline imports from Gassco's flows to "
                                               "Germany, France, Belgium and other, since ENTSOG's Norway origin captures only about 60% of them); biomethane injected into the grids is a separate supply line (France, Denmark, Netherlands and Austria from the operators; other EU27 countries from Eurostat's annual figures, spread over the year and held at the last published year); pipeline imports/exports "
                                               "from/to outside the EU; LNG and storage from GIE's EU aggregates"],
                        "EU gas balance data", "EU gas balance: supply and storage vs consumption (TWh per month)",
                        "TWh per month", GAS_BALANCE_SRC, "Notes:", label="Europe", line_cols=("Consumption",))
        for cc in [c for c in GAS_NAMES if any(col.startswith(f"{c}_") for col in gbal.columns)]:
            fix_kw, fix_cons, fix_note = point_fix_args(args.data_dir, cc, tso, bio, gni)
            cons_in = fix_cons if fix_cons is not None else (tso[cc] if (tso is not None and cc in tso) else None)
            sto_in = fix_kw.pop("storage", gsto)
            b = gas_country_balance(gbal, cc, sto_in, glng, cons_in,
                                    nor["NO_to_GB"] if (cc == "UK" and len(nor) and "NO_to_GB" in nor) else None,
                                    bio[cc] if (len(bio) and cc in bio) else None, **fix_kw)
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
            if cc in ("DE", "NL") and emden is not None:
                note = (note or "Supply less exports and storage injections against consumption.") + (
                    " Excludes the Norwegian gas that arrives at Emden: ENTSOG publishes nothing at Emden and the gas feeds both the German and "
                    "the Dutch grids, so this balance is short on its own; see the Germany + Netherlands balance.")
            if cc == "UK" and len(nor) and "NO_to_GB" in nor:
                note += (" Pipeline imports are Gassco's Norway-to-Great-Britain flows; the rest of the St Fergus and Easington entry points "
                         "(UK North Sea gas) is counted in production.")
            if fix_note:
                note = (note or "Supply less exports and storage injections against consumption.") + fix_note
            if cc == "DK":
                try:
                    b, note = denmark_gas_balance(args.data_dir), (
                        "Denmark from Energinet's own Gasflow dataset: the North Sea entry (zero until Baltic Pipe started in Oct 2022, then "
                        "mainly Norwegian gas for Baltic Pipe) and Tyra (Danish fields, small until 2024), biomethane, storage and the German border against exports to Poland and "
                        "Sweden; consumption is gas delivered to Danish consumers (including the biomethane). ENTSOG captures only a small part "
                        "of these flows.")
                except Exception as e:  # noqa: BLE001
                    gas[2].append(f"Denmark gas balance from Energinet failed ({type(e).__name__}: {e}); ENTSOG used")
            if cc == "IE":
                try:
                    b, note = ireland_gas_balance(args.data_dir), (
                        "Republic of Ireland from Gas Networks Ireland's own series: Corrib/Inch production and Moffat imports "
                        "(from Great Britain, ROI share) against ROI demand. ENTSOG's Irish totals miss most of Moffat.")
                except Exception as e:  # noqa: BLE001
                    gas[2].append(f"Ireland gas balance from GNI failed ({type(e).__name__}: {e}); ENTSOG used")
            if b.empty:
                gas[2].append(f"{GAS_NAMES[cc]} gas balance: too little data")
                continue
            total_chart(wb, used, gas, None, b, [note or ("Supply (production, pipeline imports, LNG, storage withdrawals) less exports and "
                                                 "storage injections should land near consumption (distribution + final consumers); "
                                                 "the gap is unreported or unclassified flow")],
                        f"{GAS_NAMES[cc]} gas balance data", f"{GAS_NAMES[cc]} gas balance: supply and storage vs consumption",
                        "TWh per month", GAS_BALANCE_SRC, "Notes:", label=GAS_NAMES[cc], line_cols=("Consumption",))
    if len(gbal) and emden is not None:
        try:   # Germany + Netherlands: Gassco's Emden gas is split between them in a way the raw data cannot show, so they are combined
            de_b = gas_country_balance(gbal, "DE", gsto, glng, tso["DE"] if "DE" in tso else None, None,
                                       bio["DE"] if (len(bio) and "DE" in bio) else None, emden)
            nl_b = gas_country_balance(gbal, "NL", gsto, glng, None, None, bio["NL"] if (len(bio) and "NL" in bio) else None)
            colsb = [c for c in de_b.columns if c in nl_b.columns]
            both = de_b[colsb].add(nl_b[colsb], fill_value=0)
            if len(both) >= 12:
                total_chart(wb, used, gas, None, both, [
                    "Germany + Netherlands combined. Includes the Norwegian gas that arrives at Emden (Gassco's flow to Germany minus the Dornum "
                    "volume ENTSOG reports); ENTSOG publishes nothing at Emden, and the gas feeds both grids, so the two countries are shown together. "
                    "Flows between the two countries are counted on both sides and do not cancel exactly."],
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
