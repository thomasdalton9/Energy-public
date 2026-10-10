"""
South & Central America master workbook: one file with every South and
Central American dataset this repo pulls (Central America = Guatemala to
Panama, plus Belize; Mexico is not included), and Dashboard front pages
carrying all their charts.

  Dashboard          - gas: title, index of charts (latest month, data
                       tab), then every chart (native Excel, no borders, mmm/yy)
  Dashboard - Power & Hydro - power generation by type per country, plus
                       reservoir water-year charts. Raw grid-operator data
                       (ONS, XM, CAMMESA, ...) is used wherever the repo pulls
                       it; Ember fills only the countries without a raw feed.
  <CC> <chart> data  - the table each Dashboard chart plots
  <CC> <dataset> raw - the full data sheet from each source workbook
  Sources            - where each dataset comes from, units and notes

Both dashboards name each chart's source (publisher + link) in the index
and under the chart.

Reads (doesn't refetch) the workbooks the scheduled South/Central America pulls
write to "output/Data and Chart Outputs/". Chart definitions come from
add_charts.py's registry, so the master and each country workbook always
show the same charts. A missing input is listed on the Dashboard and
skipped rather than stopping the rest.

Usage: python3 SOUTH_AMERICA_MASTER.py [--out "output/Data and Chart Outputs/Master Outputs/south_and_central_america_master.xlsx"]
"""
import argparse
import os
import sys
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import add_charts  # noqa: E402
import capacity_factors  # noqa: E402
import water_year_chart  # noqa: E402
import fundamentals  # noqa: E402
import xlsx_charts  # noqa: E402

DATA_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
DARK_BLUE = "17365D"

# (country code, country, workbook, raw sheet to copy, short dataset name)
DATASETS = [
    ("SA", "South America", "south_america_gas_balance.xlsx", "Balance (long)", "gas balance"),
    ("SA", "South & Central America", "south_america_gas_burn_power.xlsx", ("Monthly (Bcf per day)", "Countries"),
     "gas burn for power"),
    ("AR", "Argentina", "argentina_gas_monthly.xlsx", "National", "gas"),
    ("BR", "Brazil", "brazil_gas_monthly.xlsx", "Demand by segment", "gas"),
    ("BO", "Bolivia", "bolivia_gas_demand_by_sector.xlsx", "Demand by sector", "gas"),
    ("CL", "Chile", "chile_gas_imports.xlsx", "Gas imports", "gas imports"),
    ("CO", "Colombia", "colombia_gas_demand_by_sector.xlsx", "Demand by sector", "gas"),
    ("EC", "Ecuador", "ecuador_gas.xlsx", "Gas by use", "gas"),
    ("PA", "Panama", "panama_gas.xlsx", "Gas use", "gas"),
    ("PE", "Peru", "peru_gas_demand_by_sector.xlsx", "Demand by sector", "gas"),
    ("UY", "Uruguay", "uruguay_gas_demand_by_sector.xlsx", "Demand by sector", "gas"),
    ("TT", "Trinidad & Tobago", "trinidad_gas.xlsx", "Utilization by sector", "gas"),
    ("SV", "El Salvador", "el_salvador_gas.xlsx", "Gas use", "gas"),
    # Caribbean
    ("PR", "Puerto Rico", "puerto_rico_gas.xlsx", "Gas use", "gas"),
    ("JM", "Jamaica", "jamaica_gas.xlsx", "Gas use", "gas"),
    ("DO", "Dominican Republic", "dominican_republic_gas.xlsx", "Gas use", "gas"),
    ("CO", "Colombia", "south_america_coal_production.xlsx", "Colombia", "coal"),
]

# Second dashboard: (code, country, workbook, raw sheet or "*" for every data sheet, short name).
# Raw grid-operator generation comes first; Ember (one sheet per country) only charts
# the countries with no raw workbook this run - see main().
EMBER_FILES = [("SA", "South America", "south_america_power_by_type.xlsx"),
               ("CA", "Central America", "central_america_power_by_type.xlsx")]
EMBER = {f for _, _, f in EMBER_FILES}
RAW_POWER_DATASETS = [
    ("AR", "Argentina", "argentina_power_generation_daily.xlsx", "Daily", "power"),
    ("BO", "Bolivia", "bolivia_power_generation_daily.xlsx", "Daily", "power"),
    ("BR", "Brazil", "brazil_power_generation_daily.xlsx", "Daily", "power"),
    ("CL", "Chile", "chile_power_generation_daily.xlsx", "Daily", "power"),
    ("CO", "Colombia", "colombia_power_generation_daily.xlsx", "Daily", "power"),
    ("EC", "Ecuador", "ecuador_power_generation_daily.xlsx", "Daily", "power"),
    ("PE", "Peru", "peru_power_generation_daily.xlsx", "Daily", "power"),
    ("UY", "Uruguay", "uruguay_power_generation_daily.xlsx", "Daily", "power"),
    ("PY", "Paraguay", "paraguay_power_generation_daily.xlsx", "Daily", "power"),
    # Central America
    ("BZ", "Belize", "belize_power_generation_daily.xlsx", "Daily", "power"),
    ("CR", "Costa Rica", "costa_rica_power_generation_daily.xlsx", "Daily", "power"),
    ("SV", "El Salvador", "el_salvador_power_generation_daily.xlsx", "Daily", "power"),
    ("GT", "Guatemala", "guatemala_power_generation_daily.xlsx", "Daily", "power"),
    ("HN", "Honduras", "honduras_power_generation_daily.xlsx", "Daily", "power"),
    ("NI", "Nicaragua", "nicaragua_power_generation_daily.xlsx", "Daily", "power"),
    ("PA", "Panama", "panama_power_generation_daily.xlsx", "Daily", "power"),
    # Caribbean
    ("PR", "Puerto Rico", "puerto_rico_power_generation_daily.xlsx", "Daily", "power"),
    ("JM", "Jamaica", "jamaica_power_generation_daily.xlsx", "Daily", "power"),
    ("DO", "Dominican Republic", "dominican_republic_power_generation_daily.xlsx", "Daily", "power"),
]
HYDRO_DATASETS = [
    ("AR", "Argentina", "argentina_hydro_reservoirs.xlsx", "Daily", "hydro"),
    ("BR", "Brazil", "brazil_hydro_reservoirs.xlsx", "Daily", "hydro"),
    ("CL", "Chile", "chile_hydro_reservoirs.xlsx", "Daily", "hydro"),
    ("CO", "Colombia", "colombia_hydro_reservoirs.xlsx", "Daily", "hydro"),
    ("EC", "Ecuador", "ecuador_hydro_reservoirs.xlsx", "Daily", "hydro"),
    ("PA", "Panama", "gatun_lake_level.xlsx", "Daily", "hydro"),
    ("PE", "Peru", "peru_hydro_reservoirs.xlsx", "Daily", "hydro"),
    ("UY", "Uruguay", "uruguay_hydro_reservoirs.xlsx", "Daily", "hydro"),
    # NASA POWER rainfall at one point per catchment, cumulative since 1 Oct vs the 5 previous water years
    ("RF", "South America", "south_america_rainfall_daily.xlsx", ("Monthly", "Daily"), "rainfall"),
]
# The Power & Hydro dashboard shows one national hydro chart per country (each workbook's first water-year
# spec) plus these extra regional charts by spec name; the full sets stay in each country workbook.
HYDRO_EXTRA = {"brazil_hydro_reservoirs.xlsx": {"N"},
               # every catchment's rainfall chart, not just the first
               "south_america_rainfall_daily.xlsx": {"BR_SECO", "BR_S", "BR_NE", "BR_N", "CO", "EC", "PE", "CL_LAJA",
                                                     "CL_MAULE", "AR_COMAHUE", "UY", "PY_PARANA", "PA"}}
# Installed generation capacity by technology (standard <country>_power_capacity.xlsx, sheet "Monthly")
CAPACITY_DATASETS = [(code, country, f"{country.lower()}_power_capacity.xlsx", "Monthly", "capacity")
                     for code, country in [("AR", "Argentina"), ("BO", "Bolivia"), ("BR", "Brazil"), ("CL", "Chile"),
                                           ("CO", "Colombia"), ("EC", "Ecuador"), ("PE", "Peru"), ("UY", "Uruguay"),
                                           ("PY", "Paraguay")]]

# Workbooks with several charts where the dashboard shows only some (by spec name); the rest stay in the
# country workbook. Honduras: the monthly ODS history (2021 on) rather than the daily feed (June 2026 on).
DASHBOARD_ONLY = {"honduras_power_generation_daily.xlsx": {"History"},
                  "south_america_gas_burn_power.xlsx": {"Burn"}}   # per-country checks stay in that workbook


def gatun(path):
    """gatun_lake_level.xlsx draws its own chart (REGISTRY entry is None); the master needs a spec."""
    d = add_charts.by_date(add_charts.read(path, "Daily"), "date")
    return [{"name": "Gatun", "water_year": d["level_ft"], "title": "Panama Canal, Gatun Lake level",
             "units": "feet above sea level", "y_decimals": 0}]


# Chart specs for workbooks whose add_charts REGISTRY entry is None (they chart themselves)
MASTER_SPECS = {"gatun_lake_level.xlsx": gatun}

# Where each workbook's data comes from: (publisher, link). Shown on the dashboards.
SOURCES = {
    "argentina_gas_monthly.xlsx": ("Secretaría de Energía (Argentina), Series de Tiempo API (production, demand); "
                                   "ENARGAS (enargas.gob.ar): daily export and import reports, gas received by "
                                   "basin (GRT), deliveries (GETD), daily transport reports (net supply)",
                                   "https://datos.gob.ar/series/api/series/"),
    "brazil_gas_monthly.xlsx": ("MME Brazil gas bulletin (segments, to Jun-2025); ANP open data (pipeline flows, production, imports)",
                                "https://www.gov.br/anp/pt-br/centrais-de-conteudo/dados-abertos/dados-consolidados-movimentacao-de-gas-natural-em-gasodutos-de-transporte"),
    "bolivia_gas_demand_by_sector.xlsx": ("INE Bolivia",
                                          "https://www.ine.gob.bo/index.php/estadisticas-economicas/hidrocarburos-mineria/hidrocarburo-cuadros-estadisticos/"),
    "chile_gas_imports.xlsx": ("CNE Chile: Estadísticas > Hidrocarburo (import workbook by use and region, customs "
                               "data; ENAP + CEOP gas production, Ministerio de Energía data); Reporte Mensual "
                               "(imports for months after the workbook)",
                               "https://www.cne.cl/estadisticas/hidrocarburo/"),
    "colombia_gas_demand_by_sector.xlsx": ("Gestor del Mercado de Gas (BMC)", "https://www.bmcbec.com.co/informes/informes-mensuales"),
    "ecuador_gas.xlsx": ("EP Petroecuador (domestic gas); UN Comtrade, Ecuador customs HS 2711.11 (LNG imports)",
                         "https://www.eppetroecuador.ec/?p=3721"),
    "panama_gas.xlsx": ("Estimate from CND / ETESA Panama gas-fired generation (daily report) x 7.0 MMBtu/MWh heat rate",
                        "https://www.cnd.com.pa/index.php/informes/categoria/informes-de-operaciones"),
    "el_salvador_gas.xlsx": ("ESTIMATE: SIGET monthly LNG-fired generation (Ember/UT for months SIGET lacks) x 8.2 MMBtu/MWh heat rate",
                             "https://www.siget.gob.sv/gerencias/electricidad/informe-de-mercado-y-estadisticas-electricas/estadisticas-electricas-bi/"),
    "trinidad_gas.xlsx": ("Ministry of Energy and Energy Industries (MEEI), monthly bulletins",
                          "https://www.energy.gov.tt/category/publications/energy-industry-bulletins/"),
    "uruguay_gas_demand_by_sector.xlsx": ("MIEM Uruguay, VisualPEB energy balance", "https://visualpeb.miem.gub.uy/visualPEB/gas_natural"),
    "peru_gas_demand_by_sector.xlsx": ("MINEM Peru (DGH), Informes Estadisticos Upstream - Downstream",
                                       "https://www.gob.pe/institucion/minem/colecciones/17643-informes-estadisticos-upstream-downstream"),
    "argentina_power_generation_daily.xlsx": ("CAMMESA", "https://cammesaweb.cammesa.com/"),
    "bolivia_power_generation_daily.xlsx": ("CNDC Bolivia", "https://www.cndc.bo/"),
    "brazil_power_generation_daily.xlsx": ("ONS Brazil open data", "https://dados.ons.org.br/"),
    "colombia_power_generation_daily.xlsx": ("XM Colombia", "https://www.xm.com.co/"),
    "ecuador_power_generation_daily.xlsx": ("CENACE Ecuador", "https://www.cenace.gob.ec/info-operativa/InformacionOperativa.htm"),
    "peru_power_generation_daily.xlsx": ("COES Peru", "https://www.coes.org.pe/"),
    "chile_power_generation_daily.xlsx": ("CNE Chile, Generación Bruta workbook (daily, from Coordinador Eléctrico Nacional data)",
                                          "https://www.cne.cl/estadisticas/electricidad/"),
    "uruguay_power_generation_daily.xlsx": ("ADME Uruguay", "https://pronos.adme.com.uy/"),
    "paraguay_power_generation_daily.xlsx": ("ITAIPU Binacional monthly production reports (generation, supply to "
                                             "ANDE), ONS (Itaipu to Brazil), EBY / CAMMESA (Yacyreta to the SADI and "
                                             "SINP), VMME Balance Energetico; monthly, Paraguay's 50% share",
                                             "https://www.itaipu.gov.py/noticias/energia/"),
    "brazil_hydro_reservoirs.xlsx": ("ONS Brazil open data (EAR)", "https://dados.ons.org.br/dataset/ear-diario-por-subsistema"),
    "colombia_hydro_reservoirs.xlsx": ("XM Colombia", "https://www.xm.com.co/"),
    "chile_hydro_reservoirs.xlsx": ("DGA Chile (Visualizador Hidrométrico Nacional; monthly Boletín Hidrométrico before mid-2024)",
                                    "https://vipnet.mop.gob.cl/"),
    "argentina_hydro_reservoirs.xlsx": ("CAMMESA (daily lake levels and river flows, weekly programme), AIC, INA",
                                        "https://cammesaweb.cammesa.com/download/cotas-diarias/"),
    "peru_hydro_reservoirs.xlsx": ("COES Peru, Informe Semanal de Evaluación de la Operación (5.1 useful volume of "
                                   "reservoirs and lagoons)",
                                   "https://www.coes.org.pe/Portal/PostOperacion/Informes/EvaluacionSemanal"),
    "ecuador_hydro_reservoirs.xlsx": ("CELEC EP - CELEC SUR, Gráficas de Producción (Mazar / Amaluza SCADA levels)",
                                      "https://generacioncsr.celec.gob.ec/graficasproduccion/"),
    "south_america_rainfall_daily.xlsx": ("NASA POWER (Langley Research Center) daily PRECTOTCORR rainfall - a satellite/"
                                          "reanalysis-derived (MERRA-2 / IMERG) value for the grid cell at ONE point of each "
                                          "catchment, NOT a rain gauge or a basin average",
                                          "https://power.larc.nasa.gov/docs/services/api/temporal/daily/"),
    "uruguay_hydro_reservoirs.xlsx": ("ADME Uruguay (Río Negro lake levels, SCADA); INA Argentina (Salto Grande lake level)",
                                      "https://pronos.adme.com.uy/seriesbonete.php"),
    "belize_power_generation_daily.xlsx": ("Belize Electricity Ltd (BEL)", "https://www.bel.com.bz/"),
    "costa_rica_power_generation_daily.xlsx": ("ICE / CENCE Costa Rica", "https://apps.grupoice.com/CenceWeb/"),
    "el_salvador_power_generation_daily.xlsx": ("SIGET El Salvador, Estadisticas Electricas (Power BI), monthly net generation",
                                                "https://www.siget.gob.sv/gerencias/electricidad/informe-de-mercado-y-estadisticas-electricas/estadisticas-electricas-bi/"),
    "guatemala_power_generation_daily.xlsx": ("Administrador del Mercado Mayorista (AMM) Guatemala",
                                              "https://www.amm.org.gt/"),
    "honduras_power_generation_daily.xlsx": ("Operador del Sistema (ODS) Honduras", "https://www.ods.org.hn/"),
    "nicaragua_power_generation_daily.xlsx": ("CNDC / ENATREL Nicaragua", "https://www.cndc.org.ni/"),
    "panama_power_generation_daily.xlsx": ("CND / ETESA Panama", "https://www.cnd.com.pa/"),
    "argentina_power_capacity.xlsx": ("CAMMESA, Potencia Instalada", "https://cammesaweb.cammesa.com/download/potencia-instalada/"),
    "bolivia_power_capacity.xlsx": ("CNDC Bolivia, potencia (annual, December)", "https://www.cndc.bo/"),
    "brazil_power_capacity.xlsx": ("ANEEL SIGA plant register + MMGD distributed generation", "https://dadosabertos.aneel.gov.br/"),
    "chile_power_capacity.xlsx": ("CNE Chile, Capacidad Instalada de Generacion", "https://www.cne.cl/estadisticas/electricidad/"),
    "colombia_power_capacity.xlsx": ("XM Colombia, net effective capacity (CapEfecNeta)", "https://www.xm.com.co/"),
    "ecuador_power_capacity.xlsx": ("ARCONEL, Balance Nacional de Energia Electrica", "https://arconel.gob.ec/balance-nacional-de-energia-electrica/"),
    "peru_power_capacity.xlsx": ("COES Peru, annual statistics (effective capacity, SEIN)", "https://www.coes.org.pe/Portal/publicaciones/estadisticas/"),
    "paraguay_power_capacity.xlsx": ("ITAIPU (14,000 MW) and EBY (3,200 MW) at 50%; ANDE's own plants from Ember "
                                     "(ande.gov.py is captcha-protected) - annual", "https://www.eby.gov.py/datos-tecnicos/"),
    "uruguay_power_capacity.xlsx": ("MIEM / DNE Uruguay, potencia instalada (annual)",
                                    "https://www.gub.uy/ministerio-industria-energia-mineria/datos-y-estadisticas/datos/series-estadisticas-energia-electrica"),
    "gatun_lake_level.xlsx": ("Panama Canal Authority (ACP)", "https://evtms-rpts.pancanal.com/"),
    # Caribbean
    "puerto_rico_power_generation_daily.xlsx": ("EIA (US Energy Information Administration), EIA-923 monthly generation by fuel, Puerto Rico",
                                                "https://www.eia.gov/electricity/data/browser/"),
    "puerto_rico_gas.xlsx": ("EIA, EIA-923 plant-level natural gas consumption, Puerto Rico (monthly)",
                             "https://www.eia.gov/electricity/data/browser/"),
    "dominican_republic_power_generation_daily.xlsx": ("Organismo Coordinador del SENI (OC), daily generation by fuel",
                                                       "https://www.oc.org.do/Servicios/Reporte"),
    "dominican_republic_gas.xlsx": ("Superintendencia de Electricidad (SIE), monthly fuel consumption for power (datos.gob.do)",
                                    "https://datos.gob.do/dataset/energia-y-potencia-facturadas-ede"),
    "jamaica_power_generation_daily.xlsx": ("Ministry of Energy (MSET), Jamaica Energy Statistics Table 8 - JPS + IPPs, ANNUAL only",
                                            "https://www.mset.gov.jm/document-category/statistics-data/"),
    "jamaica_gas.xlsx": ("Ministry of Energy (MSET), Jamaica Energy Statistics Table 3 natural gas, ANNUAL only",
                         "https://www.mset.gov.jm/document-category/statistics-data/"),
    "south_america_gas_burn_power.xlsx": ("Published power-sector gas (SE/ENARGAS, MME, INE Bolivia, BMC Colombia, "
                                          "MINEM Peru, SIE, EIA-923) and, for months not yet published, an ESTIMATE = "
                                          "grid-operator gas-fired output x calibrated heat rate (flat assumed rate "
                                          "for Chile, Panama, El Salvador) - SA_GAS_BURN_POWER.py",
                                          "see each country's row"),
    "south_america_gas_balance.xlsx": ("Built from the country workbooks above (each official source: SE/ENARGAS, INE "
                                       "Bolivia, ANP, CNE Chile, Colombia supply report, Petroecuador/ARCERNNR, "
                                       "Perupetro/MINEM, MEEI Trinidad, URSEA) - no estimates",
                                       "see each country's row"),
    "south_america_coal_production.xlsx": ("ANM Colombia (coal production declared for royalties, datos.gov.co) and "
                                           "DANE (coal exports); EPE, SE Argentina, Cochilco, MINEM Peru annual; "
                                           "Venezuela: Energy Institute Statistical Review",
                                           "https://www.datos.gov.co/d/r85m-vv6c"),
    **{f: ("Ember monthly electricity data", "https://ember-energy.org/data/monthly-electricity-data/") for f in EMBER},
}
# The grid operator Ember compiles each country from (named on Ember-fed charts)
OPERATORS = {"Argentina": "CAMMESA", "Bolivia": "CNDC", "Brazil": "ONS", "Chile": "Coordinador Eléctrico Nacional",
             "Colombia": "XM", "Ecuador": "CENACE", "Peru": "COES", "Uruguay": "ADME", "Paraguay": "ANDE",
             "Belize": "BEL", "Costa Rica": "ICE/CENCE", "El Salvador": "UT", "Guatemala": "AMM", "Honduras": "ODS",
             "Nicaragua": "CNDC", "Panama": "CND",
             "Dominican Republic": "OC-SENI", "Jamaica": "JPS", "Puerto Rico": "PREPA/Genera PR"}


MIN_RAW_DAYS = 365   # a raw generation workbook replaces Ember only once it spans a year of history


SA_POWER_COUNTRIES = ["Argentina", "Bolivia", "Brazil", "Chile", "Colombia", "Ecuador", "Peru", "Uruguay", "Paraguay"]
# Countries whose SA-total contribution comes from a different sheet of their workbook than the dashboard chart.
# Paraguay: its workbook's 'Daily' sheet is Paraguay's 50% share of Itaipu and Yacyreta, but ONS already counts
# Itaipu's whole supply to Brazil (60 Hz + the 50 Hz energy Paraguay cedes) and CAMMESA all of Yacyreta's supply
# to Argentina (YACYHI + YACYHIPY). Only ANDE's own take is counted nowhere else, so the total adds that sheet.
SA_TOTAL_SHEET = {"Paraguay": "Not counted by ONS-CAMMESA"}
# Monthly feeds published ~10 days after month end: the total does not wait for them (like an Ember-fed
# country, a missing latest month is listed under NOT INCLUDED rather than holding the total back).
SA_LAGGING = {"Paraguay"}


def sa_total_sheet(path, sheet):
    """A standard-layout sheet (date + <Fuel>_MWh) other than 'Daily' -> monthly GWh frame like power_daily's."""
    d = add_charts.by_date(add_charts.read(path, sheet), "date")
    m = d[[c for c in d.columns if str(c).endswith("_MWh") and c != "Total_MWh"]].resample("MS").sum(min_count=1) / 1000
    m = m[m.index >= "2021-01-01"].rename(columns=lambda c: c.replace("_MWh", "_GWh"))
    m = m.rename(columns={"Oil_GWh": "Other Fossil_GWh", "Other_GWh": "Other Renewables_GWh"})
    return add_charts.power_mix(m)


def south_america_generation(data_dir, have_raw, frames_out=None):
    """South America generation by source, TWh per month: the sum of each country's monthly mix from the same
    series the dashboard shows (raw grid operator where used, Ember otherwise). Only months every country has
    in full (a daily feed's current month is left out until it is complete). Returns (frame, per-country notes)."""
    raw_files = {d[1]: d[2] for d in RAW_POWER_DATASETS}
    ember = os.path.join(data_dir, "south_america_power_by_type.xlsx")
    frames, notes = {}, []
    for country in SA_POWER_COUNTRIES:
        try:
            if country in have_raw:
                path = os.path.join(data_dir, raw_files[country])
                sheet = SA_TOTAL_SHEET.get(country, "Daily")
                if sheet != "Daily":
                    m = sa_total_sheet(path, sheet)
                else:
                    build = add_charts.REGISTRY.get(raw_files[country]) or add_charts.power_daily(country)
                    m = build(path)[0]["df"]
                dates = add_charts.by_date(add_charts.read(path, sheet), "date").index
                last = dates.max()
                monthly_rows = len(dates) > 1 and dates.to_series().diff().median().days > 25
                if not monthly_rows and last < last + pd.offsets.MonthEnd(0):   # current month not complete yet
                    m = m[m.index < last.to_period("M").to_timestamp()]
                src = SOURCES.get(raw_files[country], (raw_files[country],))[0]
                if sheet != "Daily":
                    src = f"{src} [sheet '{sheet}': ANDE's own take only, the rest is in ONS/CAMMESA]"
            else:
                m = add_charts.power_mix(add_charts.by_date(add_charts.read(ember, country), "Month"))
                src = f"Ember (compiled from {OPERATORS.get(country, 'the grid operator')})"
            m = m.apply(pd.to_numeric, errors="coerce")
            frames[country] = m
            notes.append(f"{country}: {src}, {m.index.min():%b/%y}-{m.index.max():%b/%y}")
        except Exception as e:  # noqa: BLE001
            notes.append(f"{country}: not available ({type(e).__name__}: {e})")
    if frames_out is not None:
        frames_out.update(frames)
    if not frames:
        return pd.DataFrame(), notes
    fuels = ["Hydro", "Gas", "Wind", "Solar", "Coal", "Nuclear", "Other"]
    have = {c: set(f.dropna(how="all").index) for c, f in frames.items()}
    # Run to the last month every raw-fed country has; an Ember-fed country that lags (Ember publishes months
    # late) is left out of the months it doesn't have yet, and those gaps are listed rather than estimated.
    raw_have = [h for c, h in have.items() if c in have_raw and c not in SA_LAGGING] or list(have.values())
    end = min(max(h) for h in raw_have)
    months = sorted(m for m in set.intersection(*raw_have) if m <= end)
    total = sum(f.reindex(index=months, columns=fuels).fillna(0) for f in frames.values()) / 1000.0   # GWh -> TWh
    total.index.name = "date"
    for c, h in have.items():
        gap = [m for m in months if m not in h]
        if gap:
            notes.append(f"NOT INCLUDED: {c} in {gap[0]:%b/%y}-{gap[-1]:%b/%y} ({len(gap)} months, no data yet)")
    return total, notes


def south_america_capacity(data_dir, cfg=None):
    """South America installed capacity by technology, GW: the sum of each country's capacity workbook. Capacity is
    a stock, so an annual-only country's figure is carried forward month by month until its next value. Runs to the
    last month every monthly-source country has; a country with no workbook is listed, not estimated."""
    import add_charts as ac
    cfg = cfg or sys.modules[__name__]   # the North America master passes itself
    frames, notes, monthly_last = {}, [], []
    for code, country, fname, _, _ in cfg.CAPACITY_DATASETS:
        path = os.path.join(data_dir, fname)
        if not os.path.exists(path):
            notes.append(f"NOT INCLUDED: {country} (no capacity workbook yet)")
            continue
        d = ac.by_date(ac.read(path, "Monthly"), "date")
        d = d[d.index >= "2021-01-01"]
        g = ac.capacity_groups(pd.DataFrame({f: d.get(f"{f}_MW") for f in ac.CAPACITY_FUELS}, index=d.index)) / 1000.0
        annual = len(d) > 1 and d.index.to_series().diff().median().days > 300
        if annual:   # annual rows hold the year-END value but are dated 1 January: place them in December
            g.index = g.index + pd.DateOffset(months=11)
        else:
            monthly_last.append(g.index.max())
        frames[country] = (g, annual)
        notes.append(f"{country}: {cfg.SOURCES.get(fname, (fname,))[0]}, {'annual' if annual else 'monthly'} "
                     f"{g.index.min():%b/%y}-{g.index.max():%b/%y}")
    if not frames:
        return pd.DataFrame(), notes
    end = min(monthly_last) if monthly_last else max(g.index.max() for g, _ in frames.values())
    months = pd.date_range("2021-01-01", end, freq="MS")
    total = sum(g.reindex(g.index.union(months)).sort_index().ffill().bfill().reindex(months).fillna(0)
                for g, _ in frames.values())   # bfill: months before a country's first figure take that figure
    total.index.name = "date"
    return total, notes


def raw_history_days(path):
    """Days spanned by the raw feed (first to last date), so monthly or annual feeds with a year of history
    qualify as well as daily ones."""
    try:
        d = pd.to_datetime(pd.read_excel(path, sheet_name="Daily", usecols=[0]).iloc[:, 0], errors="coerce").dropna()
        return int((d.max() - d.min()).days) + 1 if len(d) else 0
    except Exception:
        return 0


def source_of(fname, spec_name=None, cfg=None):
    cfg = cfg or sys.modules[__name__]
    publisher, url = cfg.SOURCES.get(fname, (fname, None))
    if fname in cfg.EMBER and spec_name in cfg.OPERATORS:
        publisher = f"Ember, compiled from {cfg.OPERATORS[spec_name]} (no raw feed yet)"
    return publisher, url

GAS_BCFD = True   # collect() shows gas volume charts in Bcf/d for this master only (other masters pass their own cfg)
CHART_W, CHART_H = 21.0, 11.0      # cm - room for the title, rotated date labels and axis titles
ROWS_PER_CHART = 24                # 11 cm is ~21 default rows, plus the source line and a gap
COLS = ("B", "F")                  # two charts per row (B-E are sized to ~21 cm, so F starts the second)


def sheet_name(text, used):
    name = text.replace("/", "-")[:31]
    base, k = name, 2
    while name in used:
        name = f"{base[:28]} {k}"
        k += 1
    used.add(name)
    return name


def write_frame(ws, df):
    ws.append([str(c) for c in df.columns])
    for row in df.itertuples(index=False):
        ws.append([None if (not isinstance(v, str) and pd.isna(v)) else
                   (v.to_pydatetime() if isinstance(v, pd.Timestamp) else v) for v in row])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"


def notes_text(path):
    try:
        u = pd.read_excel(path, sheet_name=0, header=None)
        return " ".join(str(x) for x in u.iloc[:, 0].dropna().tolist())
    except Exception:
        return ""


def collect(wb, datasets, data_dir, used, sources, skip=(), cfg=None):
    """Data + raw tabs for one dashboard's datasets; returns (charts, index rows, missing).
    Each chart is (chart, (publisher, url)); index rows end with the same source.
    skip: spec names not to chart (Ember countries that have a raw workbook).
    cfg: the module holding SOURCES, MASTER_SPECS, HYDRO_DATASETS, ... (this one by default; the North America
    master passes itself)."""
    cfg = cfg or sys.modules[__name__]
    charts, index_rows, missing = [], [], []
    for code, country, fname, raw_sheet, short in datasets:
        path = os.path.join(data_dir, fname)
        if not os.path.exists(path):
            missing.append(f"{country} {short} ({fname})")
            continue
        try:
            build = cfg.MASTER_SPECS.get(fname) or add_charts.REGISTRY.get(fname) or (
                add_charts.power_daily(f"{country} power generation by type")
                if fname.endswith("_power_generation_daily.xlsx") else add_charts.generic)
            specs = build(path)
            if fname in cfg.DASHBOARD_ONLY:
                specs = [sp for sp in specs if sp["name"] in cfg.DASHBOARD_ONLY[fname]] or specs
            if (code, country, fname, raw_sheet, short) in cfg.HYDRO_DATASETS:
                specs = [sp for i, sp in enumerate(specs) if i == 0 or sp["name"] in cfg.HYDRO_EXTRA.get(fname, ())]
        except Exception as e:
            missing.append(f"{country} {short} ({fname}: {type(e).__name__}: {e})")
            continue
        for s in specs:
            if s["name"] in skip:
                continue
            src = source_of(fname, s["name"], cfg)
            if "water_year" in s or "week_year" in s:
                weekly = "week_year" in s
                table, meta = (water_year_chart.week_year_table(s["week_year"]) if weekly
                               else water_year_chart.water_year_table(s["water_year"]))
                ws = wb.create_sheet(sheet_name(f"{code} {s['name']} data", used))
                water_year_chart.write_table(ws, table)
                charts.append((water_year_chart.build_chart(ws, table, meta, s["title"], s["units"],
                                                            width=CHART_W, height=CHART_H, gridlines=False,
                                                            inner=xlsx_charts.DASHBOARD_INNER, short_title=True,
                                                            y_decimals=s.get("y_decimals")), src))
                index_rows.append((country, f"{s['title']} ({'calendar-year weeks' if weekly else 'water year'})",
                                   pd.Timestamp(meta["last"]).strftime("%d/%m/%y"),
                                   ws.title, *src))
                continue
            s_df, s_units = xlsx_charts.monthly_energy_to_gw(s["df"].dropna(how="all"), s["units"], s["title"])
            if getattr(cfg, "GAS_BCFD", False):   # South America master: gas volumes in Bcf/d (owner's choice, Oct 2026)
                s_df, s_units = xlsx_charts.gas_volume_to_bcfd(s_df, s_units, s["title"])
            df, n_bars = xlsx_charts.prepare(s_df, s.get("line_cols", ()))
            if df.empty:
                continue
            label = f"{s['name']} {short} data" if raw_sheet == "*" else f"{code} {s['name']} data"
            ws = wb.create_sheet(sheet_name(label, used))
            xlsx_charts.write_table(ws, df, s["date_format"])
            charts.append((xlsx_charts.build_chart(ws, df, n_bars, s["title"], s_units, s["kind"],
                                                   s["date_format"], width=CHART_W, height=CHART_H, gridlines=False,
                                                   inner=xlsx_charts.DASHBOARD_INNER,
                                                   forecast_from=s.get("forecast_from"),
                                                   scenario_from=s.get("scenario_from")), src))
            index_rows.append((s["name"] if raw_sheet == "*" else country, s["title"], df.index.max().strftime("%b/%y"),
                               ws.title, *src))
        # raw_sheet: one sheet, "*" (every data sheet, one per country) or a tuple of sheets
        raw_sheets = ([n for n in pd.ExcelFile(path).sheet_names
                       if n.lower() not in ("units", "notes") and not n.startswith("Chart")
                       and n != water_year_chart.SHEET and n not in skip] if raw_sheet == "*"
                      else list(raw_sheet) if isinstance(raw_sheet, tuple) else [raw_sheet])
        for rs in raw_sheets:
            try:
                raw = add_charts.read(path, rs)
                label = (f"{rs} {short} raw" if raw_sheet == "*" else
                         f"{code} {rs} raw" if isinstance(raw_sheet, tuple) else f"{code} {short} raw")
                write_frame(wb.create_sheet(sheet_name(label, used)), raw)
            except Exception as e:
                missing.append(f"{country} {short} raw sheet {rs} ({e})")
        sources.append((country, short, fname, *source_of(fname, cfg=cfg), notes_text(path)))
    return charts, index_rows, missing


def draw_dashboard(dash, heading, charts, index_rows, missing):
    dash["B1"] = heading
    dash["B1"].font = Font(bold=True, size=18, color=DARK_BLUE)
    dash["B2"] = (f"Built {date.today():%d %b %Y} from the scheduled pulls in output/Data and Chart Outputs. "
                  "Every chart's data is on its own tab; full source data on the 'raw' tabs; units on 'Sources'. "
                  "Data comes straight from the publishing agency or grid operator; Ember only where no raw feed exists.")
    dash["B2"].font = Font(italic=True, color="6B6B6B")
    for j, h in enumerate(["Country", "Chart", "Latest", "Data tab", "Source"]):
        c = dash.cell(row=4, column=2 + j, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=DARK_BLUE)
    link = Font(color="2A78D6", underline="single")
    for i, (country, title, latest, tab, publisher, url) in enumerate(index_rows, start=5):
        for j, v in enumerate((country, title, latest, tab, publisher)):
            dash.cell(row=i, column=2 + j, value=v)
        dash.cell(row=i, column=5).hyperlink = f"#'{tab}'!A1"
        dash.cell(row=i, column=5).font = link
        if url:
            dash.cell(row=i, column=6).hyperlink = url
            dash.cell(row=i, column=6).font = link
    row = 5 + len(index_rows) + 1
    if missing:
        dash.cell(row=row, column=2, value="Not available / notes this run: " + "; ".join(missing)).font = Font(color="E34948")
        row += 1
    for col, w in (("B", 14), ("C", 52), ("D", 10), ("E", 38), ("F", 56)):
        dash.column_dimensions[col].width = w
    start = row + 1
    for k, (ch, (publisher, url)) in enumerate(charts):
        r = start + (k // 2) * ROWS_PER_CHART
        dash.add_chart(ch, f"{COLS[k % 2]}{r}")
        note = dash[f"{COLS[k % 2]}{r + ROWS_PER_CHART - 2}"]   # just below the chart
        note.value = f"Source: {publisher}"
        note.font = Font(italic=True, size=8, color="6B6B6B", underline="single" if url else None)
        if url:
            note.hyperlink = url


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(DATA_DIR, "Master Outputs", "south_and_central_america_master.xlsx"))
    ap.add_argument("--data-dir", default=DATA_DIR)
    args = ap.parse_args()

    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    dash2 = wb.create_sheet("Dashboard - Power & Hydro")
    used = {"Dashboard", "Dashboard - Power & Hydro", "Sources"}
    sources = []

    gas = collect(wb, DATASETS, args.data_dir, used, sources)
    ember_countries = set()
    for _, _, f in EMBER_FILES:
        try:
            ember_countries |= set(pd.ExcelFile(os.path.join(args.data_dir, f)).sheet_names)
        except Exception:
            pass
    raw_power, building = [], []
    for d in RAW_POWER_DATASETS:
        path = os.path.join(args.data_dir, d[2])
        if not os.path.exists(path):
            continue
        days = raw_history_days(path)
        if days >= MIN_RAW_DAYS or d[1] not in ember_countries:   # no Ember fallback: show the raw feed whatever its length
            raw_power.append(d)
        else:  # e.g. a source with no archive that only grows a day at a time: keep Ember until it has a year
            building.append(f"{d[1]} raw feed has {days} days so far - Ember shown until it has {MIN_RAW_DAYS}")
    have_raw = {d[1] for d in raw_power}
    print(f"power by type: raw operator data for {sorted(have_raw) or 'none'}; Ember for the rest")
    for b in building:
        print(b)
    parts = [collect(wb, raw_power, args.data_dir, used, sources),
             *(collect(wb, [(code, region, f, "*", "power")], args.data_dir, used, sources, skip=have_raw)
               for code, region, f in EMBER_FILES),
             collect(wb, HYDRO_DATASETS, args.data_dir, used, sources)]
    power = tuple(sum((p[i] for p in parts), []) for i in range(3))
    power[2].extend(building)
    # Combined South America generation by source, first on the Power & Hydro dashboard
    gen_frames = {}
    sa_total, sa_notes = south_america_generation(args.data_dir, have_raw, gen_frames)
    if not sa_total.empty:
        ws = wb.create_sheet(sheet_name("SA generation total data", used))
        sa_gw, gen_units = xlsx_charts.monthly_energy_to_gw(sa_total, "TWh per month", "South America power generation by source")
        df, n_bars = xlsx_charts.prepare(sa_gw)
        xlsx_charts.write_table(ws, df)
        ws.cell(row=1, column=df.shape[1] + 4, value="Countries summed (only months all of them have):")
        for i, note in enumerate(sa_notes, start=2):
            ws.cell(row=i, column=df.shape[1] + 4, value=note)
        src = ("Sum of the country series on this dashboard (grid operators; Ember where no raw feed yet)", None)
        power[0].insert(0, (xlsx_charts.build_chart(ws, df, n_bars, "South America power generation by source",
                                                    gen_units, "stacked_bar", width=CHART_W, height=CHART_H,
                                                    gridlines=False, inner=xlsx_charts.DASHBOARD_INNER), src))
        missing = [n.split(":", 1)[1].split(" in ")[0].strip() for n in sa_notes if n.startswith("NOT INCLUDED")]
        name = f"South America power generation by source ({len(SA_POWER_COUNTRIES)} countries" + (
            f"; {', '.join(missing)} missing in latest months)" if missing else ")")
        power[1].insert(0, ("South America", name,
                            df.index.max().strftime("%b/%y"), ws.title, *src))
        print("South America generation total:", "; ".join(sa_notes))
    # Capacity: combined South America chart (GW) right after the combined generation chart, then each country
    cap_total, cap_notes = south_america_capacity(args.data_dir)
    cap = collect(wb, [d for d in CAPACITY_DATASETS if os.path.exists(os.path.join(args.data_dir, d[2]))],
                  args.data_dir, used, sources)
    pos = 1 if not sa_total.empty else 0
    if not cap_total.empty:
        ws = wb.create_sheet(sheet_name("SA capacity total data", used))
        df, n_bars = xlsx_charts.prepare(cap_total)
        xlsx_charts.write_table(ws, df)
        ws.cell(row=1, column=df.shape[1] + 4, value="Countries summed:")
        for i, note in enumerate(cap_notes, start=2):
            ws.cell(row=i, column=df.shape[1] + 4, value=note)
        src = ("Sum of the country capacity workbooks (ANEEL, CAMMESA, CNE, XM, COES, CNDC, ADME/MIEM, ...)", None)
        power[0].insert(pos, (xlsx_charts.build_chart(ws, df, n_bars, "South America installed generation capacity",
                                                      "GW installed", "stacked_bar", width=CHART_W, height=CHART_H,
                                                      gridlines=False, inner=xlsx_charts.DASHBOARD_INNER), src))
        missing = [n.split(":", 1)[1].split("(")[0].strip() for n in cap_notes if n.startswith("NOT INCLUDED")]
        power[1].insert(pos, ("South America", "South America installed generation capacity" +
                              (f" (missing: {', '.join(missing)})" if missing else ""),
                              df.index.max().strftime("%b/%y"), ws.title, *src))
        pos += 1
        print("South America capacity total:", "; ".join(cap_notes))
    power[0][pos:pos] = cap[0]
    power[1][pos:pos] = cap[1]
    power[2].extend(cap[2])
    # Capacity factors: each country's dashboard generation over its capacity workbook. A country whose
    # SA-total series is a partial sheet (Paraguay: ANDE's own take) uses its full feed here.
    raw_files = {d[1]: d[2] for d in RAW_POWER_DATASETS}
    for country in SA_TOTAL_SHEET:
        if country in gen_frames and country in have_raw:
            path = os.path.join(args.data_dir, raw_files[country])
            build = add_charts.REGISTRY.get(raw_files[country]) or add_charts.power_daily(country)
            gen_frames[country] = build(path)[0]["df"]
    cf_chart, cf_row, cf_missing, cf_countries = capacity_factors.add_capacity_factor_sheets(
        wb, used, sheet_name, gen_frames,
        {c: os.path.join(args.data_dir, f) for _, c, f, _, _ in CAPACITY_DATASETS}, CHART_W, CHART_H,
        "Country generation (dashboard series) / installed capacity workbooks", "South America")
    if cf_chart:   # right after the capacity charts
        power[0].insert(pos + len(cap[0]), cf_chart)
        power[1].insert(pos + len(cap[1]), cf_row)
    power[2].extend(cf_missing)
    draw_dashboard(dash, "South & Central America energy - gas dashboard", *gas)
    draw_dashboard(dash2, "South & Central America energy - power generation & hydro", *power)
    cf_dash = wb.create_sheet("Dashboard - Capacity factors", 2)
    used.add("Dashboard - Capacity factors")
    cf_items = cf_countries   # regional chart first, then each country
    draw_dashboard(cf_dash, "South America - capacity factor by generation type (generation / capacity x hours)", [c for c, _ in cf_items], [r for _, r in cf_items], cf_missing)

    # long-term fundamentals: annual history (Ember yearly, EI Statistical Review, World Bank, IMF, degree days)
    fundamentals.add_long_term_dashboard(wb, used, sheet_name, "South & Central America and the Caribbean", args.data_dir, CHART_W, CHART_H,
                                         rows_per_chart=ROWS_PER_CHART, degree_days=("CDD_18",),
                                         index=sum(s.startswith("Dashboard") for s in wb.sheetnames))

    src = wb.create_sheet("Sources")
    src.append(["Country", "Dataset", "Workbook", "Publisher", "Link", "Units and notes (from the source workbook)"])
    for cell in src[1]:
        cell.font = Font(bold=True)
    for s in sources:
        src.append(list(s))
        if s[4]:
            src.cell(row=src.max_row, column=5).hyperlink = s[4]
    for col, w in (("A", 12), ("B", 14), ("C", 38), ("D", 40), ("E", 50), ("F", 160)):
        src.column_dimensions[col].width = w

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_charts.save_atomic(wb, args.out)
    print(f"Saved {args.out}: {len(gas[0])} gas charts, {len(power[0])} power/hydro charts; tabs {wb.sheetnames}")
    for label, m in (("gas", gas[2]), ("power/hydro", power[2])):
        if m:
            print(f"missing ({label}):", m)


if __name__ == "__main__":
    main()
