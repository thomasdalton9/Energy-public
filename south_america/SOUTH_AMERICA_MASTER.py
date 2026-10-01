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

Usage: python3 SOUTH_AMERICA_MASTER.py [--out "output/Data and Chart Outputs/south_and_central_america_master.xlsx"]
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
import water_year_chart  # noqa: E402
import xlsx_charts  # noqa: E402

DATA_DIR = os.path.join(ROOT, "output", "Data and Chart Outputs")
DARK_BLUE = "17365D"

# (country code, country, workbook, raw sheet to copy, short dataset name)
DATASETS = [
    ("AR", "Argentina", "argentina_gas_monthly.xlsx", "National", "gas"),
    ("BR", "Brazil", "brazil_gas_monthly.xlsx", "Demand by segment", "gas"),
    ("BO", "Bolivia", "bolivia_gas_demand_by_sector.xlsx", "Demand by sector", "gas"),
    ("CL", "Chile", "chile_gas_imports.xlsx", "Gas imports", "gas imports"),
    ("CO", "Colombia", "colombia_gas_demand_by_sector.xlsx", "Demand by sector", "gas"),
    ("EC", "Ecuador", "ecuador_gas.xlsx", "Gas by use", "gas"),
    ("UY", "Uruguay", "uruguay_gas_demand_by_sector.xlsx", "Demand by sector", "gas"),
    ("TT", "Trinidad & Tobago", "trinidad_gas.xlsx", "Utilization by sector", "gas"),
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
    ("CO", "Colombia", "colombia_power_generation_daily.xlsx", "Daily", "power"),
    ("EC", "Ecuador", "ecuador_power_generation_daily.xlsx", "Daily", "power"),
    ("PE", "Peru", "peru_power_generation_daily.xlsx", "Daily", "power"),
    ("UY", "Uruguay", "uruguay_power_generation_daily.xlsx", "Daily", "power"),
    # Central America
    ("BZ", "Belize", "belize_power_generation_daily.xlsx", "Daily", "power"),
    ("CR", "Costa Rica", "costa_rica_power_generation_daily.xlsx", "Daily", "power"),
    ("SV", "El Salvador", "el_salvador_power_generation_daily.xlsx", "Daily", "power"),
    ("GT", "Guatemala", "guatemala_power_generation_daily.xlsx", "Daily", "power"),
    ("HN", "Honduras", "honduras_power_generation_daily.xlsx", "Daily", "power"),
    ("NI", "Nicaragua", "nicaragua_power_generation_daily.xlsx", "Daily", "power"),
    ("PA", "Panama", "panama_power_generation_daily.xlsx", "Daily", "power"),
]
HYDRO_DATASETS = [
    ("BR", "Brazil", "brazil_hydro_reservoirs.xlsx", "Daily", "hydro"),
    ("CO", "Colombia", "colombia_hydro_reservoirs.xlsx", "Daily", "hydro"),
    ("PA", "Panama", "gatun_lake_level.xlsx", "Daily", "hydro"),
]


def gatun(path):
    """gatun_lake_level.xlsx draws its own chart (REGISTRY entry is None); the master needs a spec."""
    d = add_charts.by_date(add_charts.read(path, "Daily"), "date")
    return [{"name": "Gatun", "water_year": d["level_ft"], "title": "Panama Canal, Gatun Lake level",
             "units": "feet above sea level"}]


# Chart specs for workbooks whose add_charts REGISTRY entry is None (they chart themselves)
MASTER_SPECS = {"gatun_lake_level.xlsx": gatun}

# Where each workbook's data comes from: (publisher, link). Shown on the dashboards.
SOURCES = {
    "argentina_gas_monthly.xlsx": ("Secretaría de Energía (Argentina), Series de Tiempo API",
                                   "https://datos.gob.ar/series/api/series/"),
    "brazil_gas_monthly.xlsx": ("MME Brazil gas bulletin (segments, to Jun-2025); ANP open data (pipeline flows, production, imports)",
                                "https://www.gov.br/anp/pt-br/centrais-de-conteudo/dados-abertos/dados-consolidados-movimentacao-de-gas-natural-em-gasodutos-de-transporte"),
    "bolivia_gas_demand_by_sector.xlsx": ("INE Bolivia",
                                          "https://www.ine.gob.bo/index.php/estadisticas-economicas/hidrocarburos-mineria/hidrocarburo-cuadros-estadisticos/"),
    "chile_gas_imports.xlsx": ("CNE Chile, Reporte Mensual (customs data)", "https://www.cne.cl/nuestros-servicios/reportes/"),
    "colombia_gas_demand_by_sector.xlsx": ("Gestor del Mercado de Gas (BMC)", "https://www.bmcbec.com.co/informes/informes-mensuales"),
    "ecuador_gas.xlsx": ("EP Petroecuador", "https://www.eppetroecuador.ec/?p=3721"),
    "trinidad_gas.xlsx": ("Ministry of Energy and Energy Industries (MEEI), monthly bulletins",
                          "https://www.energy.gov.tt/category/publications/energy-industry-bulletins/"),
    "uruguay_gas_demand_by_sector.xlsx": ("MIEM Uruguay, VisualPEB energy balance", "https://visualpeb.miem.gub.uy/visualPEB/gas_natural"),
    "argentina_power_generation_daily.xlsx": ("CAMMESA", "https://cammesaweb.cammesa.com/"),
    "bolivia_power_generation_daily.xlsx": ("CNDC Bolivia", "https://www.cndc.bo/"),
    "brazil_power_generation_daily.xlsx": ("ONS Brazil open data", "https://dados.ons.org.br/"),
    "colombia_power_generation_daily.xlsx": ("XM Colombia", "https://www.xm.com.co/"),
    "ecuador_power_generation_daily.xlsx": ("CENACE Ecuador", "https://www.cenace.gob.ec/info-operativa/InformacionOperativa.htm"),
    "peru_power_generation_daily.xlsx": ("COES Peru", "https://www.coes.org.pe/"),
    "uruguay_power_generation_daily.xlsx": ("ADME Uruguay", "https://pronos.adme.com.uy/"),
    "brazil_hydro_reservoirs.xlsx": ("ONS Brazil open data (EAR)", "https://dados.ons.org.br/dataset/ear-diario-por-subsistema"),
    "colombia_hydro_reservoirs.xlsx": ("XM Colombia", "https://www.xm.com.co/"),
    "belize_power_generation_daily.xlsx": ("Belize Electricity Ltd (BEL)", "https://www.bel.com.bz/"),
    "costa_rica_power_generation_daily.xlsx": ("ICE / CENCE Costa Rica", "https://apps.grupoice.com/CenceWeb/"),
    "el_salvador_power_generation_daily.xlsx": ("Unidad de Transacciones (UT) El Salvador", "https://www.ut.com.sv/"),
    "guatemala_power_generation_daily.xlsx": ("Administrador del Mercado Mayorista (AMM) Guatemala",
                                              "https://www.amm.org.gt/"),
    "honduras_power_generation_daily.xlsx": ("Operador del Sistema (ODS) Honduras", "https://www.ods.org.hn/"),
    "nicaragua_power_generation_daily.xlsx": ("CNDC / ENATREL Nicaragua", "https://www.cndc.org.ni/"),
    "panama_power_generation_daily.xlsx": ("CND / ETESA Panama", "https://www.cnd.com.pa/"),
    "gatun_lake_level.xlsx": ("Panama Canal Authority (ACP)", "https://evtms-rpts.pancanal.com/"),
    **{f: ("Ember monthly electricity data", "https://ember-energy.org/data/monthly-electricity-data/") for f in EMBER},
}
# The grid operator Ember compiles each country from (named on Ember-fed charts)
OPERATORS = {"Argentina": "CAMMESA", "Bolivia": "CNDC", "Brazil": "ONS", "Chile": "Coordinador Eléctrico Nacional",
             "Colombia": "XM", "Ecuador": "CENACE", "Peru": "COES", "Uruguay": "ADME",
             "Belize": "BEL", "Costa Rica": "ICE/CENCE", "El Salvador": "UT", "Guatemala": "AMM", "Honduras": "ODS",
             "Nicaragua": "CNDC", "Panama": "CND"}


MIN_RAW_DAYS = 365   # a raw generation workbook replaces Ember only once it has a year of history


def raw_history_days(path):
    try:
        return int(pd.read_excel(path, sheet_name="Daily", usecols=[0]).iloc[:, 0].nunique())
    except Exception:
        return 0


def source_of(fname, spec_name=None):
    publisher, url = SOURCES.get(fname, (fname, None))
    if fname in EMBER and spec_name in OPERATORS:
        publisher = f"Ember, compiled from {OPERATORS[spec_name]} (no raw feed yet)"
    return publisher, url

CHART_W, CHART_H = 17.0, 8.5       # cm
ROWS_PER_CHART = 19
COLS = ("B", "F")                  # two charts per row (B-E are ~17 cm wide, so F starts the second)


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


def collect(wb, datasets, data_dir, used, sources, skip=()):
    """Data + raw tabs for one dashboard's datasets; returns (charts, index rows, missing).
    Each chart is (chart, (publisher, url)); index rows end with the same source.
    skip: spec names not to chart (Ember countries that have a raw workbook)."""
    charts, index_rows, missing = [], [], []
    for code, country, fname, raw_sheet, short in datasets:
        path = os.path.join(data_dir, fname)
        if not os.path.exists(path):
            missing.append(f"{country} {short} ({fname})")
            continue
        try:
            build = MASTER_SPECS.get(fname) or add_charts.REGISTRY.get(fname) or (
                add_charts.power_daily(f"{country} power generation by type")
                if fname.endswith("_power_generation_daily.xlsx") else add_charts.generic)
            specs = build(path)
        except Exception as e:
            missing.append(f"{country} {short} ({fname}: {type(e).__name__}: {e})")
            continue
        for s in specs:
            if s["name"] in skip:
                continue
            src = source_of(fname, s["name"])
            if "water_year" in s:
                table, meta = water_year_chart.water_year_table(s["water_year"])
                ws = wb.create_sheet(sheet_name(f"{code} {s['name']} data", used))
                water_year_chart.write_table(ws, table)
                charts.append((water_year_chart.build_chart(ws, table, meta, s["title"], s["units"],
                                                            width=CHART_W, height=CHART_H), src))
                index_rows.append((country, f"{s['title']} (water year)", pd.Timestamp(meta["last"]).strftime("%d/%m/%y"),
                                   ws.title, *src))
                continue
            df, n_bars = xlsx_charts.prepare(s["df"].dropna(how="all"), s.get("line_cols", ()))
            if df.empty:
                continue
            label = f"{s['name']} {short} data" if raw_sheet == "*" else f"{code} {s['name']} data"
            ws = wb.create_sheet(sheet_name(label, used))
            xlsx_charts.write_table(ws, df, s["date_format"])
            charts.append((xlsx_charts.build_chart(ws, df, n_bars, s["title"], s["units"], s["kind"],
                                                   s["date_format"], width=CHART_W, height=CHART_H), src))
            index_rows.append((s["name"] if raw_sheet == "*" else country, s["title"], df.index.max().strftime("%b/%y"),
                               ws.title, *src))
        raw_sheets = ([n for n in pd.ExcelFile(path).sheet_names
                       if n.lower() not in ("units", "notes") and not n.startswith("Chart")
                       and n != water_year_chart.SHEET and n not in skip] if raw_sheet == "*" else [raw_sheet])
        for rs in raw_sheets:
            try:
                raw = add_charts.read(path, rs)
                label = f"{rs} {short} raw" if raw_sheet == "*" else f"{code} {short} raw"
                write_frame(wb.create_sheet(sheet_name(label, used)), raw)
            except Exception as e:
                missing.append(f"{country} {short} raw sheet {rs} ({e})")
        sources.append((country, short, fname, *source_of(fname), notes_text(path)))
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
    for col, w in (("B", 12), ("C", 48), ("D", 9), ("E", 24), ("F", 52)):
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
    ap.add_argument("--out", default=os.path.join(DATA_DIR, "south_and_central_america_master.xlsx"))
    ap.add_argument("--data-dir", default=DATA_DIR)
    args = ap.parse_args()

    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    dash2 = wb.create_sheet("Dashboard - Power & Hydro")
    used = {"Dashboard", "Dashboard - Power & Hydro", "Sources"}
    sources = []

    gas = collect(wb, DATASETS, args.data_dir, used, sources)
    raw_power, building = [], []
    for d in RAW_POWER_DATASETS:
        path = os.path.join(args.data_dir, d[2])
        if not os.path.exists(path):
            continue
        days = raw_history_days(path)
        if days >= MIN_RAW_DAYS:
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
    draw_dashboard(dash, "South & Central America energy - gas dashboard", *gas)
    draw_dashboard(dash2, "South & Central America energy - power generation & hydro", *power)

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
