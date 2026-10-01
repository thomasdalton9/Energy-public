"""
Argentina national gas production and consumption by sector and by
distributor, from the Ministry of Economy's Series de Tiempo API
(apis.datos.gob.ar/series/api/series/) - free, public, no key - plus
gas exports by destination from ENARGAS's daily export reports.

Found via ARGENTINA_GAS_DISCOVERY.py + ARGENTINA_GAS_DISCOVERY2.py: all
these series belong to one dataset ("Producción y consumo de gas
natural", Secretaría de Energía / Ministerio de Economía), monthly
(R/P1M), 1996-01 to the latest available month. Units as published:
millones de metros cúbicos (million m3) per month.

Outputs (argentina_gas.xlsx):
  National              date, produccion_gas_natural, + 6 use-sector
                         columns (residencial, comercial, entes_oficiales,
                         industria, centrales_electricas [gas burned for
                         power generation], gnc [compressed gas for
                         vehicles]), country
  By sector (long)       the 6 use-sector columns reshaped long
  By distributor         date, one column per gas distribution company,
                         country
  By region (long)       the distributor columns reshaped long with a
                         region column added (REGION_MAP below)
  Exports by destination date (month), one column per destination country
                         (Chile, Brazil, Uruguay, ...), Total_exports - million
                         m3/month, complete months only
  Exports by point       the same months by export point (pipeline/border
                         crossing), million m3/month
  Exports daily          ENARGAS daily export reports as published, thousand
                         m3/day, one column per "country | point | route";
                         kept so later runs only fetch recent days

Exports come from ENARGAS's daily export reports ("Partes diarios de
exportacion"), found via discovery_archive/south_america/
ARGENTINA_GAS_EXPORTS_DISCOVERY*.py: two tables, "dentro del sistema"
(through the national transport system: GasAndes, NorAndino, Methanex
YPF/EGS to Chile; PetroUruguay and Cruz del Sur to Uruguay; TGM/Uruguayana
and "por Bolivia" to Brazil) and "fuera del sistema" (producers' own
pipelines: Gasoducto del Pacifico, Atacama, Methanex PAE/SIP/PTB to Chile;
"por Bolivia" to Brazil). Daily values in thousand m3 (9300 kcal/m3) are
summed to calendar months. There is no separate Bolivia column in the
reports: gas sent through Bolivia is labelled "Brasil por Bolivia" and is
counted under Brazil (its export point is listed in "Exports by point").
The Secretaria de Energia's comercio-exterior CSV and ENARGAS's monthly
Exportaciones.xlsx were checked too: the CSV has unit errors in several
months and the xlsx covers only part of the in-system points.

REGION_MAP covers the 9 classic distribution licensees from Argentina's
1992 gas-distribution privatization, each with a single well-defined
service territory. "sdb" and "redengas" are smaller/newer entities this
script couldn't confidently place in one province - left region=None
(printed as unmapped) rather than guessed. "gnc" (compressed gas for
vehicles) is a use-category, not a regional distributor, so it's not
in the distributor table at all.

Usage: python3 ARGENTINA_GAS.py [--out argentina_gas.xlsx]
"""
print("STARTING", flush=True)

import argparse
import html
import os
import re
import sys

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

HEADERS = {"User-Agent": "gas-demand-scripts/1.0"}
TIMEOUT = (10, 60)
SERIES_API = "https://apis.datos.gob.ar/series/api/series/"
COUNTRY = "Argentina"
OUT_DEFAULT = "argentina_gas.xlsx"

# column name -> confirmed monthly (R/P1M) series id, from
# ARGENTINA_GAS_DISCOVERY2.py's live output
PRODUCTION_ID = "364.3_PRODUCCIoNRAL__25"

SECTOR_IDS = {
    "residencial": "364.3_RESIDENCIAIAL__11",
    "comercial": "364.3_COMERCIALIAL__9",
    "entes_oficiales": "364.3_ENTES_OFICLES__15",
    "industria": "364.3_INDUSTRIARIA__9",
    "centrales_electricas": "364.3_CENTRALES_CAS__20",
    "gnc": "364.3_GNCGNC__3",
}

DISTRIBUTOR_IDS = {
    "metrogas": "364.3_METROGASGAS__8",
    "gas_natural_fenosa": "364.3_GAS_NATURAOSA__18",
    "distrib_gas_del_centro_ecogas": "364.3_DISTRIB._GGAS__30",
    "distrib_gas_cuyana_ecogas": "364.3_DISTRIB._GGAS__26",
    "litoral_gas": "364.3_LITORAL_GAGAS__11",
    "gasnea": "364.3_GASNEANEA__6",
    "gasnor": "364.3_GASNORNOR__6",
    "camuzzi_gas_pampeana": "364.3_CAMUZZI_GAANA__20",
    "camuzzi_gas_del_sur": "364.3_CAMUZZI_GASUR__19",
    "sdb": "364.3_SDBSDB__3",
    "redengas": "364.3_REDENGASGAS__8",
}

# The 9 classic 1992-privatization distribution licensees, each with one
# well-defined service territory. sdb/redengas deliberately left out -
# not confidently placeable in a single province.
REGION_MAP = {
    "metrogas": "CABA + Norte GBA",
    "gas_natural_fenosa": "Oeste/Sur GBA",
    "camuzzi_gas_pampeana": "Buenos Aires (provincia) + La Pampa",
    "camuzzi_gas_del_sur": "Patagonia",
    "litoral_gas": "Santa Fe + Entre Ríos",
    "distrib_gas_del_centro_ecogas": "Córdoba",
    "distrib_gas_cuyana_ecogas": "Cuyo (Mendoza/San Juan/San Luis)",
    "gasnor": "NOA",
    "gasnea": "NEA",
}


DATA_START = "2021-01-01"

# ENARGAS daily export reports (see module docstring). POST, at most 365 days per request.
ENARGAS_PAGE = "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/dod-partes-exp-imp-consulta.php"
ENARGAS_LIST = "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/partes-diarios-exp-imp-consulta-listado.php"
EXPORT_TABLES = {"exp_dentro": "transport system", "exp_fuera": "producer pipeline"}
COUNTRY_EN = {"Chile": "Chile", "Brasil": "Brazil", "Uruguay": "Uruguay", "Bolivia": "Bolivia",
              "Paraguay": "Paraguay"}
COUNTRY_ORDER = ("Chile", "Brazil", "Uruguay", "Bolivia", "Paraguay")
EXPORTS_DAILY_SHEET = "Exports daily"
REFETCH_DAYS = 45        # each run re-reads the last 45 days (late or revised daily reports); older days are kept
CHUNK_DAYS = 180


def _text(cell_html):
    return html.unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", cell_html))).strip()


def parse_export_table(page, route):
    """ENARGAS report table -> daily frame (thousand m3/day), columns 'Country | Point | route'.
    Headers read '<b>Chile</b><br>GasAndes'; the published Total column is checked, then dropped."""
    heads = [_text(re.sub(r"<br\s*/?>", "|", h)) for h in re.findall(r"<th[^>]*>(.*?)</th>", page, re.S | re.I)]
    rows = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", page, re.S | re.I):
        cells = [_text(c) for c in re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S | re.I)]
        if cells and re.fullmatch(r"\d{2}/\d{2}/\d{4}", cells[0]):
            rows.append(cells)
    if not rows:
        return pd.DataFrame()
    if len(heads) != len(rows[0]) or heads[0].lower() != "fecha":
        raise ValueError(f"unexpected ENARGAS export table layout: {heads} vs {len(rows[0])} cells")
    cols = []
    for h in heads[1:]:
        parts = [x.strip() for x in h.split("|")]
        cols.append(f"{COUNTRY_EN.get(parts[0], parts[0])} | {parts[1]} | {route}" if len(parts) == 2 else h)
    df = pd.DataFrame([r[1:] for r in rows], columns=cols,
                      index=pd.to_datetime([r[0] for r in rows], format="%d/%m/%Y"))
    df = df.apply(lambda c: pd.to_numeric(c.str.replace(",", ".", regex=False), errors="coerce"))
    total = df.pop("Total") if "Total" in df else None
    if total is not None:
        gap = (df.fillna(0).sum(axis=1) - total.fillna(0)).abs()
        if (gap > 1).any():
            print(f"  WARNING {route}: {int((gap > 1).sum())} days where the points don't add up to the "
                  f"published Total (max gap {gap.max():.0f} thousand m3)", flush=True)
    df.index.name = "date"
    return df


def fetch_exports_daily(start, end):
    """Both ENARGAS export tables, start..end (Timestamps), in requests of at most CHUNK_DAYS days."""
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                                    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"})
    s.get(ENARGAS_PAGE, params={"tipo": "exp_dentro"}, timeout=TIMEOUT)     # session cookie, as a browser gets
    frames = []
    for tipo, route in EXPORT_TABLES.items():
        parts = []
        d0 = start
        while d0 <= end:
            d1 = min(d0 + pd.Timedelta(days=CHUNK_DAYS - 1), end)
            r = s.post(ENARGAS_LIST, timeout=TIMEOUT, data={"fecha_desde": d0.strftime("%Y-%m-%d"),
                                                             "fecha_hasta": d1.strftime("%Y-%m-%d"),
                                                             "tipo_list": tipo})
            r.raise_for_status()
            part = parse_export_table(r.text, route)
            print(f"  {tipo} {d0.date()}..{d1.date()}: {len(part)} days", flush=True)
            if not part.empty:
                parts.append(part)
            d0 = d1 + pd.Timedelta(days=1)
        if parts:
            frames.append(pd.concat(parts))
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, axis=1).sort_index()
    return out[~out.index.duplicated(keep="last")]


def load_exports_daily(path):
    try:
        old = pd.read_excel(path, sheet_name=EXPORTS_DAILY_SHEET, index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    old.index = pd.to_datetime(old.index)
    old.index.name = "date"
    return old


def update_exports_daily(path, start_date):
    """Incremental: keep the archived daily table and fetch only from REFETCH_DAYS before its last day."""
    old = load_exports_daily(path)
    start = pd.Timestamp(start_date)
    end = pd.Timestamp.today().normalize()
    incremental = not old.empty and old.index.min() <= start
    if incremental:
        start = old.index.max() - pd.Timedelta(days=REFETCH_DAYS)
    print(f"Fetching ENARGAS daily export reports {start.date()}..{end.date()} "
          f"({'incremental' if incremental else 'full backfill'})...", flush=True)
    new = fetch_exports_daily(start, end)
    if not incremental or new.empty:
        daily = new if not incremental else old
    else:
        daily = pd.concat([old[old.index < new.index.min()], new]).sort_index()
        daily = daily[~daily.index.duplicated(keep="last")]
    return daily[daily.index >= pd.Timestamp(start_date)] if not daily.empty else daily


def monthly_exports(daily):
    """Daily thousand m3 -> calendar-month million m3, complete months only (every day reported)."""
    if daily.empty:
        return pd.DataFrame(), pd.DataFrame()
    month = daily.index.to_period("M")
    days = pd.Series(1, index=daily.index).groupby(month).sum()
    complete = days[days.values == days.index.days_in_month].index
    by_point = daily.fillna(0).groupby(month).sum() / 1000.0
    by_point = by_point.loc[by_point.index.isin(complete)]
    by_point = by_point.loc[:, by_point.ne(0).any()]           # points with no flow since the start date
    by_point.index = by_point.index.to_timestamp()
    by_point.index.name = "date"
    country = by_point.T.groupby(lambda c: c.split(" | ")[0]).sum().T
    country = country[[c for c in COUNTRY_ORDER if c in country] + [c for c in country if c not in COUNTRY_ORDER]]
    country["Total_exports"] = country.sum(axis=1)
    return country.round(3), by_point.round(3)


def fetch_series(ids, start_date):
    """ids: {column_name: series_id}. Returns a wide DataFrame indexed by
    date, one column per name, values in the series' native units
    (million m3/month)."""
    id_list = list(ids.values())
    r = requests.get(SERIES_API, headers=HEADERS, timeout=TIMEOUT,
                      params={"ids": ",".join(id_list), "format": "json", "limit": 5000,
                              "start_date": start_date})
    r.raise_for_status()
    payload = r.json()
    dates = [row[0] for row in payload["data"]]
    df = pd.DataFrame(payload["data"], columns=["date"] + id_list).set_index("date")
    df.index = pd.to_datetime(df.index)
    id_to_name = {v: k for k, v in ids.items()}
    return df.rename(columns=id_to_name).sort_index()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default=OUT_DEFAULT)
    parser.add_argument("--start-date", default=DATA_START, help="first month to pull (YYYY-MM-DD)")
    args = parser.parse_args()

    print(f"Fetching production + sector consumption from {args.start_date}...", flush=True)
    national = fetch_series({"produccion_gas_natural": PRODUCTION_ID, **SECTOR_IDS}, args.start_date)
    print(f"  {len(national)} months, {national.index.min().date()} to {national.index.max().date()}", flush=True)

    print("Fetching distributor consumption...", flush=True)
    distributors = fetch_series(DISTRIBUTOR_IDS, args.start_date)
    print(f"  {len(distributors)} months, {distributors.index.min().date()} to {distributors.index.max().date()}",
          flush=True)

    try:
        exports_daily = update_exports_daily(args.out, args.start_date)
    except (requests.RequestException, ValueError) as e:      # keep the archived exports if ENARGAS is down
        print(f"  ENARGAS export reports failed ({type(e).__name__}: {e}); keeping the archived daily table",
              flush=True)
        exports_daily = load_exports_daily(args.out)
    exports, exports_points = monthly_exports(exports_daily)
    if not exports.empty:
        print(f"  exports: {len(exports)} complete months, {exports.index.min().date()} to "
              f"{exports.index.max().date()}; daily reports to {exports_daily.index.max().date()}", flush=True)
        print(exports.tail(3).to_string(), flush=True)

    national_out = national.copy()
    national_out.insert(0, "country", COUNTRY)

    sector_long = national[list(SECTOR_IDS)].reset_index().melt(
        id_vars="date", var_name="sector", value_name="mmc")
    sector_long["country"] = COUNTRY
    sector_long = sector_long.dropna(subset=["mmc"])

    distributors_out = distributors.copy()
    distributors_out.insert(0, "country", COUNTRY)

    region_long = distributors.reset_index().melt(id_vars="date", var_name="distributor", value_name="mmc")
    region_long["region"] = region_long["distributor"].map(REGION_MAP)
    region_long["country"] = COUNTRY
    region_long = region_long.dropna(subset=["mmc"])
    unmapped = sorted(set(region_long.loc[region_long["region"].isna(), "distributor"]))
    if unmapped:
        print(f"  distributors: {len(unmapped)} unmapped (region=None): {unmapped}", flush=True)

    notes = [
        "UNITS",
        "All values in millones de metros cubicos (million m3) per month, as published.",
        "",
        "SECTORS",
        "'National': produccion_gas_natural (national gas production) plus 6 use-sector consumption columns - "
        "residencial, comercial, entes_oficiales (official/government bodies), industria, centrales_electricas "
        "(gas burned for power generation), gnc (compressed gas for vehicles). 'By sector (long)' reshapes the "
        "6 sector columns long.",
        "",
        "REGIONS",
        "'By distributor': consumption by each of Argentina's licensed gas distribution companies. 'By region "
        "(long)' adds a region column via REGION_MAP - the 9 classic 1992-privatization distribution licensees, "
        "each with one well-defined service territory. 'sdb' and 'redengas' are smaller/newer entities not "
        "confidently placeable in a single province - left region=None (printed as unmapped) rather than guessed.",
        "",
        "EXPORTS",
        "'Exports by destination': natural gas exports by destination country, million m3 per month (m3 of "
        "9300 kcal/m3 gas, ENARGAS's standard), one column per country in the reports (Chile, Brazil, Uruguay) "
        "plus Total_exports. Calendar months, complete months only (a month appears once ENARGAS has published "
        "every day of it). From 2021-01.",
        "'Exports by point': the same months by export point and route. 'transport system' = through the "
        "national transport system (ENARGAS 'dentro del sistema': GasAndes, NorAndino, Methanex YPF and "
        "Methanex EGS to Chile; PetroUruguay and Cruz del Sur to Uruguay; TGM/Uruguayana and 'por Bolivia' to "
        "Brazil). 'producer pipeline' = producers' own export pipelines (ENARGAS 'fuera del sistema': "
        "Gasoducto del Pacifico, Atacama and Methanex PAE/SIP/PTB to Chile; 'por Bolivia' to Brazil).",
        "Bolivia: the reports have no column for sales to Bolivia. Gas sent through Bolivia's network is "
        "reported as 'Brasil por Bolivia' and counted under Brazil (its destination); see 'Exports by point'.",
        "'Exports daily': the daily reports as published, thousand m3 per day, one column per "
        "'country | point | route'; the published Total column is checked against the points, not stored. "
        f"Each run keeps the archived days and re-reads only the last {REFETCH_DAYS} days.",
        "Validation: ENARGAS's monthly file 'Gas exportado a traves del sistema de transporte' matches these "
        "sums for GasAndes and PetroUruguay in 2021-2025 but lists only part of the in-system points; in 2026 "
        "its GasAndes figures run below the daily reports (consolidated vs daily data).",
        "",
        "SOURCE",
        "apis.datos.gob.ar/series/api/series - Secretaria de Energia, Ministerio de Economia, dataset "
        f"'Produccion y consumo de gas natural'. Monthly; this archive starts {args.start_date} (the source "
        "goes back to 1996-01 - rerun with --start-date to pull more).",
        "Exports: ENARGAS (Ente Nacional Regulador del Gas y la Electricidad), Partes diarios de exportacion, "
        "dentro del sistema and fuera del sistema: " + ENARGAS_PAGE + "?tipo=exp_dentro and ?tipo=exp_fuera",
    ]
    sheets = {
        "National": national_out,
        "By sector (long)": sector_long,
        "By distributor": distributors_out,
        "By region (long)": region_long,
    }
    if not exports.empty:
        sheets["Exports by destination"] = exports
        sheets["Exports by point"] = exports_points
        sheets[EXPORTS_DAILY_SHEET] = exports_daily
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "SECTORS", "REGIONS", "EXPORTS", "SOURCE"})
    print(f"Saved {args.out}", flush=True)
    print(national_out.tail().to_string(), flush=True)


if __name__ == "__main__":
    main()
