"""
Argentina national gas production and consumption by sector and by
distributor, from the Ministry of Economy's Series de Tiempo API
(apis.datos.gob.ar/series/api/series/) - free, public, no key - plus
gas exports by destination from ENARGAS's daily export reports, and
measured NET SUPPLY: gas injected into the transport and distribution
systems from domestic basins plus imports (LNG and pipeline).

Found via ARGENTINA_GAS_DISCOVERY.py + ARGENTINA_GAS_DISCOVERY2.py: all
the Series API series belong to one dataset ("Producción y consumo de gas
natural", Secretaría de Energía / Ministerio de Economía), monthly
(R/P1M), 1996-01 to the latest available month. Units as published:
millones de metros cúbicos (million m3) per month. The six sector series
are ENARGAS's "Gas entregado por tipo de usuario" (GETD.xlsx) summed over
distribution, transport and off-system users.

Outputs (argentina_gas.xlsx):
  National              date, produccion_gas_natural (gross production), + 6
                         use-sector columns (residencial, comercial,
                         entes_oficiales, industria, centrales_electricas
                         [gas burned for power generation], gnc [compressed
                         gas for vehicles]), country
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
  Supply net             monthly net supply, million m3/month (see below)
  Supply daily           daily injection (by basin) and imports, mcm/d
  Balance check          net supply + imports vs deliveries + exports, by month
  Imports daily          ENARGAS daily import reports as published, thousand m3/day
  Transport daily        ENARGAS daily transport reports as published (parsed), mcm/d
  Deliveries (ENARGAS)   ENARGAS gas delivered by user type (GETD), million m3/month

NET SUPPLY (found via discovery_archive/south_america/ARGENTINA_GAS_NETSUPPLY_DISCOVERY*.py):
  * Domestic injection, monthly: ENARGAS "Gas Recibido por Transportistas y
    Distribuidoras de Productores y Otros Origenes - por Cuenca" (GRT.xlsx): gas metered
    into the TGN and TGS transport systems and the distributors' own pipelines, by basin,
    thousand m3 of 9300 kcal. Gas reaches these receipt points after field use,
    reinjection, venting/flaring and gas-plant processing, so it is net of all of them.
    The "Otros Origenes" (LNG) columns are left out and imports are taken from the import
    reports instead; pipeline imports that GRT meters inside a basin column are taken out
    of it (Bolivia and NorAndino enter TGN's Norte pipeline -> GRT 'Noroeste';
    GasAndes enters TGN's Centro Oeste -> GRT TGN 'Neuquina'), as the daily transport
    reports' footnotes say.
  * Direct deliveries outside the transport system, which the sector series and the
    export series include but GRT does not: power plants supplied at the wellhead
    (GETD "Off System - Centrales electricas") and exports through producers' own
    pipelines (ENARGAS export reports "fuera del sistema").
  * Imports: ENARGAS daily import reports ("Partes de Importacion", tipo_list
    importaciones): Bolivia, GNL Bahia Blanca, GNL Escobar, Gasandes and Norandino
    (imports from Chile), thousand m3/day, summed to complete months.
  * Daily: ENARGAS "Parte Diario Operativo - Gas Natural Transportado" (one PDF per day,
    descarga.php?tipo=transporte), provisional data from TGN/TGS: injection by basin
    (Norte, Neuquina, Austral) and total, mcm/d of 9300 kcal, with footnotes giving the
    imports inside those figures (Bolivia+Norandino in Norte, GNL Escobar+Gasandes in
    Centro Oeste, GNL Bahia Blanca in Neuba II until it was replaced by the Perito
    Moreno pipeline footnote) and the line-pack. Domestic daily injection = total minus
    those imports. Each run re-reads the last 45 days and backfills missing days within
    a time budget.
  * Transport fuel, unaccounted gas and losses: ENARGAS publishes them only as % of gas
    delivered per transporter (CLP.xlsx); they are stored as published (%), not turned
    into volumes, so the charts carry no "system use & losses" segment.

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
import io
import os
import re
import sys
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

HEADERS = {"User-Agent": "gas-demand-scripts/1.0"}
BROWSER_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
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

# ENARGAS daily export/import reports (see module docstring). POST, at most 365 days per request.
ENARGAS = "https://www.enargas.gob.ar/secciones/transporte-y-distribucion/"
ENARGAS_PAGE = ENARGAS + "dod-partes-exp-imp-consulta.php"
ENARGAS_LIST = ENARGAS + "partes-diarios-exp-imp-consulta-listado.php"
EXPORT_TABLES = {"exp_dentro": "transport system", "exp_fuera": "producer pipeline"}
IMPORT_TABLES = {"importaciones": None}
COUNTRY_EN = {"Chile": "Chile", "Brasil": "Brazil", "Uruguay": "Uruguay", "Bolivia": "Bolivia",
              "Paraguay": "Paraguay"}
COUNTRY_ORDER = ("Chile", "Brazil", "Uruguay", "Bolivia", "Paraguay")
EXPORTS_DAILY_SHEET = "Exports daily"
IMPORTS_DAILY_SHEET = "Imports daily"
TRANSPORT_DAILY_SHEET = "Transport daily"
REFETCH_DAYS = 45        # each run re-reads the last 45 days (late or revised daily reports); older days are kept
CHUNK_DAYS = 180

# ENARGAS monthly operating data (Datos operativos de transporte y distribucion), whole files, small
DATOS = ENARGAS + "datos-estadisticos/"
GRT_URL = DATOS + "GRT/GRT.xlsx"        # gas received from producers and other origins, by basin
GETD_URL = DATOS + "GETD/GETD.xlsx"     # gas delivered, total system, by user type
CLP_URL = DATOS + "CPGNNC/CLP.xlsx"     # transport fuel, unaccounted gas and losses, % of gas delivered
# ENARGAS daily transport report PDFs (Parte Diario Operativo - Gas Natural Transportado)
PARTE_URL = ENARGAS + "descarga.php"
PARTES_BUDGET_S = int(os.environ.get("PARTES_BUDGET_S", "900"))   # seconds per run for daily-PDF backfill

IMPORT_COLS = ["LNG_Escobar", "LNG_BahiaBlanca", "Imports_Bolivia", "Imports_Chile_GasAndes", "Imports_Chile_NorAndino",
               "Imports_Chile", "Total_imports"]
# ENARGAS import report columns -> English names (thousand m3/day as published)
IMPORT_NAMES = {"Bolivia": "Bolivia", "GNL B. Blanca": "LNG Bahia Blanca", "GNL Escobar": "LNG Escobar",
                "Gasandes": "Chile GasAndes", "Norandino": "Chile NorAndino"}


def _text(cell_html):
    return html.unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", cell_html))).strip()


def parse_report_table(page, route):
    """ENARGAS export/import report table -> daily frame (thousand m3/day).
    Export headers read '<b>Chile</b><br>GasAndes' -> 'Chile | GasAndes | route'; import headers are single
    names ('Bolivia', 'GNL Escobar') and are kept as they are. The published Total column is checked, then
    dropped. Days not yet reported (all cells blank) are dropped."""
    heads = [_text(re.sub(r"<br\s*/?>", "|", h)) for h in re.findall(r"<th[^>]*>(.*?)</th>", page, re.S | re.I)]
    rows = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", page, re.S | re.I):
        cells = [_text(c) for c in re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S | re.I)]
        if cells and re.fullmatch(r"\d{2}/\d{2}/\d{4}", cells[0]):
            rows.append(cells)
    if not rows:
        return pd.DataFrame()
    if len(heads) != len(rows[0]) or heads[0].lower() != "fecha":
        raise ValueError(f"unexpected ENARGAS report table layout: {heads} vs {len(rows[0])} cells")
    cols = []
    for h in heads[1:]:
        parts = [x.strip() for x in h.split("|")]
        cols.append(f"{COUNTRY_EN.get(parts[0], parts[0])} | {parts[1]} | {route}" if len(parts) == 2 and route
                    else " ".join(parts))
    df = pd.DataFrame([r[1:] for r in rows], columns=cols,
                      index=pd.to_datetime([r[0] for r in rows], format="%d/%m/%Y"))
    df = df.apply(lambda c: pd.to_numeric(c.str.replace(",", ".", regex=False), errors="coerce"))
    df = df.dropna(how="all")
    total = df.pop("Total") if "Total" in df else None
    if total is not None:
        gap = (df.fillna(0).sum(axis=1) - total.fillna(0)).abs()
        if (gap > 1).any():
            print(f"  WARNING {route or 'imports'}: {int((gap > 1).sum())} days where the points don't add up to "
                  f"the published Total (max gap {gap.max():.0f} thousand m3)", flush=True)
    df.index.name = "date"
    return df


parse_export_table = parse_report_table     # name used by earlier discovery scripts


def enargas_session():
    s = requests.Session()
    s.headers.update({"User-Agent": BROWSER_UA})
    s.get(ENARGAS_PAGE, params={"tipo": "exp_dentro"}, timeout=TIMEOUT)     # session cookie, as a browser gets
    return s


def fetch_report_tables(tables, start, end):
    """ENARGAS daily export/import tables ({tipo_list: route}), start..end (Timestamps), in requests of at most
    CHUNK_DAYS days."""
    s = enargas_session()
    frames = []
    for tipo, route in tables.items():
        parts = []
        d0 = start
        while d0 <= end:
            d1 = min(d0 + pd.Timedelta(days=CHUNK_DAYS - 1), end)
            r = s.post(ENARGAS_LIST, timeout=TIMEOUT, data={"fecha_desde": d0.strftime("%Y-%m-%d"),
                                                             "fecha_hasta": d1.strftime("%Y-%m-%d"),
                                                             "tipo_list": tipo})
            r.raise_for_status()
            part = parse_report_table(r.text, route)
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


def fetch_exports_daily(start, end):
    return fetch_report_tables(EXPORT_TABLES, start, end)


def fetch_imports_daily(start, end):
    return fetch_report_tables(IMPORT_TABLES, start, end)


def load_daily(path, sheet):
    try:
        old = pd.read_excel(path, sheet_name=sheet, index_col=0)
    except (FileNotFoundError, ValueError):
        return pd.DataFrame()
    old.index = pd.to_datetime(old.index)
    old.index.name = "date"
    return old


def load_exports_daily(path):
    return load_daily(path, EXPORTS_DAILY_SHEET)


def update_daily(path, sheet, start_date, fetch, label):
    """Incremental: keep the archived daily table and fetch only from REFETCH_DAYS before its last day."""
    old = load_daily(path, sheet)
    start = pd.Timestamp(start_date)
    end = pd.Timestamp.today().normalize()
    incremental = not old.empty and old.index.min() <= start
    if incremental:
        start = old.index.max() - pd.Timedelta(days=REFETCH_DAYS)
    print(f"Fetching ENARGAS daily {label} reports {start.date()}..{end.date()} "
          f"({'incremental' if incremental else 'full backfill'})...", flush=True)
    new = fetch(start, end)
    if not incremental or new.empty:
        daily = new if not incremental else old
    else:
        daily = pd.concat([old[old.index < new.index.min()], new]).sort_index()
        daily = daily[~daily.index.duplicated(keep="last")]
    return daily[daily.index >= pd.Timestamp(start_date)] if not daily.empty else daily


def update_exports_daily(path, start_date):
    return update_daily(path, EXPORTS_DAILY_SHEET, start_date, fetch_exports_daily, "export")


def complete_months(daily, col=None):
    """Calendar months in which every day has a value (in `col`, or in any column)."""
    has = daily[col].notna() if col else daily.notna().any(axis=1)
    month = daily.index.to_period("M")
    days = has.groupby(month).sum()
    return days[days.values == days.index.days_in_month].index


def monthly_exports(daily):
    """Daily thousand m3 -> calendar-month million m3, complete months only (every day reported)."""
    if daily.empty:
        return pd.DataFrame(), pd.DataFrame()
    month = daily.index.to_period("M")
    complete = complete_months(daily)
    by_point = daily.fillna(0).groupby(month).sum() / 1000.0
    by_point = by_point.loc[by_point.index.isin(complete)]
    by_point = by_point.loc[:, by_point.ne(0).any()]           # points with no flow since the start date
    by_point.index = by_point.index.to_timestamp()
    by_point.index.name = "date"
    country = by_point.T.groupby(lambda c: c.split(" | ")[0]).sum().T
    country = country[[c for c in COUNTRY_ORDER if c in country] + [c for c in country if c not in COUNTRY_ORDER]]
    country["Total_exports"] = country.sum(axis=1)
    return country.round(3), by_point.round(3)


def monthly_imports(daily):
    """Daily import report (thousand m3/day) -> complete calendar months, million m3/month, English names."""
    if daily.empty:
        return pd.DataFrame()
    month = daily.index.to_period("M")
    complete = complete_months(daily)
    m = daily.fillna(0).groupby(month).sum() / 1000.0
    m = m.loc[m.index.isin(complete)]
    m.index = m.index.to_timestamp()
    m.index.name = "date"
    out = pd.DataFrame(index=m.index)
    out["LNG_Escobar"] = m.get("GNL Escobar", 0.0)
    out["LNG_BahiaBlanca"] = m.get("GNL B. Blanca", 0.0)
    out["Imports_Bolivia"] = m.get("Bolivia", 0.0)
    out["Imports_Chile_GasAndes"] = m.get("Gasandes", 0.0)
    out["Imports_Chile_NorAndino"] = m.get("Norandino", 0.0)
    out["Imports_Chile"] = out["Imports_Chile_GasAndes"] + out["Imports_Chile_NorAndino"]
    out["Total_imports"] = out[["LNG_Escobar", "LNG_BahiaBlanca", "Imports_Bolivia", "Imports_Chile"]].sum(axis=1)
    return out.round(3)


# ------------------------------------------------------------ ENARGAS monthly operating data (xlsx)

def _clean_head(x):
    s = re.sub(r"\(\d+\)|/\d+$", "", str(x)).replace("\n", " ")
    return re.sub(r"\s+", " ", s).strip() if s.strip().lower() not in ("nan", "none") else ""


def read_enargas_pivot(content, sheet, top_key):
    """ENARGAS 'datos estadisticos' sheet: a header row holding `top_key` (e.g. 'TGN'), the sub-header row under
    it, then monthly rows (column 0 = date). Returns a frame indexed by month, columns 'Top | Sub' (or 'Top')."""
    raw = pd.read_excel(io.BytesIO(content), sheet_name=sheet, header=None)
    hit = [i for i in range(min(40, len(raw))) if any(_clean_head(v) == top_key for v in raw.iloc[i].values)]
    if not hit:
        raise ValueError(f"{sheet}: no header row with {top_key!r}")
    r = hit[0]
    top = [_clean_head(v) for v in raw.iloc[r].values]
    sub = [_clean_head(v) for v in raw.iloc[r + 1].values]
    last = ""
    names = []
    for j in range(1, raw.shape[1]):
        last = top[j] or last
        names.append(f"{last} | {sub[j]}" if sub[j] else last)
    dates = pd.to_datetime(raw.iloc[r + 2:, 0], errors="coerce")
    body = raw.iloc[r + 2:, 1:].loc[dates.notna()]
    body.columns = names
    body.index = pd.DatetimeIndex(dates[dates.notna()]).to_period("M").to_timestamp()
    body.index.name = "date"
    return body.apply(pd.to_numeric, errors="coerce")


def fetch_xlsx(url):
    r = requests.get(url, headers={"User-Agent": BROWSER_UA}, timeout=(10, 120))
    r.raise_for_status()
    return r.content


def fetch_grt(start_date):
    """GRT 'Cuenca': gas received by transporters and distributors, thousand m3 (9300 kcal) per month."""
    g = read_enargas_pivot(fetch_xlsx(GRT_URL), "Cuenca", "TGN")
    g = g[g.index >= pd.Timestamp(start_date)]
    parts = g.drop(columns=[c for c in g.columns if c.startswith("Total")])
    if "Total general" in g:
        gap = (parts.sum(axis=1) - g["Total general"]).abs().max()
        print(f"  GRT: {len(g)} months to {g.index.max().date()}; columns {list(g.columns)}; "
              f"max gap parts vs total {gap:.0f} thousand m3", flush=True)
    return g


def fetch_getd(start_date):
    """GETD 'TipoUsuario': gas delivered by user type, thousand m3 (9300 kcal) per month. ENARGAS has the
    sheet titles of its two tabs swapped; this one holds the user types."""
    content = fetch_xlsx(GETD_URL)
    for sheet in ("TipoUsuario", "AreaLicencia"):
        try:
            d = read_enargas_pivot(content, sheet, "Dis")
        except ValueError:
            continue
        if any("Residencial" in c for c in d.columns):
            d = d[d.index >= pd.Timestamp(start_date)]
            print(f"  GETD ({sheet}): {len(d)} months to {d.index.max().date()}; columns {list(d.columns)}", flush=True)
            return d
    raise ValueError("GETD: no sheet with user types")


def fetch_clp(start_date):
    """CLP: TGN/TGS fuel, unaccounted gas (GNNC) and losses, % of gas delivered by each transporter."""
    raw = pd.read_excel(io.BytesIO(fetch_xlsx(CLP_URL)), sheet_name="CLP", header=None)
    r = next(i for i in range(min(30, len(raw))) if str(raw.iat[i, 0]).strip().lower() == "periodo")
    names = {"combustible": "fuel", "gnnc": "unaccounted", "perdidas": "losses", "pérdidas": "losses"}
    cols = []
    for v in raw.iloc[r, 1:].values:
        who, what = (str(v).split("\n") + [""])[:2]
        cols.append(f"{who.strip()}_{names.get(what.strip().lower(), what.strip().lower())}_pct")
    dates = pd.to_datetime(raw.iloc[r + 1:, 0], errors="coerce")
    body = raw.iloc[r + 1:, 1:].loc[dates.notna()].apply(pd.to_numeric, errors="coerce")
    body.columns = cols
    body.index = pd.DatetimeIndex(dates[dates.notna()]).to_period("M").to_timestamp()
    body.index.name = "date"
    return body[body.index >= pd.Timestamp(start_date)].round(3)


def col_like(df, *keys):
    """Sum of the columns whose name contains every key (case-insensitive); 0 if none."""
    hit = [c for c in df.columns if all(k.lower() in c.lower() for k in keys)]
    return df[hit].sum(axis=1, min_count=1) if hit else pd.Series(0.0, index=df.index)


# ------------------------------------------------------------ ENARGAS daily transport report (PDF)

NUM = r"(-?\d+(?:[.,]\d+)?)"


def _num(s):
    return float(s.replace(",", ".")) if s is not None else None


def parse_transport_parte(text, tables):
    """Text + tables of one 'Parte Diario Operativo - Gas Natural Transportado' PDF -> dict, mcm/d (9300 kcal).
    Injection_total includes the imports given in its footnotes: (a) Bolivia y Norandino in TGN Norte,
    (d) GNL Escobar y Gasandes in TGN Centro Oeste, (b) GNL Bahia Blanca in TGS Neuba II (later footnote (b)
    is the Perito Moreno pipeline, domestic gas), (e) the peak-shaving plant."""
    t = re.sub(r"[ \t]+", " ", text)
    out = {}
    m = re.search(r"Inyecci[oó]n\s*Total\s*\(e\)\s*" + NUM, t)
    out["Injection_total"] = _num(m.group(1)) if m else None
    m = re.search(r"Inyecci[oó]n\s*Total\s*\(e\)\s*-?[\d.,]+[^\n]*?\)\s*" + NUM + r"\s+" + NUM, t)
    out["LinePack"], out["LinePack_change"] = (_num(m.group(1)), _num(m.group(2))) if m else (None, None)
    for key, pat in (("Incl_Bolivia_NorAndino", r"\(a\)\s*Incluye[^()\n]*?Bolivia[^()\n]*?\(\s*" + NUM + r"\s*\)"),
                     ("Incl_Escobar_GasAndes", r"\(d\)\s*Incluye[^()\n]*?Escobar[^()\n]*?\(\s*" + NUM + r"\s*\)"),
                     ("Incl_LNG_BahiaBlanca", r"\(b\)\s*Incluye[^()\n]*?(?:GNL|Blanca)[^()\n]*?\(\s*" + NUM + r"\s*\)"),
                     ("Incl_PeritoMoreno", r"\(b\)\s*Incluye\s*(?:GPM|[^()\n]*?Perito)[^()\n]*?\(\s*" + NUM + r"\s*\)"),
                     ("Incl_PeakShaving", r"\(e\)\s*Incluye[^()\n]*?Peak\s*Shaving\s*\(\s*" + NUM + r"\s*\)")):
        m = re.search(pat, t, re.I)
        out[key] = _num(m.group(1)) if m else None
    for basin in ("Norte", "Neuquina", "Austral"):
        out[f"Basin_{basin}"] = None
        for tab in tables:
            for row in tab:
                for cell in row:
                    mm = re.fullmatch(rf"\s*{basin}\s*\n\s*{NUM}\s*", str(cell or ""))
                    if mm and out[f"Basin_{basin}"] is None:
                        out[f"Basin_{basin}"] = _num(mm.group(1))
        if out[f"Basin_{basin}"] is None:     # text fallback: the basin value on the line after its name
            mm = re.search(rf"(?m)^{basin}\s*\n\s*{NUM}\s*$", t)
            if mm:
                out[f"Basin_{basin}"] = _num(mm.group(1))
    known = [out[f"Basin_{b}"] for b in ("Norte", "Neuquina", "Austral")]
    if out["Injection_total"] is not None and sum(v is None for v in known) == 1:
        # the total is the sum of the three basins: a basin the PDF layout hides is the total less the other two
        miss = ("Norte", "Neuquina", "Austral")[known.index(None)]
        out[f"Basin_{miss}"] = round(out["Injection_total"] - sum(v for v in known if v is not None), 2)
    return out


def fetch_transport_day(session, day):
    """One day's transport report (PDF) -> dict, or None when ENARGAS has no report for that day."""
    import pdfplumber
    r = session.get(PARTE_URL, timeout=TIMEOUT, params={"tipo": "transporte", "path": "partes-diarios/transporte",
                                                        "file": day.strftime("%Y%m%d") + ".pdf"})
    if r.status_code != 200 or not r.content.startswith(b"%PDF"):
        return None
    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
        page = pdf.pages[0]
        rec = parse_transport_parte(page.extract_text() or "", page.extract_tables())
    return rec


def update_transport_daily(path, start_date):
    """Daily transport reports: keep archived days, re-read the last REFETCH_DAYS and backfill missing days
    (newest first) until PARTES_BUDGET_S runs out; later runs carry on with the gap."""
    old = load_daily(path, TRANSPORT_DAILY_SHEET)
    end = pd.Timestamp.today().normalize() - pd.Timedelta(days=1)
    days = pd.date_range(pd.Timestamp(start_date), end, freq="D")
    have = set(old.index[old["Injection_total"].notna()]) if "Injection_total" in old else set()
    tried = set(old.index) if not old.empty else set()
    recent = end - pd.Timedelta(days=REFETCH_DAYS)
    todo = [d for d in reversed(days) if d >= recent or d not in tried]
    print(f"Fetching ENARGAS daily transport reports: {len(have)} days archived, {len(todo)} to fetch "
          f"(budget {PARTES_BUDGET_S}s)...", flush=True)
    s = enargas_session()
    s.get(ENARGAS + "dod-partes-dist-trans.php", timeout=TIMEOUT)
    t0, recs, fails = time.time(), {}, 0
    for d in todo:
        if time.time() - t0 > PARTES_BUDGET_S:
            print(f"  time budget reached; {len(todo) - len(recs)} days left for later runs", flush=True)
            break
        try:
            rec = fetch_transport_day(s, d)
        except Exception as e:  # noqa: BLE001 - one unreadable PDF must not stop the archive; retried next run
            print(f"  {d.date()}: {type(e).__name__}: {str(e)[:120]}", flush=True)
            fails += 1
            if fails > 20:
                print("  too many failures; stopping this run", flush=True)
                break
            continue
        recs[d] = rec or {"Injection_total": None}      # None: no report that day (not retried unless recent)
        if rec and rec.get("Injection_total") is None:
            print(f"  {d.date()}: report found but not parsed", flush=True)
    new = pd.DataFrame.from_dict(recs, orient="index").apply(pd.to_numeric, errors="coerce")
    if new.empty:
        return old
    print(f"  fetched {len(new)} days ({int(new['Injection_total'].notna().sum())} parsed) "
          f"in {time.time() - t0:.0f}s", flush=True)
    daily = pd.concat([old.drop(index=[d for d in new.index if d in old.index]), new]).sort_index() \
        if not old.empty else new.sort_index()
    daily.index.name = "date"
    return daily


def supply_daily(transport, imports_daily):
    """Daily domestic injection (transport reports) and imports (import reports), mcm/d."""
    t = transport.copy()
    out = pd.DataFrame(index=t.index.union(imports_daily.index) if not imports_daily.empty else t.index)
    out.index.name = "date"
    for c in ("Basin_Norte", "Basin_Neuquina", "Basin_Austral", "Injection_total", "Incl_Bolivia_NorAndino",
              "Incl_Escobar_GasAndes", "Incl_LNG_BahiaBlanca", "Incl_PeritoMoreno", "Incl_PeakShaving",
              "LinePack", "LinePack_change"):
        if c in t:
            out[f"{c}_mcmd"] = t[c]
    imp = t[[c for c in ("Incl_Bolivia_NorAndino", "Incl_Escobar_GasAndes", "Incl_LNG_BahiaBlanca") if c in t]]
    if "Injection_total" in t:
        dom = t["Injection_total"] - imp.fillna(0).sum(axis=1)
        out["Domestic_injection_mcmd"] = dom.where(t["Injection_total"].notna())
        if "Basin_Norte" in t and "Incl_Bolivia_NorAndino" in t:
            out["Domestic_Norte_mcmd"] = t["Basin_Norte"] - t["Incl_Bolivia_NorAndino"].fillna(0)
    if not imports_daily.empty:
        for raw, name in IMPORT_NAMES.items():
            if raw in imports_daily:
                out[f"{name} (thousand m3)"] = imports_daily[raw]
        k = imports_daily.reindex(out.index)
        out["LNG_Escobar_mcmd"] = k.get("GNL Escobar") / 1000
        out["LNG_BahiaBlanca_mcmd"] = k.get("GNL B. Blanca") / 1000
        out["Imports_Bolivia_mcmd"] = k.get("Bolivia") / 1000
        out["Imports_Chile_mcmd"] = (k.get("Gasandes").fillna(0) + k.get("Norandino").fillna(0)).where(
            k.notna().any(axis=1)) / 1000
        out["Total_imports_mcmd"] = out[["LNG_Escobar_mcmd", "LNG_BahiaBlanca_mcmd", "Imports_Bolivia_mcmd",
                                         "Imports_Chile_mcmd"]].sum(axis=1, min_count=1)
        if "Domestic_injection_mcmd" in out:
            out["Total_supply_mcmd"] = out["Domestic_injection_mcmd"] + out["Total_imports_mcmd"]
    return out.dropna(how="all").round(3)


def supply_net(grt, getd, exports_points, imports_m, clp, transport):
    """Monthly net supply, million m3/month (see module docstring)."""
    idx = grt.index.union(imports_m.index) if not imports_m.empty else grt.index
    g = (grt / 1000.0).reindex(idx)
    imp = imports_m.reindex(index=idx, columns=IMPORT_COLS)
    out = pd.DataFrame(index=idx)
    out.index.name = "date"
    out["Domestic_Neuquina"] = (col_like(g, "TGN", "Neuquina") + col_like(g, "TGS", "Neuquina")
                                - imp["Imports_Chile_GasAndes"].fillna(0))
    out["Domestic_Noroeste"] = (col_like(g, "TGN", "Noroeste") - imp["Imports_Bolivia"].fillna(0)
                                - imp["Imports_Chile_NorAndino"].fillna(0))
    out["Domestic_SanJorge"] = col_like(g, "San Jorge")
    out["Domestic_Austral"] = col_like(g, "Austral")
    out["Domestic_distributor_pipelines"] = col_like(g, "Distribuidoras", "Propios")
    domestic = ["Domestic_Neuquina", "Domestic_Noroeste", "Domestic_SanJorge", "Domestic_Austral",
                "Domestic_distributor_pipelines"]
    out["Domestic_injection"] = out[domestic].sum(axis=1)
    out.loc[~idx.isin(grt.index), domestic + ["Domestic_injection"]] = None          # GRT not published yet
    out.loc[imp["Total_imports"].isna(), ["Domestic_Neuquina", "Domestic_Noroeste", "Domestic_injection"]] = None
    out["Wellhead_power_plants"] = (col_like(getd.reindex(idx), "Off", "Centrales") / 1000.0
                                    if not getd.empty else float("nan"))
    out["Exports_producer_pipelines"] = float("nan")
    if not exports_points.empty:
        fuera = [c for c in exports_points.columns if c.endswith("producer pipeline")]
        out["Exports_producer_pipelines"] = exports_points[fuera].sum(axis=1).reindex(idx)
    out["Net_domestic_supply"] = out[["Domestic_injection", "Wellhead_power_plants", "Exports_producer_pipelines"]] \
        .sum(axis=1, min_count=3)
    for c in ("LNG_Escobar", "LNG_BahiaBlanca", "Imports_Bolivia", "Imports_Chile", "Total_imports"):
        out[c] = imp[c]
    out["Total_net_supply"] = out["Net_domestic_supply"] + out["Total_imports"]
    out["GRT_LNG_other_origins_reported"] = col_like(g, "Otros Orígenes").where(idx.isin(grt.index))
    if not transport.empty and "Injection_total" in transport:
        sd = supply_daily(transport, pd.DataFrame())
        ok = complete_months(sd, "Domestic_injection_mcmd")
        m = sd["Domestic_injection_mcmd"].groupby(sd.index.to_period("M")).sum()
        m = m[m.index.isin(ok)]
        m.index = m.index.to_timestamp()
        out["Domestic_injection_transport_reports"] = m.reindex(idx)
    if not clp.empty:
        out = out.join(clp.reindex(idx))
    return out.dropna(subset=["Domestic_injection", "Total_imports"], how="all").round(3)


def balance_check(national, getd, exports, net):
    """Net supply + imports vs deliveries (sector series + other GETD users) + exports, million m3/month."""
    sectors = national[list(SECTOR_IDS)].sum(axis=1, min_count=6)
    b = pd.DataFrame(index=net.index)
    b.index.name = "date"
    b["Deliveries_by_sector"] = sectors.reindex(b.index)
    if not getd.empty:
        g = getd.reindex(b.index) / 1000.0
        b["Other_deliveries_subdistributors_Cerri"] = (col_like(g, "Subdistribuidor") + col_like(g, "Tra", "RTP"))
    b["Exports"] = exports["Total_exports"].reindex(b.index) if "Total_exports" in exports else None
    b["Deliveries_plus_exports"] = b[[c for c in b.columns]].sum(axis=1, min_count=len(b.columns))
    b["Total_net_supply"] = net["Total_net_supply"]
    b["Residual"] = b["Total_net_supply"] - b["Deliveries_plus_exports"]
    b["Residual_pct_of_supply"] = 100 * b["Residual"] / b["Total_net_supply"]
    return b.dropna(subset=["Residual"]).round(3)


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
    df = pd.DataFrame(payload["data"], columns=["date"] + id_list).set_index("date")
    df.index = pd.to_datetime(df.index)
    id_to_name = {v: k for k, v in ids.items()}
    return df.rename(columns=id_to_name).sort_index()


def keep_archived(fn, path, sheet, label):
    """Run a fetch; on failure keep the archived sheet so one ENARGAS outage doesn't wipe the history."""
    try:
        return fn()
    except (requests.RequestException, ValueError, KeyError, StopIteration) as e:
        print(f"  {label} failed ({type(e).__name__}: {e}); keeping the archived table", flush=True)
        return load_daily(path, sheet)


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

    exports_daily = keep_archived(lambda: update_exports_daily(args.out, args.start_date), args.out,
                                  EXPORTS_DAILY_SHEET, "ENARGAS export reports")
    exports, exports_points = monthly_exports(exports_daily)
    if not exports.empty:
        print(f"  exports: {len(exports)} complete months, {exports.index.min().date()} to "
              f"{exports.index.max().date()}; daily reports to {exports_daily.index.max().date()}", flush=True)
        print(exports.tail(3).to_string(), flush=True)

    imports_daily = keep_archived(lambda: update_daily(args.out, IMPORTS_DAILY_SHEET, args.start_date,
                                                       fetch_imports_daily, "import"),
                                  args.out, IMPORTS_DAILY_SHEET, "ENARGAS import reports")
    imports_m = monthly_imports(imports_daily)
    if not imports_m.empty:
        print(f"  imports: {len(imports_m)} complete months to {imports_m.index.max().date()}", flush=True)

    transport = keep_archived(lambda: update_transport_daily(args.out, args.start_date), args.out,
                              TRANSPORT_DAILY_SHEET, "ENARGAS daily transport reports")

    print("Fetching ENARGAS monthly operating data (GRT, GETD, CLP)...", flush=True)
    try:
        grt = fetch_grt(args.start_date)
    except (requests.RequestException, ValueError) as e:
        print(f"  GRT failed ({type(e).__name__}: {e})", flush=True)
        grt = pd.DataFrame()
    try:
        getd = fetch_getd(args.start_date)
    except (requests.RequestException, ValueError) as e:
        print(f"  GETD failed ({type(e).__name__}: {e})", flush=True)
        getd = pd.DataFrame()
    try:
        clp = fetch_clp(args.start_date)
    except (requests.RequestException, ValueError, StopIteration) as e:
        print(f"  CLP failed ({type(e).__name__}: {e})", flush=True)
        clp = pd.DataFrame()

    net = supply_net(grt, getd, exports_points, imports_m, clp, transport) if not grt.empty else pd.DataFrame()
    daily_supply = supply_daily(transport, imports_daily) if not (transport.empty and imports_daily.empty) \
        else pd.DataFrame()
    check = balance_check(national, getd, exports, net) if not net.empty else pd.DataFrame()
    if not net.empty:
        print("Supply net (million m3/month), last months:", flush=True)
        print(net[["Domestic_injection", "Net_domestic_supply", "Total_imports", "Total_net_supply",
                   "GRT_LNG_other_origins_reported"] +
                  (["Domestic_injection_transport_reports"] if "Domestic_injection_transport_reports" in net else [])]
              .tail(16).to_string(), flush=True)
    if not check.empty:
        print("Balance check (million m3/month):", flush=True)
        print(check.to_string(), flush=True)
    if not getd.empty:   # the SE sector series are GETD's user types summed over Dis/Tra/Off: check
        g = getd.reindex(national.index) / 1000.0
        se_power = national["centrales_electricas"]
        print(f"  check: SE power vs GETD power, max abs diff "
              f"{(se_power - col_like(g, 'Centrales')).abs().max():.1f} million m3", flush=True)

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
        "Monthly sheets: millones de metros cubicos (million m3) per month, as published (m3 of 9300 kcal for the "
        "ENARGAS data). Daily sheets: as published - thousand m3/day for the export and import reports, million "
        "m3/day (mcm/d) for the transport reports - plus mcm/d columns. Charts show monthly averages in mcm/d "
        "(million m3 per month / days in the month).",
        "",
        "SECTORS",
        "'National': produccion_gas_natural (GROSS national gas production, at the wellhead, before field use, "
        "reinjection, venting/flaring and processing) plus 6 use-sector consumption columns - "
        "residencial, comercial, entes_oficiales (official/government bodies), industria, centrales_electricas "
        "(gas burned for power generation, including plants supplied at the wellhead), gnc (compressed gas for "
        "vehicles). These six are ENARGAS's 'gas entregado por tipo de usuario' (GETD) summed over distribution, "
        "transport and off-system users. 'By sector (long)' reshapes the 6 sector columns long.",
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
        "",
        "NET SUPPLY",
        "'Supply net' (million m3/month, from 2021) - measured, nothing estimated or netted off:",
        "Domestic_* / Domestic_injection: ENARGAS 'Gas Recibido por Transportistas y Distribuidoras de "
        "Productores y Otros Origenes - por Cuenca' (GRT.xlsx): gas metered into the TGN and TGS transport "
        "systems and into the distributors' own pipelines (Camuzzi Gas del Sur, Cuyana), by basin. It is net of "
        "field use, reinjection, venting/flaring and gas-plant shrinkage, which all happen before these receipt "
        "points. Neuquina = TGN + TGS Neuquina; Noroeste = TGN Noroeste; San Jorge and Austral = TGS; "
        "Domestic_distributor_pipelines = 'Distribuidoras - Propios'. GRT's 'Otros Origenes' (LNG) columns are "
        "left out (kept for reference as GRT_LNG_other_origins_reported; they show Escobar until 2025 and 0 in "
        "2026). Pipeline imports metered inside GRT's basin columns are taken out: Bolivia and NorAndino from "
        "Noroeste (they enter TGN's Norte pipeline), GasAndes from Neuquina (TGN Centro Oeste), per the "
        "footnotes of ENARGAS's daily transport reports.",
        "Wellhead_power_plants: power plants supplied directly at the wellhead, outside the transport system "
        "(ENARGAS GETD 'Off System - Centrales electricas'). Exports_producer_pipelines: exports through "
        "producers' own pipelines (ENARGAS export reports 'fuera del sistema'). Both are in the sector and export "
        "series but not in GRT. Net_domestic_supply = Domestic_injection + Wellhead_power_plants + "
        "Exports_producer_pipelines.",
        "Imports (ENARGAS 'Partes de Importacion', daily, summed to complete months): LNG_Escobar (GNL Escobar "
        "FSRU), LNG_BahiaBlanca (GNL Bahia Blanca FSRU, 2021-2023 winters), Imports_Bolivia, Imports_Chile "
        "(GasAndes + NorAndino back-flows). Total_net_supply = Net_domestic_supply + Total_imports.",
        "Domestic_injection_transport_reports: the same domestic injection from the daily transport reports "
        "(provisional, TGN + TGS only, without the distributors' own pipelines), complete months - a cross-check.",
        "TGN_/TGS_fuel_pct, _unaccounted_pct, _losses_pct: transport fuel (compressors), gas not accounted for "
        "(GNNC) and losses, in % of gas delivered by each transporter, as ENARGAS publishes them (CLP.xlsx). "
        "ENARGAS gives no volumes for them, so they are not turned into a volume here.",
        "'Supply daily': Basin_*_mcmd and Injection_total_mcmd = ENARGAS 'Parte Diario Operativo - Gas Natural "
        "Transportado' (one PDF per day, provisional data from TGN and TGS): injection by basin and total, mcm/d "
        "of 9300 kcal. Incl_* = imports and other sources the report says are inside those figures (Bolivia + "
        "NorAndino in Norte; GNL Escobar + GasAndes in Centro Oeste; GNL Bahia Blanca in Neuba II; later the "
        "Perito Moreno pipeline (domestic) in Neuba II; peak-shaving plant). Domestic_injection_mcmd = "
        "Injection_total minus Bolivia+NorAndino, Escobar+GasAndes and Bahia Blanca LNG. LinePack_mcmd = "
        "transport line-pack (million m3) and its change on the day. Import columns: the daily import reports, "
        "thousand m3/day as published, and mcm/d. Each run re-reads the last 45 days and backfills missing days "
        f"within a {PARTES_BUDGET_S}s budget.",
        "'Balance check': Total_net_supply against deliveries by sector (the 'National' sector columns) + "
        "sub-distributors and the Cerri processing plant (GETD) + exports. The residual is not a loss estimate: "
        "it holds transport fuel and losses (~2% of deliveries, CLP), line-pack changes and the distributors' "
        "billing-cycle timing (their deliveries are billed volumes, which lag physical flows in winter).",
        "'Imports daily' / 'Transport daily' / 'Deliveries (ENARGAS)': the raw archives behind the above.",
        "",
        "SOURCE",
        "apis.datos.gob.ar/series/api/series - Secretaria de Energia, Ministerio de Economia, dataset "
        f"'Produccion y consumo de gas natural'. Monthly; this archive starts {args.start_date} (the source "
        "goes back to 1996-01 - rerun with --start-date to pull more).",
        "Exports and imports: ENARGAS (Ente Nacional Regulador del Gas y la Electricidad), Partes diarios de "
        "exportacion (dentro / fuera del sistema) and de importacion: " + ENARGAS_PAGE +
        "?tipo=exp_dentro, ?tipo=exp_fuera and ?tipo=importaciones",
        "Net supply: ENARGAS Datos operativos de transporte y distribucion - " + GRT_URL + ", " + GETD_URL + ", " +
        CLP_URL + "; daily transport reports: " + ENARGAS + "dod-partes-dist-trans.php (column 'Transporte').",
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
    if not net.empty:
        sheets["Supply net"] = net
    if not daily_supply.empty:
        sheets["Supply daily"] = daily_supply
    if not check.empty:
        sheets["Balance check"] = check
    if not imports_daily.empty:
        sheets[IMPORTS_DAILY_SHEET] = imports_daily
    if not transport.empty:
        sheets[TRANSPORT_DAILY_SHEET] = transport
    if not getd.empty:
        sheets["Deliveries (ENARGAS)"] = (getd / 1000.0).round(3)
    xlsx_notes.write_workbook(args.out, sheets, notes,
                              {"UNITS", "SECTORS", "REGIONS", "EXPORTS", "NET SUPPLY", "SOURCE"})
    print(f"Saved {args.out}", flush=True)
    print(national_out.tail().to_string(), flush=True)


if __name__ == "__main__":
    main()
