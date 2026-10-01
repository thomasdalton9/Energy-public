"""
Peru natural gas demand by sector, monthly from 2021, plus production by
lot - from the Ministry of Energy and Mines (MINEM, Direccion General de
Hidrocarburos), gob.pe collection "Informes Estadisticos Upstream -
Downstream":
  https://www.gob.pe/institucion/minem/colecciones/17643-informes-estadisticos-upstream-downstream

Demand: each monthly "Informe Estadistico Downstream" carries
'distribucion-<mes>-<anio>.xlsx' (Informe de distribucion de gas natural):
one sheet per distribution concession - Calidda (Lima & Callao), Contugas
(Ica), Quavii (Concesion Norte), Petroperu (Concesion Suroeste, ex-Naturgy)
and Gasnorp (Piura) - with "VOLUMEN DE GAS NATURAL DISTRIBUIDO POR SECTOR"
in MMPCD (million cubic feet per day, monthly average) for Residencial,
Comercial, Industrial, GNV, Generacion Electrica, Instituciones publicas
and the total, from Jan-2022 to the latest month. 2021 is not in those
files; it comes once (backfill) from MINEM's Anuario Estadistico de
Hidrocarburos 2021 (gob.pe collection 25333), whose tables give the same
concession x sector volumes in MMPC per month (divided by days in month).

Production: each monthly "Informe Estadistico Upstream" carries
'7-produccion-fiscalizada-de-gas-natural.xlsx' - fiscalised gas production
by lot in MPCD (thousand cubic feet per day) from Jan-2019: Camisea lots
88 and 56 (Pluspetrol) and 57 (Repsol), Aguaytia (31C) and the northwest
(Talara coast + offshore Zocalo).

Sectors written (mcm/d = million m3 per day; MMPCD x 0.0283168):
  Power                  = Generacion Electrica (category GE)
  Industrial             = Industrial (incl. Talara refinery and fishmeal in Gasnorp / Quavii)
  Vehicle_CNG            = GNV (CNG filling stations)
  Residential_commercial = Residencial + Comercial + Instituciones publicas
  Total                  = sum of the concessions' TOTAL rows
No petrochemical use is published (Peru has no gas petrochemical plant).

Cross-check: the 'Power cross-check' sheet sets the Power sector against
COES gas-fired generation (peru_power_generation_daily.xlsx, if present)
and gives the implied heat rate at 1,065 Btu/scf (Camisea gas, Perupetro).

Incremental: keeps every month already in the workbook's raw sheets,
downloads only the latest downstream and upstream files (each holds the
full window), refreshes the last REFRESH_MONTHS months from them and adds
new months; the 2021 backfill runs only while 2021 months are missing.
Not reachable from the editing sandbox; runs in GitHub Actions.
"""
import argparse
import datetime
import io
import os
import re
import sys
import time
import unicodedata

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root
import xlsx_charts  # noqa: E402
import xlsx_notes  # noqa: E402

GOB = "https://www.gob.pe"
COLL = GOB + "/institucion/minem/colecciones/17643-informes-estadisticos-upstream-downstream"
ANUARIO_COLL = GOB + "/institucion/minem/colecciones/25333-anuario-estadistico-de-hidrocarburos"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept-Language": "es-PE,es;q=0.9"}
T = (15, 120)
DATA_START = pd.Timestamp(2021, 1, 1)
REFRESH_MONTHS = 6
MCF_TO_M3 = 0.0283168466  # 1 cubic foot in m3, so MMPCD x this = million m3/day
BTU_PER_SCF = 1065.0  # Camisea gas HHV: Perupetro daily production report, Lot 88 MMBtu / Mcf ~ 1.065-1.068
MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
         "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
MES3 = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "sep": 9,
        "oct": 10, "nov": 11, "dic": 12}
SECTOR_MAP = {"residencial": "Residential", "comercial": "Commercial", "instituciones publicas": "Public_institutions",
              "industrial": "Industrial", "gnv": "Vehicle_CNG", "generacion electrica": "Power"}
SECTORS = ["Power", "Industrial", "Vehicle_CNG", "Residential_commercial"]
CONCESSIONS = {"calidda": "Calidda", "contugas": "Contugas", "quavii": "Quavii", "petroperu": "Petroperu",
               "gasnorp": "Gasnorp"}
DEMAND_SHEET, RAW_SHEET = "Demand by sector", "By concession (MMPCD)"
PROD_SHEET, CHECK_SHEET = "Production by lot", "Power cross-check"
LNG_SHEET, DAILY_SHEET = "LNG exports", "Daily power burn (est.)"
PROD_COLS = ["Lot_88", "Lot_56", "Lot_57", "Aguaytia_31C", "Northwest", "Other", "Total"]
COES_FILE = "peru_power_generation_daily.xlsx"


def out(*a):
    print(*a, flush=True)


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return " ".join(s.replace("\n", " ").split())


def get(url, **kw):
    for attempt in range(3):
        try:
            r = requests.get(url, headers=H, timeout=T, **kw)
            if r.status_code == 200:
                return r
            out(f"  {url}: HTTP {r.status_code}")
        except requests.RequestException as e:
            out(f"  {url}: {type(e).__name__}: {str(e)[:150]}")
        time.sleep(5 * (attempt + 1))
    return None


# ------------------------------------------------------------------ gob.pe

def publications(coll, pages=3):
    """[(path, title-slug)] of a gob.pe collection, newest first."""
    seen = []
    for page in range(1, pages + 1):
        r = get(f"{coll}?sheet={page}")
        if r is None:
            break
        new = [u for u in re.findall(r'href="(/institucion/minem/informes-publicaciones/\d+-[^"?#]+)"', r.text)
               if u not in seen]
        seen += new
        if not new:
            break
    return seen


def slug_month(path):
    """'...-informe-estadistico-downstream-agosto-2026' -> Timestamp(2026-08-01)."""
    m = re.search(r"(" + "|".join(MESES) + r")-?(\d{4})$", path)
    return pd.Timestamp(int(m.group(2)), MESES[m.group(1)], 1) if m else None


def latest_attachment(kind, pattern):
    """Newest publication of `kind` (downstream/upstream) and its attachment matching `pattern`."""
    pubs = [(slug_month(p), p) for p in publications(COLL) if f"informe-estadistico-{kind}" in p]
    pubs = sorted([x for x in pubs if x[0] is not None], reverse=True)
    for month, path in pubs[:4]:
        r = get(GOB + path)
        if r is None:
            continue
        files = sorted(set(re.findall(r'https://cdn\.www\.gob\.pe/uploads/document/file/\d+/[^"?#\s]+', r.text)))
        hit = [f for f in files if re.search(pattern, f, re.I)]
        if hit:
            out(f"{kind}: {path} -> {hit[0]}")
            return month, hit[0]
        out(f"{kind}: {path} has no file matching {pattern}")
    return None, None


# ------------------------------------------------------------------ parsers

def parse_distribution(content):
    """Distribution workbook -> wide frame: Month index, columns '<Concession>|<Sector>' in MMPCD."""
    frames = {}
    for sheet, df in pd.read_excel(io.BytesIO(content), sheet_name=None, header=None).items():
        conc = next((v for k, v in CONCESSIONS.items() if k in norm(sheet)), sheet.strip().title())
        cells = df.map(norm)
        vol = [i for i in df.index if cells.loc[i].str.contains("volumen de gas natural distribuido").any()]
        if not vol:
            out(f"  sheet {sheet}: no volume table")
            continue
        hdr = next(i for i in df.index if i > vol[0] and (cells.loc[i] == "sector").any())
        scol = int((cells.loc[hdr] == "sector").idxmax())
        dates = {c: pd.Timestamp(v).to_period("M").to_timestamp() for c, v in df.loc[hdr].items()
                 if c > scol and isinstance(v, (datetime.datetime, pd.Timestamp))}
        for i in range(hdr + 1, len(df)):
            label = cells.loc[i, scol]
            if not label or label == "nan":
                break
            key = "Total" if label.startswith("total") else SECTOR_MAP.get(label)
            if key is None:
                out(f"  sheet {sheet}: unknown sector {label!r}")
                key = re.sub(r"\W+", "_", label).strip("_").title()
            frames[f"{conc}|{key}"] = pd.Series({d: pd.to_numeric(df.loc[i, c], errors="coerce")
                                                 for c, d in dates.items()})
            if key == "Total":
                break
    wide = pd.DataFrame(frames).sort_index()
    wide.index.name = "Month"
    return wide


LOTS = {"88": "Lot_88", "56": "Lot_56", "57": "Lot_57", "31c": "Aguaytia_31C"}


def parse_production(content):
    """Fiscalised gas production workbook -> Month x lot frame in MMPCD."""
    df = pd.read_excel(io.BytesIO(content), header=None)
    cells = df.map(norm)
    hdr = next(i for i in df.index if (cells.loc[i] == "lote").any())
    years = df.loc[hdr - 1].ffill()
    cols = {}
    for c, v in df.loc[hdr].items():
        m = MES3.get(norm(v)[:3]) if isinstance(v, str) else None
        if m and pd.notna(years[c]) and str(years[c]).strip().isdigit():
            cols[c] = pd.Timestamp(int(years[c]), m, 1)
    lot_col = int((cells.loc[hdr] == "lote").idxmax())
    comp_col = lot_col - 1
    rows = {}
    for i in range(hdr + 1, len(df)):
        lot, comp = cells.loc[i, lot_col], cells.loc[i, comp_col]
        name = LOTS.get(lot.replace(" ", "")) if lot != "nan" else None
        if comp.startswith("total costa"):
            name = "Coast"
        elif comp.startswith("total zocalo"):
            name = "Offshore"
        elif comp.startswith("total pais"):
            name = "Total"
        if name:
            rows[name] = pd.Series({d: pd.to_numeric(df.loc[i, c], errors="coerce") for c, d in cols.items()})
    p = pd.DataFrame(rows).sort_index() / 1000.0  # MPCD -> MMPCD
    p["Northwest"] = p.get("Coast", 0) + p.get("Offshore", 0)
    named = [c for c in ("Lot_88", "Lot_56", "Lot_57", "Aguaytia_31C", "Northwest") if c in p]
    p["Other"] = (p["Total"] - p[named].sum(axis=1)).round(6)
    p.index.name = "Month"
    return p[named + ["Other", "Total"]]


ANUARIO_CONCESSIONS = {"lima": "Calidda", "ica": "Contugas", "norte": "Quavii", "sur oeste": "Petroperu",
                       "suroeste": "Petroperu", "piura": "Gasnorp"}


def parse_anuario(content, year):
    """Anuario 'Tablas' workbook -> wide frame like parse_distribution for `year`.
    Tables 'Volumen de distribucion de Gas Natural por sector economico en la Concesion <X> (MMPC), <year>'
    give MMPC per month; divided by days in month -> MMPCD."""
    frames = {}
    for sheet, df in pd.read_excel(io.BytesIO(content), sheet_name=None, header=None).items():
        cells = df.map(norm)
        for i in df.index:
            for c in df.columns:
                t = cells.loc[i, c]
                if not (t.startswith("tabla") and "volumen de distribucion de gas natural por sector economico" in t
                        and str(year) in t):
                    continue
                conc = next((v for k, v in ANUARIO_CONCESSIONS.items() if f"concesion {k}" in t
                             or f"concesion de {k}" in t), None)
                if conc is None:
                    continue
                hdr = next((j for j in range(i + 1, min(i + 6, len(df)))
                            if sum(isinstance(v, (datetime.datetime, pd.Timestamp)) for v in df.loc[j]) >= 12), None)
                if hdr is None:
                    continue
                dates = {k: pd.Timestamp(v).to_period("M").to_timestamp() for k, v in df.loc[hdr].items()
                         if isinstance(v, (datetime.datetime, pd.Timestamp)) and pd.Timestamp(v).year == year}
                lab = min(dates) - 1
                for j in range(hdr + 1, min(hdr + 12, len(df))):
                    label = re.sub(r"\d+$", "", cells.loc[j, lab]).strip()
                    key = "Total" if label.startswith("total") else SECTOR_MAP.get(label)
                    if key is None:
                        continue
                    ser = pd.Series({d: pd.to_numeric(df.loc[j, k], errors="coerce") for k, d in dates.items()})
                    frames[f"{conc}|{key}"] = ser / ser.index.days_in_month
                    if key == "Total":
                        break
    wide = pd.DataFrame(frames).sort_index()
    wide.index.name = "Month"
    return wide


def anuario_backfill(year):
    """MINEM's Anuario Estadistico de Hidrocarburos for `year` (its 'Tablas' xlsx) -> concession x sector MMPCD."""
    for path in publications(ANUARIO_COLL):
        if not path.endswith(str(year)):
            continue
        r = get(GOB + path)
        if r is None:
            continue
        files = sorted(set(re.findall(r'https://cdn\.www\.gob\.pe/uploads/document/file/\d+/[^"?#\s]+', r.text)))
        for f in [f for f in files if re.search(r"tabla.*\.xlsx?$", f, re.I)]:
            rr = get(f)
            if rr is None:
                continue
            wide = parse_anuario(rr.content, year)
            if not wide.empty:
                out(f"anuario {year}: {f} -> {wide.shape[1]} series")
                return wide, f
    out(f"anuario {year}: tables not found")
    return pd.DataFrame(), None


# ------------------------------------------------------------------ archive

def load_sheet(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=0)
        df.index = pd.to_datetime(df.index.astype(str))
        out(f"existing '{sheet}': {len(df)} months {df.index.min():%Y-%m}..{df.index.max():%Y-%m}")
        return df
    except (FileNotFoundError, ValueError) as e:
        out(f"no existing '{sheet}' ({e})")
        return pd.DataFrame()


def merge(old, new, refresh=REFRESH_MONTHS):
    """Keep archived months; add new months; overwrite the last `refresh` archived months with new values."""
    if old.empty:
        return new.sort_index()
    if new.empty:
        return old.sort_index()
    recent = sorted(old.index)[-refresh:]
    base = old.drop(index=[m for m in recent if m in new.index])
    return new.combine_first(base).sort_index()


def sectors(raw):
    """Concession x sector MMPCD -> national sectors (MMPCD)."""
    def total(sector):
        cols = [c for c in raw.columns if c.split("|")[1] == sector]
        return raw[cols].sum(axis=1, min_count=1) if cols else pd.Series(0.0, index=raw.index)
    s = pd.DataFrame(index=raw.index)
    s["Power"] = total("Power")
    s["Industrial"] = total("Industrial")
    s["Vehicle_CNG"] = total("Vehicle_CNG")
    s["Residential_commercial"] = (total("Residential").fillna(0) + total("Commercial").fillna(0)
                                   + total("Public_institutions").fillna(0))
    s["Total"] = total("Total")
    s["Sector_sum_check"] = s[SECTORS].sum(axis=1) - s["Total"]
    return s


def power_check(sec, coes_path):
    """Monthly Power sector vs COES gas-fired generation; also returns the COES daily series."""
    try:
        d = pd.read_excel(coes_path, sheet_name="Daily")
    except (FileNotFoundError, ValueError) as e:
        out(f"COES cross-check skipped ({e})")
        return pd.DataFrame(), pd.Series(dtype=float)
    daily = d.assign(date=pd.to_datetime(d["date"])).set_index("date")["Gas_MWh"].sort_index()
    daily = daily[daily.index >= DATA_START]
    m = daily.resample("MS").agg(["mean", "count"])
    m = m[m["count"] >= 25]
    c = pd.DataFrame(index=sec.index)
    c["Power_MMPCD"] = sec["Power"]
    c["COES_gas_generation_GWh_per_day"] = (m["mean"] / 1000).reindex(c.index)
    # MMPCD x 1e6 scf x Btu/scf over GWh/d x 1e6 kWh
    c["Implied_heat_rate_Btu_per_kWh"] = (c["Power_MMPCD"] * BTU_PER_SCF / c["COES_gas_generation_GWh_per_day"])
    return c.dropna(how="all", subset=["COES_gas_generation_GWh_per_day"]), daily


def daily_power_estimate(check, daily):
    """ESTIMATE: COES daily gas-fired MWh x heat rate -> daily power-sector gas burn.
    Heat rate = that month's implied heat rate where MINEM has published the month (so the daily values average
    to the published monthly Power figure), else the mean of the latest 12 published months."""
    if check.empty or daily.empty:
        return pd.DataFrame()
    hr = check["Implied_heat_rate_Btu_per_kWh"].dropna()
    hr = hr[(hr > 4000) & (hr < 20000)]
    if hr.empty:
        return pd.DataFrame()
    fallback = hr.iloc[-12:].mean()
    month = daily.index.to_period("M").to_timestamp()
    rate = pd.Series(month.map(hr).values, index=daily.index)
    basis = pd.Series("monthly MINEM", index=daily.index).where(rate.notna(), "trailing 12-month mean")
    rate = rate.fillna(fallback)
    e = pd.DataFrame({"COES_gas_generation_MWh": daily.round(1), "Heat_rate_Btu_per_kWh": rate.round(0),
                      "Heat_rate_basis": basis})
    e["Power_burn_estimate_MMPCD"] = (daily * 1000 * rate / BTU_PER_SCF / 1e6).round(3)
    e["Power_burn_estimate_mcm_per_day"] = (e["Power_burn_estimate_MMPCD"] * MCF_TO_M3).round(4)
    e.index.name = "Date"
    return e


# ------------------------------------------------------------------ Perupetro LNG cargoes

PP_LNG = "https://www.perupetro.com.pe/ExportaGAS/EmbarqueGasServlet"
# Perupetro's server omits its intermediate certificate; add GlobalSign's published intermediate (it chains to
# GlobalSign Root R3 in certifi, so a tampered download would simply fail verification - TLS checks stay on).
GLOBALSIGN_INTERMEDIATE = "http://secure.globalsign.com/cacert/gsrsaovsslca2018.crt"
LNG_COLS = ["Cargoes", "LNG_m3", "LNG_tonnes", "MMBtu", "Gas_MMcf"]


def perupetro_verify():
    try:
        requests.head("https://www.perupetro.com.pe/", headers=H, timeout=T)
        return True
    except requests.exceptions.SSLError:
        pass
    except requests.RequestException as e:
        out(f"perupetro unreachable ({type(e).__name__})")
        return None
    import ssl
    import tempfile
    import certifi
    der = requests.get(GLOBALSIGN_INTERMEDIATE, headers=H, timeout=T).content
    pem = der.decode() if der.startswith(b"-----") else ssl.DER_cert_to_PEM_cert(der)
    f = tempfile.NamedTemporaryFile("w", suffix=".pem", delete=False)
    f.write(open(certifi.where()).read() + "\n" + pem)
    f.close()
    return f.name


def lng_cargoes():
    """Every Peru LNG cargo Perupetro lists (date, m3 LNG, tonnes, MMBtu, MPC) -> monthly totals from 2021."""
    import json
    verify = perupetro_verify()
    if verify is None:
        return pd.DataFrame()
    today = datetime.date.today()
    data = {"fechaInicio": DATA_START.strftime("%d/%m/%Y"), "fechaFin": today.strftime("%d/%m/%Y"), "accion": "Todos"}
    for attempt in range(3):
        try:
            r = requests.post(PP_LNG, data=data, headers=H, timeout=T, verify=verify)
            m = re.search(r"\.columns\(\{\s*data:\s*(\[.*?\])\s*,\s*\n", r.text, re.S)
            if r.status_code == 200 and m:
                break
            out(f"  LNG cargoes: HTTP {r.status_code}, data block {'found' if m else 'missing'}")
        except requests.RequestException as e:
            out(f"  LNG cargoes: {type(e).__name__}: {str(e)[:150]}")
        time.sleep(10 * (attempt + 1))
    else:
        return pd.DataFrame()
    c = pd.DataFrame(json.loads(m.group(1)))
    for k in ("m3", "tm", "ccalorico", "mpc"):
        c[k] = pd.to_numeric(c[k].astype(str).str.replace(",", ""), errors="coerce")
    c["fecha"] = pd.to_datetime(c["fecha"], format="%d/%m/%Y", errors="coerce")
    c = c[c["fecha"] >= DATA_START]
    out(f"LNG cargoes: {len(c)} since {DATA_START:%Y-%m}, latest {c['fecha'].max():%Y-%m-%d}")
    g = c.set_index("fecha").resample("MS").agg({"id_embarque": "count", "m3": "sum", "tm": "sum", "ccalorico": "sum",
                                                  "mpc": "sum"})
    g.columns = LNG_COLS
    g["Gas_MMcf"] = g["Gas_MMcf"] / 1000.0  # MPC (thousand ft3) -> million ft3
    g = g[g.index < pd.Timestamp(today).to_period("M").to_timestamp()]  # complete months only
    g.index.name = "Month"
    return g


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/peru_gas_demand_by_sector.xlsx")
    ap.add_argument("--coes", default=None, help="COES workbook for the power cross-check (default: next to --out)")
    args = ap.parse_args()
    coes = args.coes or os.path.join(os.path.dirname(os.path.abspath(args.out)), COES_FILE)

    raw = load_sheet(args.out, RAW_SHEET)
    prod = load_sheet(args.out, PROD_SHEET)
    if not prod.empty:
        prod = prod[[f"{c}_MMPCD" for c in PROD_COLS if f"{c}_MMPCD" in prod]].rename(columns=lambda c: c[:-6])
    lng = load_sheet(args.out, LNG_SHEET)
    if not lng.empty:
        lng = lng[[c for c in LNG_COLS if c in lng]]
    sources = {}

    month, url = latest_attachment("downstream", r"distribucion.*\.xlsx$")
    if url:
        r = get(url)
        if r is not None:
            new = parse_distribution(r.content)
            out(f"distribution file: {new.shape[1]} series, {new.index.min():%Y-%m}..{new.index.max():%Y-%m}")
            raw = merge(raw, new[new.index >= DATA_START])
            sources["distribution"] = url
    missing21 = [m for m in pd.date_range(DATA_START, "2021-12-01", freq="MS") if raw.empty or m not in raw.index
                 or raw.loc[m].isna().all()]
    if missing21:
        old, f = anuario_backfill(2021)
        if not old.empty:
            raw = old.combine_first(raw) if not raw.empty else old
            sources["anuario"] = f
    if raw.empty:
        raise SystemExit("no distribution data fetched and no existing archive")
    order = list(CONCESSIONS.values())
    raw = raw[sorted(raw.columns, key=lambda c: (order.index(c.split("|")[0]) if c.split("|")[0] in order else 99,
                                                 c))]

    month_up, url_up = latest_attachment("upstream", r"produccion-fiscalizada-de-gas-natural.*\.xlsx$")
    if url_up:
        r = get(url_up)
        if r is not None:
            p = parse_production(r.content)
            prod = merge(prod, p[p.index >= DATA_START])
            sources["production"] = url_up
            out(f"production file: {p.index.min():%Y-%m}..{p.index.max():%Y-%m}")

    try:
        new_lng = lng_cargoes()
    except Exception as e:  # noqa: BLE001 - optional sheet; keep the archive if Perupetro misbehaves
        out(f"LNG cargoes failed: {type(e).__name__}: {e}")
        new_lng = pd.DataFrame()
    lng = merge(lng, new_lng) if not new_lng.empty else lng

    sec = sectors(raw)
    gap = sec["Sector_sum_check"].abs().max()
    out(f"{len(sec)} months {sec.index.min():%Y-%m}..{sec.index.max():%Y-%m}; max |sector sum - total| = {gap:.3f} MMPCD")
    out(sec.tail(4).round(1).to_string())

    table = pd.concat([(sec[SECTORS + ["Total"]] * MCF_TO_M3).round(4).add_suffix("_mcm_per_day"),
                       sec[SECTORS + ["Total"]].round(3).add_suffix("_MMPCD")], axis=1)
    check, coes_daily = power_check(sec, coes)
    if not check.empty:
        out(check.tail(6).round(1).to_string())
    est = daily_power_estimate(check, coes_daily)

    sheets = {DEMAND_SHEET: table}
    if not prod.empty:
        pt = pd.concat([(prod * MCF_TO_M3).round(4).add_suffix("_mcm_per_day"), prod.round(3).add_suffix("_MMPCD")],
                       axis=1)
        dom = sec["Total"].reindex(prod.index)
        pt["Domestic_distribution_MMPCD"] = dom.round(3)
        pt["Production_less_domestic_MMPCD"] = (prod["Total"] - dom).round(3)
        sheets[PROD_SHEET] = pt
    if not lng.empty:
        lt = lng.copy()
        lt["LNG_exports_MMPCD"] = (lt["Gas_MMcf"] / lt.index.days_in_month).round(3)
        lt["LNG_exports_mcm_per_day"] = (lt["LNG_exports_MMPCD"] * MCF_TO_M3).round(4)
        lt[["LNG_m3", "LNG_tonnes", "MMBtu", "Gas_MMcf"]] = lt[["LNG_m3", "LNG_tonnes", "MMBtu", "Gas_MMcf"]].round(1)
        sheets[LNG_SHEET] = lt
    if not check.empty:
        sheets[CHECK_SHEET] = check.round(3)
    if not est.empty:
        e = est.copy()
        e.index = e.index.strftime("%Y-%m-%d")
        sheets[DAILY_SHEET] = e
    sheets[RAW_SHEET] = raw.round(4)
    for name, df in sheets.items():
        if name != DAILY_SHEET:
            df.index = pd.DatetimeIndex(df.index).strftime("%Y-%m")
            df.index.name = "Month"

    span = lambda d: f"{d.index.min():%Y-%m}..{d.index.max():%Y-%m}" if not d.empty else "none"  # noqa: E731
    notes = [
        "UNITS",
        "*_mcm_per_day = million cubic metres per day (monthly average); *_MMPCD = million cubic feet per day as "
        "published by MINEM. mcm/d = MMPCD x 0.0283168 (1 ft3 = 0.0283168 m3).",
        f"'{RAW_SHEET}': every concession x sector series as published, MMPCD (2021: Anuario MMPC per month / days).",
        f"'{PROD_SHEET}': fiscalised production by lot (MINEM publishes MPCD = thousand ft3/day; divided by 1,000).",
        f"'{LNG_SHEET}': Peru LNG (Melchorita) cargoes loaded per month - count, m3 of LNG, tonnes, MMBtu and gas "
        "equivalent (Perupetro's MPC / 1,000 = MMcf); LNG_exports_MMPCD = MMcf / days in month. Cargo-based, so lumpy "
        "month to month (one cargo is ~3.3-3.7 Bcf).",
        f"'{CHECK_SHEET}': Power sector MMPCD vs COES gas-fired generation (GWh/day, monthly mean of daily SCADA "
        f"totals) and the implied heat rate at {BTU_PER_SCF:,.0f} Btu/scf.",
        f"'{DAILY_SHEET}': ESTIMATE, not published data - COES daily gas-fired MWh x heat rate. The heat rate is the "
        "month's implied heat rate where MINEM has published the month (so daily values average to the published "
        "Power figure), otherwise the mean of the latest 12 published months.",
        "",
        "SECTORS",
        "Power: 'Generacion Electrica' (tariff category GE) - gas-fired power plants supplied through the "
        "distribution concessions (Chilca, Ventanilla, Santa Rosa, Fenix, Kallpa, Independencia ...).",
        "Industrial: 'Industrial' (categories B-industrial, C, D, E; Gasnorp also Talara refinery REF and fishmeal "
        "PESCA; Quavii and Petroperu supply industry by truck-borne LNG/CNG).",
        "Vehicle_CNG: 'GNV' (CNG filling stations).",
        "Residential_commercial: 'Residencial' + 'Comercial' + 'Instituciones publicas'.",
        "Petrochemical/other: not published - Peru has no gas-based petrochemical plant; no separate line exists.",
        "Total: sum of the concessions' published TOTAL rows (Calidda, Contugas, Quavii, Petroperu, Gasnorp); the "
        "four sector columns add up to it exactly (checked every run). Gasnorp (Piura) starts in 2022 (tiny until "
        "Nov-2022) and is absent from the 2021 Anuario.",
        "Not included: gas used outside the distribution concessions (Camisea field/plant own use, Aguaytia's "
        "Termoselva plant, northwest operators' own use) and Peru LNG feed gas.",
        f"'{PROD_SHEET}': Lot_88 (Camisea, Pluspetrol - mainly domestic market), Lot_56 (Camisea, Pluspetrol - LNG "
        "export), Lot_57 (Repsol, Kinteroni/Sagari), Aguaytia_31C, Northwest = coast (Talara basin) + offshore Zocalo; "
        "Other = remaining lots. Production_less_domestic_MMPCD = Total production - domestic distribution (LNG feed "
        "plus field use and non-concession users).",
        "",
        "COVERAGE",
        f"Demand monthly {span(sec)}; production {span(prod)}; LNG cargoes {span(lng)}; daily power estimate "
        f"{(est.index.min().date() if not est.empty else 'none')}..{(est.index.max().date() if not est.empty else '')}. "
        "MINEM publishes about four weeks after month end. Demand Jan-2022 on from the monthly distribution files "
        f"(the last {REFRESH_MONTHS} months are re-read each run); 2021 from the Anuario Estadistico de Hidrocarburos "
        "2021 tables (fetched once).",
        "",
        "SOURCE",
        f"MINEM / DGH, Informes Estadisticos Upstream - Downstream: {COLL}",
        f"Latest distribution file: {sources.get('distribution', 'n/a this run (archive kept)')}",
        f"Latest production file: {sources.get('production', 'n/a this run (archive kept)')}",
        f"Anuario Estadistico de Hidrocarburos (2021 backfill): {sources.get('anuario', ANUARIO_COLL)}",
        "Perupetro, Embarques de Gas Natural para fines de exportacion: "
        "https://www.perupetro.com.pe/ExportaGAS/Relacion_ES.jsp",
        "COES gas-fired generation (cross-check / daily estimate): "
        "https://www.coes.org.pe/Portal/portalinformacion/generacion",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "SECTORS", "COVERAGE", "SOURCE"})
    chart = (sec[SECTORS] * MCF_TO_M3).rename(columns={"Vehicle_CNG": "Vehicle CNG",
                                                       "Residential_commercial": "Residential & commercial"})
    xlsx_charts.add_chart_sheet(args.out, chart, "Peru gas demand by sector", "million m3/day", kind="stacked_bar")
    out(f"Saved {args.out}")


if __name__ == "__main__":
    main()
