"""
South America coal production (and exports where published), from 2021.

Colombia (about 55-65 Mt/yr, mostly exported thermal coal) - quarterly by department, exports monthly:
  Production - Agencia Nacional de Mineria (ANM) open data on datos.gov.co, dataset r85m-vv6c
  "ANM Volumen de Explotacion de Minerales Asociados a Pagos de Regalias":
    https://www.datos.gov.co/d/r85m-vv6c
  tonnes of coal produced (volumen de explotacion) declared for royalty liquidation, by municipality, department,
  mining title / project and mineral (CARBON, CARBON TERMICO, CARBON METALURGICO, CARBON ANTRACITA), per
  liquidation period 1-12 (the month; small producers liquidate a quarter at a time, on the quarter's last month).
  Hence the series is published by quarter: a quarter appears once its last month has been liquidated. ANM marks
  the figures preliminary (late declarations are added), so each run re-pulls the last REFRESH_YEARS years.
  Volumes are on REGALIAS rows; COMPENSACION rows repeat payments with no tonnage. Some large-mine rows carry
  royalties with zero tonnage (e.g. "CERREJON CONTRATO DE ASOCIACION"): those tonnes are declared on the mine's
  other rows (CZN-CEMT, Oreganal, Patilla, CDC), so totals are not double counted.
  Exports - DANE (from DIAN customs records), monthly series of traditional exports, "Carbon" in tonnes and
  thousand USD FOB (file anex-EXPORTACIONES-SerieCafeCarbonPetroleoNotradicionales-<mes><yyyy>.xlsx):
    https://www.dane.gov.co/index.php/estadisticas-por-tema/comercio-internacional/exportaciones
  Cross-check: UN Comtrade (Colombia's monthly customs returns), HS 2701 coal (excludes coke), net weight.

Other countries - annual (their coal output is small):
  Brazil - EPE, Balanco Energetico Nacional, historical series chapter 2: tables 2.4 steam coal and 2.5
    metallurgical coal, row "Producao", thousand tonnes.
  Argentina - Secretaria de Energia, Balance Energetico Nacional (one xlsx per year): "Carbon Mineral" production
    (YCRT Rio Turbio) in thousand toe, converted to tonnes at the BEN's own net calorific value for domestic coal
    (4,022 kcal/kg, BEN 2021 synthesis table of calorific values): t = ktoe x 10^10 / 4,022 / 1000. Labelled
    as a conversion. www.energia.gob.ar serves an incomplete TLS certificate chain, so its files are fetched over
    https first and over the plain-http address the ministry's own data portal lists only if https fails
    (certificate verification is never switched off).
  Chile - Cochilco, Anuario de Estadisticas del Cobre y Otros Minerales (database xlsx): coal production, t.
  Peru - MINEM, Anuario Minero: coal (carbon antracita / bituminoso) production, t.
  Venezuela - no official production statistics are published; Energy Institute figures only, labelled as such.
Annual cross-check for all countries: Energy Institute Statistical Review of World Energy, "Coal Production - mt".
  energyinst.org refuses scripted downloads (HTTP 403), so EI's all-data xlsx is read from the copy kept in Our
  World in Data's ETL snapshot store, located through the snapshot's .dvc file in github.com/owid/etl (EI's own
  file, unchanged; md5 checked by OWID).

Incremental: the ANM rows, DANE / Comtrade months and annual figures already in the workbook are kept; each run
re-pulls only the last REFRESH_YEARS years of ANM rows, reads DANE's single series file (refreshing the last
REFRESH_MONTHS months, which DANE revises, and adding new ones), asks Comtrade only for months it has not yet
published, and reads the annual sources only for years still missing (EI only when a new edition appears).
Not reachable from the editing sandbox; runs in GitHub Actions (.github/workflows/south_america_coal.yml).
Found via discovery_archive/south_america/SA_COAL_DISCOVERY*.py.
"""
import argparse
import datetime
import html
import io
import os
import re
import sys
import time
import unicodedata
from urllib.parse import quote

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root
import xlsx_notes  # noqa: E402

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "es,en;q=0.8"}
T = (15, 180)
START_YEAR = 2021
REFRESH_YEARS = 2
REFRESH_MONTHS = 6
COMTRADE_MAX_CALLS = 40

ANM_DATASET = "r85m-vv6c"
ANM_API = f"https://www.datos.gov.co/resource/{ANM_DATASET}.json"
ANM_PAGE = f"https://www.datos.gov.co/d/{ANM_DATASET}"
DANE_PAGE = "https://www.dane.gov.co/index.php/estadisticas-por-tema/comercio-internacional/exportaciones"
DANE_FILE = ("https://www.dane.gov.co/files/operaciones/EXPORTACIONES/"
             "anex-EXPORTACIONES-SerieCafeCarbonPetroleoNotradicionales-{m}{y}.xlsx")
COMTRADE = "https://comtradeapi.un.org/public/v1/preview/C/M/HS"
EPE_PAGE = "https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes/BEN-Series-Historicas-Completas"
AR_PAGE = "https://www.argentina.gob.ar/economia/energia/planeamiento-energetico/balances-energeticos"
AR_KCAL_PER_KG = 4022.0
COCHILCO_PAGE = "https://www.cochilco.cl/web/anuario-de-estadisticas-del-cobre-y-otros-minerales/"
PE_ANUARIO = "https://www.gob.pe/institucion/minem/colecciones/2400-anuario-minero"
OWID_API = "https://api.github.com/repos/owid/etl/contents/snapshots/energy_institute"
OWID_DVC = ("https://raw.githubusercontent.com/owid/etl/master/snapshots/energy_institute/{v}/"
            "statistical_review_of_world_energy.xlsx.dvc")
OWID_FILE = "https://snapshots.owid.io/{a}/{b}"
EI_PAGE = "https://www.energyinst.org/statistical-review/resources-and-data-downloads"
MES3 = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]

# Departments kept as their own columns (the rest are "Other departments")
DEPARTMENTS = ["La Guajira", "Cesar", "Boyaca", "Cundinamarca", "Norte de Santander", "Cordoba", "Antioquia",
               "Santander"]
COUNTRIES = ["Colombia", "Brazil", "Argentina", "Chile", "Peru", "Venezuela"]
EI_ROWS = {"Colombia": "Colombia", "Brazil": "Brazil", "Venezuela": "Venezuela",
           "Other S. & Cent. America": "Other S. & Cent. America", "Total S. & Cent. America": "Total S. & Cent. America"}

S_CO, S_CO_EXP, S_CO_MINE, S_CO_RAW = "Colombia", "Colombia exports", "Colombia by mine", "Colombia ANM raw"
S_BR, S_AR, S_CL, S_PE, S_VE = "Brazil", "Argentina", "Chile", "Peru", "Venezuela"
S_ANNUAL, S_EI, S_CHECK = "Annual by country", "EI Statistical Review", "EI cross-check"


def out(*a):
    print(*a, flush=True)


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return " ".join(s.split())


def get(url, tries=3, quiet404=False, **kw):
    for attempt in range(tries):
        try:
            r = requests.get(url, headers=H, timeout=T, **kw)
            if r.status_code == 200:
                return r
            if not (quiet404 and r.status_code == 404):
                out(f"  {url[:170]}: HTTP {r.status_code}")
            if r.status_code in (401, 403, 404, 410):
                return None
        except requests.exceptions.SSLError as e:
            out(f"  {url[:170]}: SSLError {str(e)[:120]}")
            return None
        except requests.RequestException as e:
            out(f"  {url[:170]}: {type(e).__name__}: {str(e)[:150]}")
        time.sleep(5 * (attempt + 1))
    return None


def get_https_first(url):
    """https, then the plain-http address for hosts with a broken certificate chain (never unverified TLS)."""
    https = re.sub(r"^http://", "https://", url)
    r = get(https, tries=2)
    if r is None and "energia.gob.ar" in url:
        r = get(re.sub(r"^https://", "http://", url), tries=2)
    return r


def load_sheet(path, sheet, index_col=0):
    try:
        df = pd.read_excel(path, sheet_name=sheet, index_col=index_col)
        return df.drop(columns=[c for c in df.columns if str(c).startswith("Unnamed")])
    except (FileNotFoundError, ValueError, KeyError):
        return pd.DataFrame()


def tonnes(s):
    """ANM numbers are text: '1,234,567.89', '- 0', '-' -> float."""
    s = str(s).strip()
    if s in ("", "-", "- 0", "nan", "None"):
        return 0.0
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return float("nan")


# ------------------------------------------------------------------ Colombia production (ANM)

ANM_COLS = {"a_o_liquidado": "Year", "periodo_liquidado2": "Period", "nombre_departamento": "Department",
            "nombre_municipio": "Municipality", "codigo_dane": "DANE_code", "id_nombre_del_proyecto2": "Title_id",
            "nombre_del_proyecto": "Project", "recurso_natural": "Mineral", "contraprestacion": "Payment",
            "volumenes_de_explotacion": "Volume_t", "regalias_pagadas": "Royalties_COP"}


def anm_rows(from_year):
    """All coal rows liquidated in from_year or later (None if the API fails)."""
    rows, off = [], 0
    where = ("recurso_natural like 'CARBON%' AND recurso_natural != 'CARBONATO DE CALCIO' "
             f"AND a_o_liquidado >= '{from_year}'")
    while True:
        r = get(ANM_API, params={"$where": where, "$limit": 50000, "$offset": off, "$order": ":id"})
        if r is None:
            return None
        d = r.json()
        rows += d
        if len(d) < 50000:
            break
        off += 50000
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    for c in ANM_COLS:
        if c not in df:
            df[c] = None
    df = df[list(ANM_COLS)].rename(columns=ANM_COLS)
    df["Volume_t"] = df["Volume_t"].map(tonnes)
    df["Royalties_COP"] = df["Royalties_COP"].map(tonnes)
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce")
    df["Period"] = pd.to_numeric(df["Period"], errors="coerce")
    df = df.dropna(subset=["Year", "Period"])
    df[["Year", "Period"]] = df[["Year", "Period"]].astype(int)
    df["Department"] = df["Department"].map(norm)
    df["Project"] = df["Project"].map(lambda s: " ".join(str(s).split()))
    return df.sort_values(["Year", "Period", "Department", "Municipality", "Project"]).reset_index(drop=True)


def royalty_rows(raw):
    return raw[raw["Payment"].astype(str).str.upper().str.startswith("REGAL") & raw["Volume_t"].notna()].copy()


def colombia_quarterly(raw):
    """ANM rows -> quarterly Mt by department; a quarter is kept once its last month has been liquidated."""
    reg = royalty_rows(raw)
    reg["Quarter"] = [pd.Timestamp(y, 3 * ((p - 1) // 3) + 1, 1) for y, p in zip(reg["Year"], reg["Period"])]
    dept = reg["Department"].where(reg["Department"].isin(DEPARTMENTS), "Other departments")
    q = reg.groupby(["Quarter", dept])["Volume_t"].sum().unstack().fillna(0.0) / 1e6
    q = q[[d for d in DEPARTMENTS + ["Other departments"] if d in q.columns]]
    q.columns = [f"{c.replace(' ', '_')}_Mt" for c in q.columns]
    q["Total_production_Mt"] = q.sum(axis=1)
    ends = {(y, p) for y, p in zip(reg["Year"], reg["Period"])}
    done = [i for i in q.index if (i.year, i.month + 2) in ends]
    q = q.loc[done]
    q.index.name = "Quarter"
    return q


def colombia_annual(raw):
    """Calendar-year totals (Mt) for years whose December has been liquidated."""
    reg = royalty_rows(raw)
    years = sorted({y for y, p in zip(reg["Year"], reg["Period"]) if p == 12})
    return (reg[reg["Year"].isin(years)].groupby("Year")["Volume_t"].sum() / 1e6).rename("Colombia")


def colombia_by_mine(raw):
    """Annual Mt by department and project (mines over 0.2 Mt in any year; the rest grouped)."""
    reg = royalty_rows(raw)
    p = reg.groupby(["Department", "Project", "Year"])["Volume_t"].sum().unstack().fillna(0.0) / 1e6
    big = p[p.max(axis=1) >= 0.2]
    rest = p.drop(big.index)
    if not rest.empty:
        small = rest.groupby(level="Department").sum()
        small.index = pd.MultiIndex.from_tuples([(d, "Other titles / small producers") for d in small.index],
                                                names=["Department", "Project"])
        big = pd.concat([big, small])
    big = big.sort_index()
    big.columns = [f"{c}_Mt" for c in big.columns]
    return big.reset_index().set_index("Department")


# ------------------------------------------------------------------ Colombia exports (DANE, Comtrade)

def dane_series():
    """Monthly coal exports from DANE's traditional-exports series file (newest one linked on the exports page,
    else the last few months' file names)."""
    urls = []
    r = get(DANE_PAGE)
    if r is not None:
        urls = re.findall(r'href="([^"]*SerieCafeCarbonPetroleoNotradicionales[^"]*\.xlsx)"', r.text)
        urls = [("https://www.dane.gov.co" + u) if u.startswith("/") else u for u in urls]
    today = pd.Timestamp(datetime.date.today())
    for k in range(1, 6):
        d = today - pd.DateOffset(months=k)
        urls.append(DANE_FILE.format(m=MES3[d.month - 1], y=d.year))
    for u in dict.fromkeys(urls):
        r = get(u, tries=1, quiet404=True)
        if r is None:
            continue
        df = pd.read_excel(io.BytesIO(r.content), header=None)
        hdr = next(i for i in range(40) if any(norm(v).lower() == "carbon" for v in df.iloc[i].astype(str)))
        col = next(j for j, v in enumerate(df.iloc[hdr].astype(str)) if norm(v).lower() == "carbon")
        units = [norm(v).lower() for v in df.iloc[hdr + 1].astype(str)]
        t_col = next(j for j in range(col, col + 3) if "tonelada" in units[j])
        usd_col = next((j for j in range(col, col + 3) if "dolares" in units[j]), None)
        dates = pd.to_datetime(df.iloc[:, 0], errors="coerce")
        ok = dates.notna() & (dates.dt.day == 1)
        s = pd.DataFrame({"DANE_exports_Mt": pd.to_numeric(df.loc[ok, t_col], errors="coerce").values / 1e6,
                          "DANE_exports_USD_million_FOB": (pd.to_numeric(df.loc[ok, usd_col], errors="coerce").values
                                                           / 1e3 if usd_col is not None else float("nan"))},
                         index=pd.DatetimeIndex(dates[ok].values, name="Month"))
        s = s[s.index >= pd.Timestamp(START_YEAR, 1, 1)].dropna(how="all")
        out(f"DANE coal exports: {u} -> {s.index.min():%Y-%m}..{s.index.max():%Y-%m}")
        return s, u
    return pd.DataFrame(), None


def comtrade_month(m):
    """Colombia HS 2701 exports (Mt) for month m; None if Comtrade has not published it; NaN on failure."""
    params = {"reporterCode": 170, "period": m.strftime("%Y%m"), "cmdCode": "2701", "flowCode": "X"}
    for attempt in range(3):
        try:
            r = requests.get(COMTRADE, params=params, headers=H, timeout=(10, 60))
            if r.status_code == 200:
                rows = [d for d in r.json().get("data", []) if d.get("partnerCode") == 0]
                return float(rows[0].get("netWgt") or 0) / 1e9 if rows else None
            out(f"  comtrade {m:%Y-%m}: HTTP {r.status_code}")
        except requests.RequestException as e:
            out(f"  comtrade {m:%Y-%m}: {type(e).__name__}")
        time.sleep(5 * (attempt + 1))
    return float("nan")


def colombia_exports(old, dane):
    """Merge DANE (new months + last REFRESH_MONTHS refreshed) and fill Comtrade gaps."""
    ex = old.copy() if not old.empty else pd.DataFrame()
    if not ex.empty:
        ex.index = pd.to_datetime(ex.index)
    if not dane.empty:
        cut = dane.index.max() - pd.DateOffset(months=REFRESH_MONTHS - 1)
        new = dane if ex.empty else dane[(dane.index >= cut) | ~dane.index.isin(ex.index)]
        ex = new if ex.empty else ex.reindex(ex.index.union(new.index))
        for c in new.columns:
            ex.loc[new.index, c] = new[c]
    if ex.empty:
        return ex
    for c in ("Comtrade_HS2701_Mt", "Comtrade_status"):
        if c not in ex:
            ex[c] = float("nan") if c.endswith("Mt") else ""
    ex["Comtrade_status"] = ex["Comtrade_status"].astype(object).where(ex["Comtrade_status"].notna(), "")
    todo = [m for m in ex.index if ex.at[m, "Comtrade_status"] != "published"][-COMTRADE_MAX_CALLS:]
    for m in todo:
        v = comtrade_month(m)
        if v is None:
            ex.at[m, "Comtrade_status"] = "not yet published"
        elif pd.isna(v):
            ex.at[m, "Comtrade_status"] = "request failed"
        else:
            ex.at[m, "Comtrade_HS2701_Mt"] = v
            ex.at[m, "Comtrade_status"] = "published"
        time.sleep(1.5)
    out(f"Comtrade: asked {len(todo)} months; {int((ex['Comtrade_status'] == 'published').sum())} published in total")
    ex.index.name = "Month"
    return ex.sort_index()


# ------------------------------------------------------------------ Brazil (EPE BEN)

def brazil_epe():
    """EPE BEN chapter 2: steam and metallurgical coal production, thousand t, by year."""
    r = get(EPE_PAGE)
    if r is None:
        return pd.DataFrame(), None
    links = [html.unescape(h) for h in re.findall(r'href="([^"]+\.xlsx)"', r.text)]
    links = [h for h in links if re.search(r"Cap.tulo 2 ", h)]
    if not links:
        out("EPE: chapter 2 link not found")
        return pd.DataFrame(), None
    url = links[0] if links[0].startswith("http") else "https://www.epe.gov.br" + links[0]
    url = quote(url, safe=":/")
    x = get(url)
    if x is None:
        return pd.DataFrame(), None
    df = pd.read_excel(io.BytesIO(x.content), header=None)
    labels = [norm(v).upper() for v in df.iloc[:, 0].astype(str)]

    def years_above(i):
        for k in range(i, -1, -1):
            row = df.iloc[k]
            yrs = {j: int(v) for j, v in enumerate(row.values) if isinstance(v, (int, float)) and not pd.isna(v)
                   and 1970 <= v <= 2100 and float(v).is_integer()}
            if len(yrs) > 20:
                return yrs
        return {}

    res = {}
    for key, col in (("CARVAO VAPOR", "Steam_coal_kt"), ("CARVAO METALURGICO", "Metallurgical_coal_kt")):
        t = next((i for i, s in enumerate(labels) if s.startswith(key) and "UNIDADE" in " ".join(
            norm(v).upper() for v in df.iloc[i].astype(str))), None)
        if t is None:
            out(f"EPE: table {key} not found")
            continue
        p = next((i for i in range(t + 1, t + 30) if labels[i].startswith("PRODUC")), None)
        yrs = years_above(p) if p is not None else {}
        if p is None or not yrs:
            out(f"EPE: production row for {key} not found")
            continue
        res[col] = {y: pd.to_numeric(df.iat[p, j], errors="coerce") for j, y in yrs.items()}
    if not res:
        return pd.DataFrame(), None
    b = pd.DataFrame(res)
    b.index.name = "Year"
    b = b[b.index >= START_YEAR]
    b["Total_Mt"] = b.sum(axis=1, min_count=1) / 1000
    out(f"EPE BEN: Brazil coal production {b.index.min()}..{b.index.max()}")
    return b, url


# ------------------------------------------------------------------ Argentina (SE BEN)

def argentina_links():
    r = get(AR_PAGE)
    found = {}
    if r is not None:
        for h in re.findall(r'href="([^"]+\.xlsx)"', r.text):
            h = html.unescape(h)
            m = re.search(r"balance_(\d{4})_v", h, re.I)
            if m and "BEP" not in h:
                y = int(m.group(1))
                found.setdefault(y, ("https://www.argentina.gob.ar" + h) if h.startswith("/") else h)
    return found


def argentina_year(url):
    """'Carbon Mineral' production (ktoe) from one BEN horizontal workbook."""
    r = get_https_first(url)
    if r is None:
        return None
    df = pd.read_excel(io.BytesIO(r.content), header=None)
    cells = df.astype(str).map(lambda v: norm(v).upper())
    hdr = [(i, j) for i in range(min(15, len(df))) for j in range(df.shape[1]) if cells.iat[i, j] == "PRODUCCION"]
    row = [i for i in range(len(df)) if any(cells.iat[i, j] == "CARBON MINERAL" for j in range(min(4, df.shape[1])))]
    if not hdr or not row:
        return None
    return float(pd.to_numeric(df.iat[row[0], hdr[0][1]], errors="coerce"))


def argentina(old):
    have = set(old.index) if not old.empty else set()
    rows = {}
    for y, u in sorted(argentina_links().items()):
        if y < START_YEAR or (y in have and not pd.isna(old.at[y, "Production_ktoe"])):
            continue
        v = argentina_year(u)
        if v is not None:
            rows[y] = {"Production_ktoe": v, "Source_file": u}
            out(f"Argentina BEN {y}: coal production {v} ktoe")
    new = pd.DataFrame.from_dict(rows, orient="index")
    a = pd.concat([old, new]) if not old.empty else new
    if a.empty:
        return a
    a = a[~a.index.duplicated(keep="last")].sort_index()
    a.index.name = "Year"
    a["Production_kt_converted"] = a["Production_ktoe"] * 1e7 / AR_KCAL_PER_KG / 1000
    a["Production_Mt_converted"] = a["Production_kt_converted"] / 1000
    return a[["Production_ktoe", "Production_kt_converted", "Production_Mt_converted", "Source_file"]]


# ------------------------------------------------------------------ Chile (Cochilco), Peru (MINEM)

def chile(old):
    """Cochilco yearbook database: coal production rows (t)."""
    r = get(COCHILCO_PAGE)
    if r is None:
        return old, None
    links = re.findall(r'href="([^"]*base-de-datos-anuario[^"]*\.xlsx)"', r.text)
    if not links:
        return old, None
    x = get(links[0])
    if x is None:
        return old, None
    found = cochilco_coal(x.content)
    if found.empty:
        out("Cochilco: no coal rows found")
        return old, links[0]
    c = found if old.empty else found.combine_first(old)
    return c, links[0]


def cochilco_coal(content):
    """Coal production by year (t) from the Cochilco database workbook: Tabla 1.1 'Produccion minera de Chile'
    (national; the 7.x tables are by region), row 'Carbon (TM netas) / Coal (net MT)'."""
    xl = pd.ExcelFile(io.BytesIO(content))
    order = sorted(xl.sheet_names, key=lambda s: (norm(s).lower() != "tabla 1.1", xl.sheet_names.index(s)))
    for s in order:
        df = xl.parse(s, header=None)
        lab = df.iloc[:, :3].astype(str).map(lambda v: norm(v).lower())
        for i in range(len(df)):
            if not any(re.match(r"carbon \(tm netas\)", lab.iat[i, j]) for j in range(lab.shape[1])):
                continue
            for k in range(i, -1, -1):
                yrs = {j: int(v) for j, v in enumerate(df.iloc[k].values) if isinstance(v, (int, float))
                       and not pd.isna(v) and 1990 <= v <= 2100 and float(v).is_integer()}
                if len(yrs) > 5:
                    vals = {y: pd.to_numeric(df.iat[i, j], errors="coerce") for j, y in yrs.items() if y >= START_YEAR}
                    c = pd.DataFrame({"Production_t": vals})
                    c.index.name = "Year"
                    c["Production_Mt"] = c["Production_t"] / 1e6
                    out(f"Cochilco coal [{s} r{i}]: {c['Production_t'].to_dict()}")
                    return c
    return pd.DataFrame()


def peru_anuarios():
    """{year: publication path} of MINEM's Anuario Minero (gob.pe collection), newest first."""
    found = {}
    for page in (1, 2):
        r = get(f"{PE_ANUARIO}?sheet={page}")
        if r is None:
            break
        for path, y in re.findall(r'href="(/institucion/minem/informes-publicaciones/\d+-anuario-minero-(\d{4}))"', r.text):
            found.setdefault(int(y), path)
    if not found:   # the collection page is rendered client-side at times: fall back to the site search
        r = get("https://www.gob.pe/busquedas.json", params={"term": "anuario minero", "institucion[]": "minem"})
        if r is not None:
            for path, y in re.findall(r'(/institucion/minem/informes-publicaciones/\d+-anuario-minero-(\d{4}))', r.text):
                found.setdefault(int(y), path)
    return dict(sorted(found.items(), reverse=True))


def peru_annex(content):
    """Coal rows of the Anuario's Excel annex ('Produccion' sheet): t by year for anthracite and bituminous."""
    xl = pd.ExcelFile(io.BytesIO(content))
    for s in [s for s in xl.sheet_names if norm(s).lower().strip().startswith("produccion")
              and "mundial" not in norm(s).lower()]:
        df = xl.parse(s, header=None)
        lab = df.astype(str).map(lambda v: norm(v).lower())
        res = {}
        for i in range(len(df)):
            j0 = next((j for j in range(min(6, df.shape[1])) if re.match(r"carbon (antracita|bituminoso)", lab.iat[i, j])), None)
            if j0 is None:
                continue
            kind = "Anthracite_t" if "antracita" in lab.iat[i, j0] else "Bituminous_t"
            for k in range(i, -1, -1):
                yrs = {j: int(v) for j, v in enumerate(df.iloc[k].values) if isinstance(v, (int, float))
                       and not pd.isna(v) and 1990 <= v <= 2100 and float(v).is_integer() and j > j0}
                if len(yrs) >= 2:
                    res.setdefault(kind, {y: pd.to_numeric(df.iat[i, j], errors="coerce") for j, y in yrs.items()})
                    out(f"Peru annex [{s} r{i}] {kind}: {res[kind]}")
                    break
        if res:
            p = pd.DataFrame(res)
            p.index.name = "Year"
            return p
    return pd.DataFrame()


def peru(old):
    """MINEM Anuario Minero Excel annex: coal production (t) by year; older anuarios only for missing years."""
    have = set(old.index) if not old.empty else set()
    want = set(range(START_YEAR, datetime.date.today().year))
    if want <= have:
        return old, None
    p, src = old.copy(), None
    for y, path in peru_anuarios().items():
        if y < START_YEAR or not (want - have):
            break
        r = get("https://www.gob.pe" + path)
        if r is None:
            continue
        files = re.findall(r'https://cdn\.www\.gob\.pe/uploads/document/file/\d+/[^"?#\s]+\.xlsx', r.text)
        for f in dict.fromkeys(files):
            x = get(f)
            if x is None:
                continue
            got = peru_annex(x.content)
            if got.empty:
                continue
            got = got[got.index >= START_YEAR]
            got["Source_file"] = f
            p = got if p.empty else got.combine_first(p)
            have |= set(got.index)
            src = src or f
            break
    if p.empty:
        return p, src
    p = p.sort_index()
    p["Production_t"] = p[[c for c in ("Anthracite_t", "Bituminous_t") if c in p]].sum(axis=1, min_count=1)
    p["Production_Mt"] = p["Production_t"] / 1e6
    cols = [c for c in ("Anthracite_t", "Bituminous_t", "Production_t", "Production_Mt", "Source_file") if c in p]
    return p[cols], src


# ------------------------------------------------------------------ Energy Institute Statistical Review

def ei_latest_version():
    hdr = dict(H)
    if os.environ.get("GITHUB_TOKEN"):
        hdr["Authorization"] = f"Bearer {os.environ['GITHUB_TOKEN']}"
    try:
        r = requests.get(OWID_API, headers=hdr, timeout=T)
        if r.status_code == 200:
            vers = sorted(d["name"] for d in r.json()
                          if d["type"] == "dir" and re.match(r"\d{4}-\d{2}-\d{2}$", d["name"]))
            return vers[-1] if vers else None
        out(f"  OWID snapshot list: HTTP {r.status_code}")
    except requests.RequestException as e:
        out(f"  OWID snapshot list: {type(e).__name__}")
    return None


def ei_production(version):
    """EI 'Coal Production - mt' rows for South & Central America (Mt by year, 2021 on) from the all-data xlsx
    of the given OWID snapshot version (= EI publication date)."""
    r = get(OWID_DVC.format(v=version))
    if r is None:
        return pd.DataFrame()
    md5 = re.search(r"md5:\s*([0-9a-f]{32})", r.text)
    if not md5:
        return pd.DataFrame()
    x = get(OWID_FILE.format(a=md5.group(1)[:2], b=md5.group(1)[2:]))
    if x is None:
        return pd.DataFrame()
    df = pd.read_excel(io.BytesIO(x.content), sheet_name="Coal Production - mt", header=None)
    hdr = next(i for i in range(10) if str(df.iloc[i, 0]).strip().lower().startswith("million tonnes"))
    years = {j: int(v) for j, v in enumerate(df.iloc[hdr]) if isinstance(v, (int, float)) and not pd.isna(v)
             and 1900 < v < 2100 and float(v).is_integer()}
    rows = {}
    for _, row in df.iterrows():
        name = str(row.iloc[0]).strip()
        if name in EI_ROWS:
            rows[EI_ROWS[name]] = {y: pd.to_numeric(row.iloc[j], errors="coerce") for j, y in years.items()}
    ei = pd.DataFrame(rows)
    ei.index.name = "Year"
    ei = ei[ei.index >= START_YEAR]
    ei["Edition"] = f"EI Statistical Review {version[:4]} (published {version})"
    out(f"EI Statistical Review {version}: {list(rows)} {ei.index.min()}..{ei.index.max()}")
    return ei


# ------------------------------------------------------------------ assembly

def annual_table(co_annual, br, ar, cl, pe, ei):
    """Mt/yr by country with the source of every cell."""
    years = range(START_YEAR, datetime.date.today().year + 1)
    t = pd.DataFrame(index=pd.Index(list(years), name="Year"))

    def put(country, series, source):
        t[f"{country}_Mt"] = series.reindex(t.index) if series is not None else float("nan")
        t[f"{country}_source"] = [source if not pd.isna(v) else "" for v in t[f"{country}_Mt"]]

    put("Colombia", co_annual, "ANM royalty production volumes (datos.gov.co r85m-vv6c)")
    put("Brazil", br["Total_Mt"] if not br.empty else None, "EPE Balanco Energetico Nacional (steam + metallurgical)")
    put("Argentina", ar["Production_Mt_converted"] if not ar.empty else None,
        "SE Balance Energetico Nacional (ktoe converted at 4,022 kcal/kg)")
    put("Chile", cl["Production_Mt"] if not cl.empty else None, "Cochilco yearbook")
    put("Peru", pe["Production_Mt"] if not pe.empty else None, "MINEM Anuario Minero")
    ve = ei["Venezuela"] if not ei.empty and "Venezuela" in ei else None
    put("Venezuela", ve, "Energy Institute Statistical Review (no official national data)")
    t["Total_Mt"] = t[[f"{c}_Mt" for c in COUNTRIES]].sum(axis=1, min_count=1)
    t["Countries_reported"] = t[[f"{c}_Mt" for c in COUNTRIES]].notna().sum(axis=1)
    return t.dropna(subset=["Colombia_Mt", "Brazil_Mt"], how="all")


def ei_check(annual, ei):
    if ei.empty:
        return pd.DataFrame()
    c = pd.DataFrame(index=ei.index)
    for country in ("Colombia", "Brazil"):
        c[f"{country}_ours_Mt"] = annual[f"{country}_Mt"].reindex(c.index)
        c[f"{country}_EI_Mt"] = ei.get(country)
        c[f"{country}_diff_pct"] = (c[f"{country}_ours_Mt"] / c[f"{country}_EI_Mt"] - 1) * 100
    c["Venezuela_EI_Mt"] = ei.get("Venezuela")   # the annual table uses EI for Venezuela: nothing to compare
    small = annual[["Argentina_Mt", "Chile_Mt", "Peru_Mt"]].reindex(c.index)
    c["AR_CL_PE_ours_Mt"] = small.sum(axis=1, min_count=1)
    c["Other_SCA_EI_Mt"] = ei.get("Other S. & Cent. America")
    c["Total_ours_Mt"] = annual["Total_Mt"].reindex(c.index)
    c["Total_SCA_EI_Mt"] = ei.get("Total S. & Cent. America")
    c["Edition"] = ei.get("Edition")
    return c


def span(df):
    return f"{df.index.min():%b %Y}..{df.index.max():%b %Y}" if not df.empty else "none"


def yspan(df):
    return f"{df.index.min()}..{df.index.max()}" if not df.empty else "none"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output/Data and Chart Outputs/south_america_coal_production.xlsx")
    args = ap.parse_args()
    sources = {}

    # Colombia production: keep archived ANM rows, re-pull the last REFRESH_YEARS years
    raw = load_sheet(args.out, S_CO_RAW, index_col=None)
    from_year = START_YEAR if raw.empty else max(START_YEAR, int(raw["Year"].max()) - REFRESH_YEARS + 1)
    new = anm_rows(from_year)
    if new is None or new.empty:
        out(f"ANM: no rows fetched from {from_year} (archive kept)")
    else:
        out(f"ANM: {len(new)} coal rows liquidated {from_year}..{new['Year'].max()}")
        raw = pd.concat([raw[raw["Year"] < from_year], new]) if not raw.empty else new
    if raw.empty:
        raise SystemExit("no ANM data fetched and no existing archive")
    for c in ("Year", "Period"):
        raw[c] = raw[c].astype(int)
    q = colombia_quarterly(raw)
    co_annual = colombia_annual(raw)
    mines = colombia_by_mine(raw)

    # Colombia exports
    dane, dane_url = dane_series()
    sources["dane"] = dane_url
    exports = colombia_exports(load_sheet(args.out, S_CO_EXP), dane)
    if not exports.empty:
        ex = exports["DANE_exports_Mt"]
        qx = ex.groupby(ex.index.to_period("Q")).agg(["sum", "count"])
        qx = qx[qx["count"] == 3]["sum"]
        qx.index = qx.index.to_timestamp()
        q["Exports_DANE_Mt"] = qx.reindex(q.index)

    # Annual sources
    br_old = load_sheet(args.out, S_BR)
    last_full = datetime.date.today().year - 1
    br = br_old
    if br_old.empty or last_full not in br_old.index:
        b, sources["epe"] = brazil_epe()
        br = b if not b.empty else br_old
    ar = argentina(load_sheet(args.out, S_AR))
    cl_old = load_sheet(args.out, S_CL)
    cl = cl_old
    if cl_old.empty or last_full not in cl_old.index:
        cl, sources["cochilco"] = chile(cl_old)
    pe, sources["minem"] = peru(load_sheet(args.out, S_PE))

    ei_old = load_sheet(args.out, S_EI)
    version = ei_latest_version()
    ei = ei_old
    if version and (ei_old.empty or not ei_old["Edition"].astype(str).str.contains(version).any()):
        e = ei_production(version)
        ei = e if not e.empty else ei_old
    ve = pd.DataFrame({"Production_Mt": ei["Venezuela"]}) if not ei.empty and "Venezuela" in ei else pd.DataFrame()
    if not ve.empty:
        ve["Source"] = ei["Edition"]

    annual = annual_table(co_annual, br, ar, cl, pe, ei)
    check = ei_check(annual, ei)

    sheets = {S_CO: q, S_CO_EXP: exports, S_CO_MINE: mines, S_BR: br, S_AR: ar, S_CL: cl, S_PE: pe, S_VE: ve,
              S_ANNUAL: annual, S_EI: ei, S_CHECK: check, S_CO_RAW: raw.set_index("Year")}
    sheets = {k: v for k, v in sheets.items() if v is not None and not v.empty}

    last_q = q.index.max() if not q.empty else None
    notes = [
        "South America coal production (and Colombia exports), from 2021",
        "",
        "UNITS",
        "Mt = million tonnes; kt = thousand tonnes; ktoe = thousand tonnes of oil equivalent (10^7 kcal per toe).",
        f"'{S_CO}': coal production per calendar quarter (Mt) by producing department, from ANM royalty "
        "declarations (sum of CARBON, CARBON TERMICO, CARBON METALURGICO, CARBON ANTRACITA; REGALIAS rows), plus "
        "Exports_DANE_Mt = DANE coal exports summed over the quarter's three months.",
        f"'{S_CO_EXP}': monthly exports. DANE_exports_Mt / DANE_exports_USD_million_FOB = DANE traditional "
        "exports 'Carbon' (coal; DANE's group also takes in coke and briquettes, CUCI 32). Comtrade_HS2701_Mt = "
        "UN Comtrade HS 2701 (coal only) net weight, Colombia as reporter - cross-check.",
        f"'{S_CO_MINE}': annual Mt by department and mining title / project (titles above 0.2 Mt in any year; "
        "the rest per department as 'Other titles / small producers'); for cross-checks with company reports "
        "(Cerrejon = the CERREJON / CDC / CARBONES DEL CERREJON rows in La Guajira; Drummond = El Descanso, La "
        "Loma, El Corozo in Cesar).",
        f"'{S_ANNUAL}': Mt per calendar year by country, with the source of each figure beside it; Colombia only "
        "for years whose December has been liquidated.",
        f"'{S_BR}': EPE production of steam coal (carvao vapor) and metallurgical coal, kt; Total_Mt = sum.",
        f"'{S_AR}': BEN 'Carbon Mineral' production in ktoe as published; Production_kt_converted / _Mt_converted "
        f"= ktoe x 10^7 kcal / {AR_KCAL_PER_KG:,.0f} kcal/kg (the BEN's net calorific value for domestic coal) - a "
        "conversion, not a published tonnage.",
        f"'{S_VE}': Energy Institute figure (Mt) - Venezuela publishes no official coal production statistics.",
        f"'{S_EI}': EI 'Coal Production - mt' rows for South & Central America; '{S_CHECK}': our figures against "
        "EI (diff_pct = ours / EI - 1). EI lists Colombia, Brazil and Venezuela; Argentina, Chile and Peru are "
        "inside 'Other S. & Cent. America'.",
        f"'{S_CO_RAW}': every ANM coal row used (tonnes, royalties in COP), the incremental archive.",
        "",
        "COVERAGE",
        f"Colombia production quarterly {span(q)} (latest quarter {last_q:%b %Y} onwards is preliminary; ANM "
        f"revises as late declarations arrive - the last {REFRESH_YEARS} years are re-pulled each run)." if last_q
        is not None else "Colombia production: none",
        f"Colombia exports monthly {span(exports)} (DANE publishes about 6 weeks after month end; Comtrade later).",
        f"Brazil {yspan(br)}; Argentina {yspan(ar)}; Chile {yspan(cl)}; Peru {yspan(pe)}; Venezuela (EI) {yspan(ve)}.",
        "Chile: Mina Invierno (the last large mine) stopped in 2020; remaining output is negligible. Peru: small "
        "anthracite / bituminous output (Ancash, La Libertad). Annual figures only for both, where published.",
        "Monthly production by department is not published: ANM's royalty data are monthly only for the large "
        "mines that liquidate monthly; small producers liquidate by quarter.",
        "",
        "SOURCE",
        f"ANM, Volumen de Explotacion de Minerales Asociados a Pagos de Regalias (datos.gov.co): {ANM_PAGE}",
        f"DANE exports (series file): {sources.get('dane') or DANE_PAGE}",
        f"UN Comtrade public API (HS 2701, reporter Colombia): {COMTRADE}",
        f"EPE, Balanco Energetico Nacional - Series Historicas Completas: {sources.get('epe') or EPE_PAGE}",
        f"Secretaria de Energia (Argentina), Balances Energeticos: {AR_PAGE}",
        f"Cochilco, Anuario de Estadisticas del Cobre y Otros Minerales: {sources.get('cochilco') or COCHILCO_PAGE}",
        f"MINEM Peru, Anuario Minero: {sources.get('minem') or PE_ANUARIO}",
        f"Energy Institute, Statistical Review of World Energy: {EI_PAGE} (all-data xlsx read from OWID's snapshot "
        "store, https://github.com/owid/etl/tree/master/snapshots/energy_institute)",
    ]
    titles = {"UNITS", "COVERAGE", "SOURCE"}
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, titles)
    out(f"Saved {args.out}")
    if not q.empty:
        out(f"Colombia latest quarter {q.index.max():%Y-%m}: {q['Total_production_Mt'].iloc[-1]:.2f} Mt")
    out(annual[[c for c in annual.columns if c.endswith('_Mt')]].round(3).to_string())
    if not check.empty:
        out(check.round(2).to_string())


if __name__ == "__main__":
    main()
