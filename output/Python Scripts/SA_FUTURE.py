"""
South America "future" workbook (one-off, run on demand): generation projects in the pipeline.

  Brazil     ANEEL SIGA plant register (the one brazil_power_capacity uses): plants with DscFaseUsina
             "Construção" (under construction) and "Construção não iniciada" (granted, construction not started),
             with granted capacity (MdaPotenciaOutorgadaKw) and any expected-operation date in the register
  Argentina  Secretaría de Energía open data "Obras de generación de electricidad" (generation works):
             http://datos.energia.gob.ar/ ... obras-generacion.csv

Writes south_america_future.xlsx. Other countries: no machine-readable pipeline found yet (Chile's Coordinador
blocks automated access; Colombia UPME and Peru COES publish PDF / web lists).

Usage: python3 future/SA_FUTURE.py [--out "output/Data and Chart Outputs/south_america_future.xlsx"]
"""
import argparse
import io
import os
import re
import sys
from datetime import date

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "south_america"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_notes  # noqa: E402
from future_common import FUELS, by_year_fuel, col, fuel_from_text, get  # noqa: E402

RALIE_CKAN = "https://dadosabertos.aneel.gov.br/api/3/action/package_search"
AR_CKAN = "https://datos.energia.gob.ar/api/3/action/package_search"
AR_OBRAS = ("https://datos.energia.gob.ar/dataset/84bf14b4-47bf-4efb-ad7c-0d70684a50d3/resource/"
            "535ed5fc-ca6d-463e-9f88-2fe33b7b3790/download/obras-generacion.csv")
DEFAULT_OUT = os.path.join(ROOT, "output", "Data and Chart Outputs", "south_america_future.xlsx")
PHASES = {"Construção": "Under construction", "Construção não iniciada": "Granted, construction not started"}


def brazil():
    import BRAZIL_ANEEL_CAPACITY as siga
    from BRAZIL_ANEEL_CAPACITY_MONTHLY import siga_fuel

    raw = siga.fetch_raw()
    print(f"SIGA: {len(raw):,} rows; phases {raw['DscFaseUsina'].value_counts().to_dict()}", flush=True)
    print(f"  date columns: {[c for c in raw.columns if c.startswith('Dat')]}", flush=True)
    p = raw[raw["DscFaseUsina"].isin(PHASES)].copy()
    p["Phase"] = p["DscFaseUsina"].map(PHASES)
    p["MW"] = siga.to_float_br(p["MdaPotenciaOutorgadaKw"]) / 1000.0
    p["Fuel"] = [siga_fuel(t, f) for t, f in zip(p["SigTipoGeracao"], p["DscFonteCombustivel"])]
    sheets = {}
    ph = p.pivot_table(index="Phase", columns="Fuel", values="MW", aggfunc="sum").fillna(0)
    ph = ph.reindex(columns=[f for f in FUELS if f in ph.columns]).round(0)
    ph.index.name = "Phase"
    sheets["BR pipeline by phase"] = ph
    # SIGA's DatEntradaOperacao is blank (read as day 0) for plants not yet operating: expected dates come from RALIE
    keep = [c for c in ("NomEmpreendimento", "SigUFPrincipal", "SigTipoGeracao", "DscFonteCombustivel", "Fuel",
                        "Phase", "MW") if c in p.columns]
    sheets["BR pipeline projects"] = p[keep].sort_values("MW", ascending=False).set_index(keep[0])
    print(f"  pipeline {p['MW'].sum() / 1000:,.1f} GW: {p.groupby('Fuel')['MW'].sum().round(0).to_dict()}", flush=True)
    return sheets


def _ckan_csv(api, query, pattern):
    """URL of the first CSV resource matching `pattern` in a CKAN search."""
    for pkg in get(api, params={"q": query, "rows": 10}).json()["result"]["results"]:
        for res in pkg.get("resources", []):
            u = res.get("url", "")
            if re.search(pattern, u, re.I) and u.lower().endswith(".csv"):
                return u
    raise RuntimeError(f"no CSV matching {pattern} for '{query}'")


def _csv(content):
    for enc in ("utf-8", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(content), encoding=enc, sep=None, engine="python")
        except UnicodeDecodeError:
            continue
    raise ValueError("unreadable CSV")


def brazil_ralie():
    """ANEEL RALIE (generation expansion monitoring): projects under construction / granted with their forecast
    commercial-operation date -> MW by forecast year and fuel."""
    url = _ckan_csv(RALIE_CKAN, "ralie", r"ralie.*usina|usina.*ralie|ralie")
    d = _csv(get(url).content)
    print(f"RALIE {url}: {d.shape}; columns {list(d.columns)}\n{d.head(3).to_string()[:1500]}", flush=True)
    mw = col(d, r"MdaPotenciaOutorgada|Potencia.*Outorgada|Potencia|MW")
    dates = [c for c in d.columns if re.search(r"Prev.*Oper.*Comerc|OperacaoComercial|Prev.*Comerc|Prev", c, re.I)]
    print(f"  date columns: {dates}", flush=True)
    if not dates:
        raise KeyError("no forecast-date column")
    when = pd.to_datetime(d[dates[0]], errors="coerce", dayfirst=True)
    for c in dates[1:]:   # first forecast date given among the candidates
        when = when.fillna(pd.to_datetime(d[c], errors="coerce", dayfirst=True))
    tipo = col(d, r"SigTipoGeracao|TipoGera", required=False)
    fonte = col(d, r"DscFonte|Combust|Fonte", required=False)
    text = (d[tipo].fillna("").astype(str) if tipo else "") + " " + (d[fonte].fillna("").astype(str) if fonte else "")
    from BRAZIL_ANEEL_CAPACITY_MONTHLY import siga_fuel
    fuel = [siga_fuel(t, f) for t, f in zip(d[tipo].fillna(""), d[fonte].fillna(""))] if tipo and fonte else \
        text.map(fuel_from_text)
    kw = pd.to_numeric(d[mw].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
                       errors="coerce")
    t = d.assign(Fuel=fuel, MW=kw / 1000.0 if "kw" in mw.lower() else kw, Year=when.dt.year)
    t = t[t["Year"] >= date.today().year - 1]
    print(f"  {len(t)} dated projects, {t['MW'].sum() / 1000:,.1f} GW: {t.groupby('Year')['MW'].sum().round(0).to_dict()}",
          flush=True)
    return {"BR expected additions": by_year_fuel(t, "Year", "Fuel", "MW", end=2040)}, url


CNE_MEDIA = "https://www.cne.cl/wp-json/wp/v2/media"
ES_MONTHS = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6, "jul": 7, "ago": 8, "sep": 9, "set": 9,
             "oct": 10, "nov": 11, "dic": 12}


def _es_date(v):
    """Excel date, 'nov-25' / 'nov-2025' style text, or blank -> Timestamp."""
    if isinstance(v, (pd.Timestamp, date)):
        return pd.Timestamp(v)
    m = re.match(r"\s*([a-zA-Z]{3})[a-z]*[-/ .]+(\d{2,4})", str(v))
    if m and m.group(1).lower() in ES_MONTHS:
        y = int(m.group(2))
        return pd.Timestamp(year=y + 2000 if y < 100 else y, month=ES_MONTHS[m.group(1).lower()], day=1)
    t = str(v).strip()
    return pd.to_datetime(t, errors="coerce", dayfirst=not re.match(r"\d{4}-", t))


def _num(v):
    """'1,0' / '1.234,5' / 12.3 -> float."""
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace(" ", "")
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return float("nan")


def chile():
    """CNE 'Instalaciones declaradas en construcción' (monthly workbook): generation (PMGD small distributed, PMG /
    large generation, isolated systems) and BESS projects with estimated interconnection date and net MW."""
    import requests
    from future_common import HEADERS
    media = requests.get(CNE_MEDIA, params={"search": "Declaracion-Construccion", "per_page": 30, "orderby": "date"},
                         headers=HEADERS, timeout=(10, 120)).json()
    urls = [m["source_url"] for m in media if re.search(r"Declaracion-Construccion.*\.xlsx$", m.get("source_url", ""), re.I)]
    if not urls:
        raise RuntimeError("no 'Tablas-Declaracion-Construccion' workbook in CNE's media library")
    url = sorted(urls, key=lambda u: re.search(r"/uploads/(\d{4}/\d{2})/", u).group(1) if re.search(r"/uploads/(\d{4}/\d{2})/", u) else "")[-1]
    xl = pd.ExcelFile(io.BytesIO(get(url).content))
    print(f"CNE declared-in-construction {url}: sheets {xl.sheet_names}", flush=True)
    rows = []
    for sh, kind in (("PMGD", "Small distributed (PMGD)"), ("P.Generación", "Generation"), ("BESS", "Storage (BESS)"),
                     ("P.Generación SSMM", "Isolated systems")):
        if sh not in xl.sheet_names:
            continue
        raw = pd.read_excel(xl, sh, header=None)
        h = next(i for i in range(min(len(raw), 15)) if raw.iloc[i].astype(str).str.strip().eq("Proyecto").any())
        d = raw.iloc[h + 1:].copy()
        d.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in raw.iloc[h]]
        d = d[d["Proyecto"].notna()]
        tech = col(d, r"Tipo de Tecnolog")
        mw = col(d, r"Potencia Neta")
        when = col(d, r"Fecha Estimada")
        t = pd.DataFrame({"Project": d["Proyecto"], "Owner": d.get(col(d, r"Propietario", required=False)),
                          "Category": kind, "Technology": d[tech],
                          "Fuel": "Battery storage" if kind.startswith("Storage") else d[tech].fillna("").astype(str).map(fuel_from_text),
                          "MW": d[mw].map(_num), "Expected": d[when].map(_es_date),
                          "Region": d.get(col(d, r"Ubicaci|Regi", required=False))})
        print(f"  {sh}: {len(t)} projects, {t['MW'].sum():,.0f} MW", flush=True)
        rows.append(t)
    p = pd.concat(rows, ignore_index=True)
    p["Year"] = p["Expected"].dt.year
    sheets = {"CL in construction by year": by_year_fuel(p, "Year", "Fuel", "MW"),
              "CL in construction projects": p.sort_values("MW", ascending=False).set_index("Project")}
    cat = p.pivot_table(index="Category", columns="Fuel", values="MW", aggfunc="sum").fillna(0).round(0)
    cat.index.name = "Category"
    sheets["CL in construction by type"] = cat
    print(f"  total {p['MW'].sum() / 1000:,.1f} GW: {p.groupby('Fuel')['MW'].sum().round(0).to_dict()}", flush=True)
    return sheets, url


UPME_REG = ("https://docs.upme.gov.co/SIMEC/Energia%20Electrica/Informes_Registro_Proyectos_Generacion/"
            "Informe_registros_activos_de_proyectos_generacion_electrica_{m}_{y}.xlsx")
ES_MONTH_NAMES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
                  "noviembre", "diciembre"]


def colombia():
    """UPME register of generation projects (active registrations, monthly workbook): phase (FASE 1 prefeasibility,
    2 feasibility, 3 detailed design / construction), resource, technology, MW and expected start of operation."""
    today = date.today()
    content = url = None
    for back in range(0, 6):   # newest published month first
        m = (today.month - 1 - back) % 12
        y = today.year + (today.month - 1 - back) // 12
        u = UPME_REG.format(m=ES_MONTH_NAMES[m], y=y)
        try:
            r = get(u)
        except Exception:  # noqa: BLE001
            continue
        if r.content[:2] == b"PK":
            content, url = r.content, u
            break
    if content is None:
        raise RuntimeError("no UPME active-registrations workbook found")
    raw = pd.read_excel(io.BytesIO(content), header=None)
    h = next(i for i in range(min(len(raw), 15)) if raw.iloc[i].astype(str).str.contains("Codigo Proyecto|Código Proyecto").any())
    d = raw.iloc[h + 1:].copy()
    d.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in raw.iloc[h]]
    d = d[pd.to_numeric(d[col(d, r"C[oó]digo Proyecto")], errors="coerce").notna()]
    text = lambda c: d[col(d, c)].fillna("").astype(str).str.strip()  # noqa: E731
    fuel = (text(r"^Recurso") + " " + text(r"^Tipo$") + " " + text(r"Tecnolog")).map(
        lambda t: "Gas" if re.search("GAS", t) else "Hydro" if re.search("HIDR|AGUA", t) else "Wind" if re.search("E[OÓ]LIC|VIENTO", t)
        else "Solar" if re.search("SOLAR|SOL ", t + " ") else "Bioenergy" if re.search("BIOMASA|BAGAZO", t)
        else "Coal" if re.search("CARB", t) else fuel_from_text(t))
    p = pd.DataFrame({"Project": text(r"Nombre"), "Phase": text(r"^Estado"), "Resource": text(r"^Recurso"),
                      "Technology": text(r"Tecnolog"), "Fuel": fuel.values,
                      "MW": d[col(d, r"Capacidad")].map(_num), "Department": text(r"Departamento"),
                      "Expected operation": d[col(d, r"entrada en operaci")].map(_es_date)})
    p = p[p["Project"].ne("") & p["Phase"].ne("")]   # the sheet ends with a totals row
    p["Year"] = p["Expected operation"].dt.year
    ph = p.pivot_table(index="Phase", columns="Fuel", values="MW", aggfunc="sum").fillna(0).round(0)
    ph.index.name = "Phase"
    print(f"UPME {url}: {len(p)} projects, {p['MW'].sum() / 1000:,.1f} GW; by phase "
          f"{p.groupby('Phase')['MW'].sum().round(0).to_dict()}", flush=True)
    return {"CO registered by phase": ph, "CO registered by year": by_year_fuel(p, "Year", "Fuel", "MW"),
            "CO registered projects": p.sort_values("MW", ascending=False).set_index("Project")}, url


COES_PORTAL = "https://www.coes.org.pe/Portal/"
COES_OC = "Planificación/Nuevos Proyectos/Operación Comercial de unidades o centrales de generación/"
COES_EPO = "Planificación/Nuevos Proyectos/Estudios de Pre Operatividad/1. Modelo Eléctrico del SEIN para la elaboración de EPO/"


def _coes_browse(S, path):
    import html as _html
    r = S.post(COES_PORTAL + "browser/vistadatos", data={"baseDirectory": path, "url": path, "indicador": "",
                                                        "initialLink": "", "orderFolder": ""}, timeout=(10, 120))
    r.raise_for_status()
    out = []
    for m in re.finditer(r"openBlob\('([^']+)',\s*'(\w)'", r.text):
        it = (_html.unescape(m.group(1)), m.group(2))
        if it not in out:
            out.append(it)
    return out


def peru():
    """COES: (1) units/plants granted commercial operation (OC list, 2001 on) -> additions by year since 2021;
    (2) projects in COES's SEIN model for pre-operability studies (EPO model, 10-year horizon, PDF) -> pipeline."""
    import requests
    from urllib.parse import quote
    from future_common import HEADERS
    S = requests.Session()
    S.headers.update(HEADERS)
    sheets, src = {}, []
    oc = [p for p, k in _coes_browse(S, COES_OC + "2. Lista de unidades o centrales de generación con conformidad de OC/")
          if p.lower().endswith((".xlsx", ".xls"))]
    if oc:
        content = S.get(COES_PORTAL + "browser/download?url=" + quote(oc[-1]), timeout=(10, 300)).content
        xl = pd.ExcelFile(io.BytesIO(content))
        print(f"COES OC list {oc[-1]}: sheets {xl.sheet_names}", flush=True)
        raw = pd.read_excel(xl, xl.sheet_names[0], header=None)
        h = next(i for i in range(min(len(raw), 20)) if raw.iloc[i].astype(str).str.contains(r"(?i)potencia|MW").any())
        d = raw.iloc[h + 1:].copy()
        d.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in raw.iloc[h]]
        print(f"  columns {list(d.columns)}\n{d.head(5).to_string(max_colwidth=25)[:1500]}", flush=True)
        mw = col(d, r"Potencia.*MW|MW|Potencia")
        when = col(d, r"Fecha.*(Operaci|OC|inicio)|Fecha")
        tech = col(d, r"Tecnolog|Tipo|Fuente|Combust|Recurso", required=False)
        name = col(d, r"Central|Unidad|Nombre|Proyecto", required=False)
        t = pd.DataFrame({"Plant": d[name] if name else "", "Technology": d[tech] if tech else "",
                          "Fuel": (d[tech].fillna("").astype(str) + " " + (d[name].fillna("").astype(str) if name else "")).map(fuel_from_text) if tech else "Other",
                          "MW": d[mw].map(_num), "Date": d[when].map(_es_date)})
        t = t.dropna(subset=["Date"])
        t["Year"] = t["Date"].dt.year
        sheets["PE commercial operation by year"] = by_year_fuel(t, "Year", "Fuel", "MW", start=2021)
        sheets["PE commercial operation list"] = t[t["Year"] >= 2021].sort_values("Date").set_index("Plant")
        src.append(f"COES, unidades o centrales con conformidad de Operación Comercial: {oc[-1]}")
    epo = [p for p, k in _coes_browse(S, COES_EPO) if re.search(r"Modelo_SEIN_EPO|Lo Nuevo", p, re.I) and
           re.search(r"\.(pdf|pfd)$", p, re.I)]
    print(f"COES EPO model files: {epo}", flush=True)
    rows = []
    for f in epo:
        try:
            import pdfplumber
            content = S.get(COES_PORTAL + "browser/download?url=" + quote(f), timeout=(10, 300)).content
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                print(f"  {f.rsplit('/', 1)[-1]}: {len(pdf.pages)} pages; page 1 text:\n"
                      f"{(pdf.pages[0].extract_text() or '')[:1500]}", flush=True)
                for pg in pdf.pages:
                    for tb in pg.extract_tables() or []:
                        for r in tb:
                            cells = [re.sub(r"\s+", " ", str(c or "")).strip() for c in r]
                            if any(re.search(r"(?i)\bMW\b|^\d+([.,]\d+)?$", c) for c in cells) and \
                                    any(re.search(r"20[2-4]\d", c) for c in cells):
                                rows.append(cells + [f.rsplit("/", 1)[-1]])
        except Exception as e:  # noqa: BLE001
            print(f"  {f}: {type(e).__name__}: {e}", flush=True)
    ref = COES_EPO + "2. Información Referencial para EPOs/"
    try:
        items = _coes_browse(S, ref)
        print(f"COES EPO reference folder: {items[:40]}", flush=True)
        for sub in [p for p, k in items if k.upper() == "D"][:6]:
            print(f"  {sub}: {_coes_browse(S, sub)[:30]}", flush=True)
    except Exception as e:  # noqa: BLE001
        print(f"  EPO reference folder: {type(e).__name__}: {e}", flush=True)
    prefix = [(r"^C\.?\s?E\b|E[OÓ]LIC", "Wind"), (r"^C\.?\s?S\.?\s?F|^C\.?\s?S\b|SOLAR", "Solar"),
              (r"^C\.?\s?H\b|HIDRO", "Hydro"), (r"^C\.?\s?T\b|TERMO|GAS|CICLO", "Gas"), (r"BESS|ALMACEN", "Battery storage")]
    proj = []
    for r in rows:
        name = next((c for c in r if re.search(r"\d+(?:[.,]\d+)?\s*MW", c)), "")
        m = re.search(r"(\d+(?:[.,]\d+)?)\s*MW", name)
        when = next((c for c in r if re.match(r"\d{2}/\d{2}/\d{4}$", c)), "")
        fuel = next((f for pat, f in prefix if re.search(pat, name, re.I)), fuel_from_text(name))
        if m:
            proj.append({"Code": r[0], "Project": re.sub(r"\s*\d+(?:[.,]\d+)?\s*MW.*", "", name).strip(), "Fuel": fuel,
                         "MW": _num(m.group(1)), "Model entry date": pd.to_datetime(when, dayfirst=True, errors="coerce"),
                         "Connection": r[3] if len(r) > 3 else "", "File": r[-1]})
    if proj:
        t = pd.DataFrame(proj).drop_duplicates(["Project", "MW"])
        print(f"  EPO model new projects: {len(t)}, {t['MW'].sum():,.0f} MW: {t.groupby('Fuel')['MW'].sum().to_dict()}",
              flush=True)
        sheets["PE EPO model new projects"] = t.set_index("Code")
        t["Year"] = t["Model entry date"].dt.year
        sheets["PE EPO new projects by year"] = by_year_fuel(t, "Year", "Fuel", "MW")
        src.append("COES, Modelo Eléctrico del SEIN para EPO - 'Lo Nuevo' (projects added to the model): " + "; ".join(epo))
    if not sheets:
        raise RuntimeError("nothing parsed from COES")
    return sheets, " ; ".join(src)


def argentina():
    import requests
    from future_common import HEADERS
    s = requests.Session()   # the portal redirects through a cookie check: a session keeps the cookie
    s.headers.update(HEADERS)
    content = None
    for url in (AR_OBRAS, AR_OBRAS.replace("https://", "http://")):
        try:
            r = s.get(url, timeout=(10, 300))
            r.raise_for_status()
            content = r.content
            break
        except Exception as e:  # noqa: BLE001
            print(f"  AR {url[:40]}...: {type(e).__name__}", flush=True)
    if content is None:   # the direct link moves: look it up in the open-data catalogue
        res = s.get(AR_CKAN, params={"q": "obras generacion", "rows": 10}, timeout=(10, 120)).json()
        url = next(x["url"] for p in res["result"]["results"] for x in p.get("resources", [])
                   if re.search("obra", x.get("url", ""), re.I) and x["url"].lower().endswith(".csv"))
        content = s.get(url, timeout=(10, 300)).content
    d = _csv(content)
    print(f"AR obras: {d.shape}; columns {list(d.columns)}\n{d.head(3).to_string()[:1500]}", flush=True)
    tech = col(d, r"tecnolog|fuente|tipo")
    mw = col(d, r"potencia|mw")
    d["Fuel"] = d[tech].map(fuel_from_text)
    d["MW"] = pd.to_numeric(d[mw].astype(str).str.replace(",", "."), errors="coerce")
    sheets = {}
    st = col(d, r"estado|situaci", required=False)
    if st:
        t = d.pivot_table(index=st, columns="Fuel", values="MW", aggfunc="sum").fillna(0).round(0)
        t.index.name = "Status"
        sheets["AR works by status"] = t
    yc = col(d, r"a[ñn]o|fecha.*(habilit|operaci|fin)|habilitaci", required=False)
    if yc:
        yr = pd.to_numeric(d[yc], errors="coerce")
        if yr.isna().all():
            yr = pd.to_datetime(d[yc], errors="coerce", dayfirst=True).dt.year
        if yr.notna().any():
            sheets["AR works by year"] = by_year_fuel(d.assign(_y=yr), "_y", "Fuel", "MW", start=2021)
    sheets["AR works list"] = d.set_index(d.columns[0])
    return sheets


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=DEFAULT_OUT)
    args = ap.parse_args()
    sheets, src = {}, []
    for name, fn, s in (("Brazil", brazil, "ANEEL SIGA plant register (phase: under construction / not started)"),
                        ("Brazil RALIE", brazil_ralie, "ANEEL RALIE generation expansion monitoring"),
                        ("Chile", chile, "CNE, Instalaciones declaradas en construcción"),
                        ("Colombia", colombia, "UPME, registro de proyectos de generación (registros activos)"),
                        ("Peru", peru, "COES"),
                        ("Argentina", argentina, f"Secretaría de Energía, Obras de generación: {AR_OBRAS}")):
        try:
            got = fn()
            if isinstance(got, tuple):
                got, url = got
                s = f"{s}: {url}"
            sheets.update(got)
            src.append(s)
        except Exception as e:  # noqa: BLE001
            print(f"{name} failed: {type(e).__name__}: {e}", flush=True)
    if not sheets:
        raise SystemExit("nothing fetched")
    notes = [
        "UNITS",
        "MW by fuel. Brazil: ANEEL granted capacity of plants under construction or granted but not started "
        "(many granted wind/solar projects never get built - read 'not started' as an upper bound). Expected "
        "additions: ANEEL RALIE projects by forecast commercial-operation year. Argentina: "
        "generation works listed by the Secretaría de Energía, MW as published.",
        "",
        "COVERAGE",
        "Brazil, Chile (CNE: projects declared in construction, incl. small distributed PMGD and BESS), Colombia "
        "(UPME project register: FASE 1 prefeasibility, FASE 2 feasibility, FASE 3 design/construction - a "
        "registration is not a commitment), Peru (COES), Argentina. One-off snapshot, not refreshed on a schedule.",
        "",
        "SOURCE",
        *src,
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    xlsx_notes.write_workbook(args.out, sheets, notes, {"UNITS", "COVERAGE", "SOURCE"})
    print(f"Saved {args.out}: {list(sheets)}", flush=True)


if __name__ == "__main__":
    main()
