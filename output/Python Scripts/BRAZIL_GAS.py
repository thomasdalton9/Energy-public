"""
Brazil monthly natural gas: demand by end-use segment (MME, frozen at
2025-06) plus current pipeline-grid demand by consumer type and supply by
source (ANP open data, to the latest published month).

Why two publishers. MME's monthly gas bulletin annex
"historico-balanco-boletim.xlsx" is the only public national split by
end-use segment (industrial, vehicle CNG, residential, commercial, power,
cogeneration, other). It stops at 2025-06; the PDF bulletins stop at
Sep-2025 and only carry an interconnected-grid table; the 2026 bulletin
pages return 404; the Observatorio "Boletim do Gas" that replaced them needs
a gov.br login; ABEGAS publishes annual totals only; EPE has no monthly gas
series (BRAZIL_GAS_2026_DISCOVERY*.py). So "Demand by segment" stays MME and
is NOT extended with another series.

What is current is ANP open data:
  - "Movimentacao de Gas Natural em Gasodutos de Transporte": one CSV per
    month, daily "Volume Realizado" (mil m3) at every receipt (PTR) and
    delivery (PTE) point of the transport pipelines (TAG, NTS, TBG, TSB, GOM).
    Delivery points are classified with ANP's own unit register (read from
    ANP's public Power BI panel on the same data; name rules as fallback):
    city gates to state distributors, thermal power plants, refineries,
    fertiliser plants. Pipeline-to-pipeline interconnections are dropped so
    nothing is counted twice; Cabiunas (TECAB) and Itaborai are processing-
    plant outlets despite their "Interconexao" labels and count as domestic
    receipts.
  - production by state ("gas natural disponivel", mil m3/month) and total
    gas imports (mil m3/month).
Grid flows cover gas that moves through the transport pipelines only: they
leave out isolated systems (e.g. Parnaiba reservoir-to-wire plants), LNG
terminals that feed power plants or distributors directly, and gas sold
straight from processing plants. The "Overlap vs MME" sheet compares both
publishers month by month for 2021-01..2025-06.

Units: million m3/day (monthly average) throughout.

Incremental: the per-point monthly table ("Grid points (ANP)") is kept in
the workbook; each run fetches only months missing from it, months whose
published file name changed (ANP revisions, "_rev1"), and the latest two.

Outputs (brazil_gas_monthly.xlsx):
  Demand by segment   MME end-use segments, 2021-01..2025-06, Source column
  Grid demand (ANP)   pipeline deliveries by consumer type, to latest month
  Supply (ANP)        domestic available gas, pipeline imports (Bolivia,
                      Argentina), total imports, LNG implied, to latest month
  Overlap vs MME      both publishers side by side where they overlap
  Balance (all rows)  every labelled row of MME's balance sheet
  Grid points (ANP)   month x point raw table (feeds the grid sheets)
"""
import argparse
import base64
import io
import json
import os
import re
import sys
import unicodedata

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root, for xlsx_notes

import xlsx_notes

URL = ("https://www.gov.br/mme/pt-br/assuntos/secretarias/petroleo-gas-natural-e-biocombustiveis/publicacoes-1/"
       "boletim-mensal-de-acompanhamento-da-industria-de-gas-natural/anexos/historico-balanco-boletim.xlsx")
ANP_DA = "https://www.gov.br/anp/pt-br/centrais-de-conteudo/dados-abertos"
ANP_MOV_PAGE = f"{ANP_DA}/dados-consolidados-movimentacao-de-gas-natural-em-gasodutos-de-transporte"
ANP_PROD_PAGE = f"{ANP_DA}/producao-de-petroleo-e-gas-natural-por-estado-e-localizacao"
ANP_IMP_PAGE = f"{ANP_DA}/importacoes-e-exportacoes"
ANP_AVAILABLE = f"{ANP_DA}/arquivos/ppgn-el/gn-disponivel-1000m3.csv"
ANP_IMPORTS = f"{ANP_DA}/arquivos/ie/gn/importacao-gas-natural-2000-2025.csv"
ANP_PBI_PAGE = ("https://www.gov.br/anp/pt-br/centrais-de-conteudo/paineis-dinamicos-da-anp/"
                "painel-dinamico-de-movimentacao-de-gas-natural-em-gasodutos-de-transporte")
PBI_API = "https://wabi-brazil-south-api.analysis.windows.net"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 120)
DATA_START = "2021-01-01"
REFRESH_LATEST = 2          # always re-read the newest N grid months (late revisions)
OUT_DEFAULT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output", "brazil_gas_monthly.xlsx")

SRC_MME = "MME historical balance annex"
SRC_GRID = "ANP pipeline movement open data"
SRC_SUPPLY = "ANP production + imports open data; ANP pipeline movement (Bolivia/Argentina)"

SEGMENTS = {  # normalised label prefix -> output column
    "industrial": "Industrial",
    "automotivo": "Automotive",
    "residencial": "Residential",
    "comercial": "Commercial",
    "geracao eletrica": "Power_Generation",
    "cogeracao": "Cogeneration",
    "outros": "Other_incl_CNG",
    "demanda total": "Total_Demand",
}
GRID_DEMAND = ["Distributors", "Power_Generation", "Refineries", "Fertiliser", "Other"]
PT_MONTHS = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "out": 10,
             "nov": 11, "dez": 12}
PT_MONTH_NAMES = {"janeiro": 1, "fevereiro": 2, "marco": 3, "abril": 4, "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
                  "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12}

S = requests.Session()
S.headers.update(HEADERS)


def log(*a):
    print(*a, flush=True)


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def decode(content):
    try:
        return content.decode("utf-8-sig")
    except UnicodeDecodeError:
        return content.decode("latin-1")


def num(series):
    """Brazilian-format numbers ("1.234,56") -> float."""
    s = series.astype(str).str.strip()
    s = s.where(~s.str.contains(","), s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(s, errors="coerce")


def get(url):
    r = S.get(url, timeout=TIMEOUT)
    r.raise_for_status()
    return r


# ------------------------------------------------------------------ MME (segments, frozen)

def to_month(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 30000 < v < 60000:  # Excel serial date
        return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(v))).to_period("M").to_timestamp()
    if isinstance(v, (pd.Timestamp,)) or hasattr(v, "year"):
        try:
            return pd.Timestamp(v).to_period("M").to_timestamp()
        except Exception:
            return None
    m = re.match(r"^\s*([A-Za-zçÇ]{3})[a-zç]*[\s/\-.]*(\d{2,4})\s*$", str(v))
    if m and norm(m.group(1))[:3] in PT_MONTHS:
        y = int(m.group(2))
        y = y + 2000 if y < 100 else y
        return pd.Timestamp(y, PT_MONTHS[norm(m.group(1))[:3]], 1)
    return None


def fetch_mme():
    r = get(URL)
    if r.content[:2] != b"PK":
        raise RuntimeError(f"Not an xlsx (content-type {r.headers.get('content-type')}) - MME may have moved the file")
    xl = pd.ExcelFile(io.BytesIO(r.content))
    raw = xl.parse(xl.sheet_names[0], header=None)
    best_row, best_n = None, 0
    for i in range(min(12, len(raw))):
        n = sum(to_month(v) is not None for v in raw.iloc[i, 1:])
        if n > best_n:
            best_row, best_n = i, n
    if best_row is None or best_n < 12:
        raise RuntimeError("Could not find the month header row")
    months = [to_month(v) for v in raw.iloc[best_row, 1:]]
    rows = {}
    for i in range(best_row + 1, len(raw)):
        label = raw.iloc[i, 0]
        if pd.isna(label) or not str(label).strip():
            continue
        vals = pd.to_numeric(raw.iloc[i, 1:], errors="coerce")
        if vals.notna().sum() == 0:
            continue
        rows[str(label).strip()] = pd.Series(vals.to_numpy(), index=months)
    bal = pd.DataFrame(rows)
    bal = bal[bal.index.notna()]
    bal.index = pd.DatetimeIndex(bal.index, name="date")
    bal = bal.sort_index().dropna(how="all")
    log(f"MME annex: {len(bal.columns)} rows, {bal.index.min():%Y-%m} to {bal.index.max():%Y-%m}")
    return bal


def mme_segments(bal):
    demand = pd.DataFrame(index=bal.index)
    for col in bal.columns:
        n = norm(col)
        for prefix, out in SEGMENTS.items():
            if n.startswith(prefix) and out not in demand.columns:
                demand[out] = bal[col]
    missing = set(SEGMENTS.values()) - set(demand.columns)
    if missing:
        raise RuntimeError(f"segments not found in MME sheet: {missing} - labels were {list(bal.columns)}")
    parts = [c for c in demand.columns if c != "Total_Demand"]
    gap = (demand[parts].sum(axis=1) - demand["Total_Demand"]).abs().max()
    log(f"MME segments sum vs DEMANDA TOTAL: max abs diff {gap:.3f} million m3/d")
    return demand[list(SEGMENTS.values())].dropna(how="all")


def mme_row(bal, prefix):
    for c in bal.columns:
        if norm(c).startswith(prefix):
            return bal[c]
    return pd.Series(dtype=float, index=bal.index)


# ------------------------------------------------------------------ ANP grid movement

def list_movement_files():
    """{month: url} from the ANP dataset page; one file per month (a revised file wins)."""
    html = get(ANP_MOV_PAGE).text
    found = {}
    for href in sorted(set(re.findall(r'href="([^"]+\.csv)"', html, re.I))):
        m = re.search(r"/(20\d\d)/([^/]+)\.csv$", href, re.I)
        if not m:
            continue
        name = norm(m.group(2).replace("_", " ").replace("-", " "))
        mon = next((v for k, v in PT_MONTH_NAMES.items() if re.search(rf"\b{k}\b", name)), None)
        if mon is None:
            continue
        key = pd.Timestamp(int(m.group(1)), mon, 1)
        if key not in found or "rev" in name:
            found[key] = href
    log(f"ANP movement files listed: {len(found)}, {min(found):%Y-%m} to {max(found):%Y-%m}")
    return found


def read_movement(url):
    """One monthly CSV -> one row per (flow, pipeline, point) with the month's mean mcm/d."""
    df = pd.read_csv(io.StringIO(decode(get(url).content)), sep=";", dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    days = [c for c in df.columns if re.match(r"\d{2}/\d{2}/\d{4}$", c)]
    if not days:
        raise RuntimeError(f"no day columns in {url}")
    month = pd.to_datetime(days[0], dayfirst=True).to_period("M").to_timestamp()
    ndays = month.days_in_month

    def col(pattern):
        c = next((c for c in df.columns if re.search(pattern, norm(c))), None)
        if c is None:
            raise RuntimeError(f"column /{pattern}/ not in {list(df.columns)[:14]}")
        return c

    var, tipo = col(r"^nome da variavel"), col(r"^tipo da instalacao")
    pipe, point, code, uf = (col(r"^nome da instalacao de transporte"), col(r"^nome da instalacao de gasoduto"),
                             col(r"^codigo da instalacao de gasoduto"), col(r"^nome da uf"))
    v = df[df[var].astype(str).map(norm).str.startswith("volume realizado")].copy()
    vals = v[days].apply(num)
    v["mcm_d"] = vals.sum(axis=1, min_count=1) / ndays / 1000.0     # mil m3/day summed over days -> mcm/d average
    v["flow"] = v[tipo].astype(str).map(norm).map(
        lambda s: "Receipt" if "receb" in s else "Delivery" if "entrega" in s else "Unknown")
    v["point_code"] = pd.to_numeric(v[code], errors="coerce").astype("Int64").astype(str)
    out = (v.groupby(["flow", pipe, point, "point_code", uf], dropna=False)["mcm_d"].sum().reset_index()
           .rename(columns={pipe: "pipeline", point: "point", uf: "uf"}))
    out.insert(0, "month", month)
    out["file"] = url.rsplit("/", 1)[-1]
    return out


def pbi_units():
    """ANP's own point register (type / classification / consumer type) keyed by public point code, read from
    the public Power BI panel on the same data. {} if unavailable - the name rules then decide."""
    try:
        tok = re.findall(r"app\.powerbi\.com/view\?r=([A-Za-z0-9=_%-]+)", get(ANP_PBI_PAGE).text)[0]
        k = json.loads(base64.b64decode(tok + "=" * (-len(tok) % 4)))["k"]
        hdr = {"X-PowerBI-ResourceKey": k, "Content-Type": "application/json;charset=UTF-8"}
        me = S.get(f"{PBI_API}/public/reports/{k}/modelsAndExploration?preferReadOnlySession=true", headers=hdr,
                   timeout=TIMEOUT).json()
        model = me["models"][0]
        props = ["COD_PBLCO_UNIDADE", "DSC_TIPO_UNIDADE", "DSC_CLASSIFICACAO", "DSC_TIPO_CONSUMIDOR", "DSC_UNIDADE"]
        sel = [{"Column": {"Expression": {"SourceRef": {"Source": "u"}}, "Property": p}, "Name": f"u.{p}"} for p in props]
        body = {"version": "1.0.0", "queries": [{"Query": {"Commands": [{"SemanticQueryDataShapeCommand": {
            "Query": {"Version": 2, "From": [{"Name": "u", "Entity": "DM_SIMP_UNIDADE", "Type": 0}], "Select": sel},
            "Binding": {"Primary": {"Groupings": [{"Projections": list(range(len(props)))}]},
                        "DataReduction": {"DataVolume": 4, "Primary": {"Window": {"Count": 30000}}}, "Version": 1}}}]},
            "ApplicationContext": {"DatasetId": model["dbName"],
                                   "Sources": [{"ReportId": me["exploration"]["report"]["objectId"]}]}}],
                "cancelQueries": [], "modelId": model["id"]}
        resp = S.post(f"{PBI_API}/public/reports/querydata?synchronous=true", headers=hdr, data=json.dumps(body),
                      timeout=TIMEOUT).json()
        ds = resp["results"][0]["result"]["data"]["dsr"]["DS"][0]
        dicts, prev, schema, units = ds.get("ValueDicts", {}), [None] * len(props), None, {}
        for r in ds["PH"][0].get("DM0", []):
            schema = r.get("S", schema)
            c, row = list(r.get("C", [])), []
            for i in range(len(props)):
                if r.get("R", 0) >> i & 1:
                    row.append(prev[i])
                elif r.get("Ø", 0) >> i & 1:
                    row.append(None)
                else:
                    val = c.pop(0) if c else None
                    dn = schema[i].get("DN") if schema and i < len(schema) else None
                    row.append(dicts[dn][val] if dn and isinstance(val, int) else val)
            prev = row
            if row[0] is not None:
                units[str(int(float(row[0])))] = {"tipo": norm(row[1] or ""), "cls": norm(row[2] or ""),
                                                  "cons": norm(row[3] or ""), "name": row[4]}
        log(f"ANP Power BI point register: {len(units)} points")
        return units
    except Exception as e:
        log(f"ANP Power BI point register unavailable ({type(e).__name__}: {str(e)[:150]}) - name rules only")
        return {}


INTERCONNECT = re.compile(r"interconex|>>|\bemed\b|\bemr\b|\becgm\b|\besbc\b")
PROCESSING = re.compile(r"\(\s*tecab\s*>>|itabora")      # matched on the lower-case raw name
REFINERY = re.compile(r"\b(replan|reduc|refap|regap|repar|revap|recap|rlam|rnest|rpbc|reman|lubnor)\b|upgn candeias")


def classify(flow, point, pipeline, code, units):
    """Category of one point. Name rules for interconnections and processing outlets win over the register
    (it files some interconnections as PROPRIO and the TECAB outlets as third-party interconnections)."""
    n, raw = norm(point), str(point).lower()
    u = units.get(str(code), {})
    if flow not in ("Receipt", "Delivery") or n in ("", "nan", "none"):
        return "Unidentified"
    if flow == "Receipt":
        if PROCESSING.search(raw):
            return "Domestic"
        if INTERCONNECT.search(n) or ">>" in raw:
            return "Interconnection"
        if norm(pipeline).startswith("uruguaiana") and "canoas" in n:
            return "Interconnection"          # Canoas: GASBOL gas entering the Uruguaiana-Porto Alegre line
        if re.search(r"corumb|mutun|caceres|bolivia", n):
            return "Bolivia"
        if "uruguaiana" in n:
            return "Argentina"
        if u.get("cls") == "gnl" or re.search(r"\bgnl\b|trba|terminal|pecem", n):
            return "LNG"
        if u.get("cls") == "nacional de terceiros":
            return "Interconnection"
        return "Domestic"
    if INTERCONNECT.search(n) or ">>" in raw or u.get("cons") == "autoprodutor":
        return "Interconnection"
    cons = u.get("cons")
    if cons == "termeletrica":
        return "Power_Generation"
    if cons == "refinaria":
        return "Refineries"
    if cons == "fafen":
        return "Fertiliser"
    if cons == "concessionaria estadual":
        return "Distributors"
    if REFINERY.search(n):
        return "Refineries"
    if re.search(r"fafen|\bufn\b", n):
        return "Fertiliser"
    if re.search(r"\bute\b|termo|\bterm\b", n) and not re.search(r"\bceg\b|city", n):
        return "Power_Generation"
    if cons == "nao informado" and u:
        return "Other"
    return "Distributors"


def update_points(existing, start):
    files = list_movement_files()
    files = {m: u for m, u in files.items() if m >= pd.Timestamp(start)}
    have = {}
    if existing is not None and not existing.empty:
        have = existing.groupby("month")["file"].first().to_dict()
    latest = sorted(files)[-REFRESH_LATEST:]
    todo = [m for m in sorted(files) if m not in have or have[m] != files[m].rsplit("/", 1)[-1] or m in latest]
    log(f"grid months kept: {len([m for m in have if m not in todo])}, to fetch: {[f'{m:%Y-%m}' for m in todo]}")
    new = []
    for m in todo:
        try:
            d = read_movement(files[m])
            new.append(d)
            log(f"  {m:%Y-%m}: {len(d)} points, receipts {d[d.flow == 'Receipt'].mcm_d.sum():.1f}, "
                f"deliveries {d[d.flow == 'Delivery'].mcm_d.sum():.1f} mcm/d (gross, incl. interconnections)")
        except Exception as e:
            log(f"  {m:%Y-%m}: FAILED {type(e).__name__}: {str(e)[:200]} - keeping any stored copy")
    keep = existing[~existing["month"].isin([d["month"].iloc[0] for d in new])] if have else None
    pts = pd.concat([x for x in [keep, *new] if x is not None and not x.empty], ignore_index=True)
    return pts.sort_values(["month", "flow", "pipeline", "point"]).reset_index(drop=True)


def grid_tables(pts, units):
    cached = {}
    if "category" in pts:
        cached = pts.dropna(subset=["category"]).groupby(["flow", "point_code"])["category"].last().to_dict()
    cats = []
    for r in pts.itertuples(index=False):
        c = classify(r.flow, r.point, r.pipeline, r.point_code, units)
        if not units and (r.flow, r.point_code) in cached and not INTERCONNECT.search(norm(r.point)):
            c = cached[(r.flow, r.point_code)]     # register down this run: keep last run's register-based call
        cats.append(c)
    pts = pts.assign(category=cats)
    odd = pts[pts.category == "Unidentified"].groupby("month")["mcm_d"].sum()
    if odd.abs().max() > 0.05:
        log(f"WARNING rows with no point name/flow left out, mcm/d: {odd[odd.abs() > 0.05].round(2).to_dict()}")
    piv = pts.pivot_table(index="month", columns=["flow", "category"], values="mcm_d", aggfunc="sum").fillna(0.0)
    dem = pd.DataFrame(index=piv.index)
    for c in GRID_DEMAND:
        dem[c] = piv[("Delivery", c)] if ("Delivery", c) in piv else 0.0
    dem["Total_Grid_Deliveries"] = dem[GRID_DEMAND].sum(axis=1)
    rec = pd.DataFrame(index=piv.index)
    for c in ["Domestic", "Bolivia", "Argentina", "LNG"]:
        rec[c] = piv[("Receipt", c)] if ("Receipt", c) in piv else 0.0
    rec["Total"] = rec.sum(axis=1)
    dem.index.name = rec.index.name = "date"
    imb = rec["Total"] - dem["Total_Grid_Deliveries"]
    log("grid receipts minus deliveries (system use, line pack, losses), mcm/d: "
        f"median {imb.median():.2f}, min {imb.min():.2f} ({imb.idxmin():%Y-%m}), max {imb.max():.2f} ({imb.idxmax():%Y-%m})")
    return pts, dem.round(3), rec.round(3), imb.round(3)


# ------------------------------------------------------------------ ANP production / imports

def find_link(page, pattern, fallback):
    try:
        links = re.findall(r'href="([^"]+\.csv)"', get(page).text, re.I)
        hit = [l for l in links if re.search(pattern, l, re.I)]
        return hit[-1] if hit else fallback
    except requests.RequestException:
        return fallback


def monthly_csv(url, value_hint):
    d = pd.read_csv(io.StringIO(decode(get(url).content)), sep=";", dtype=str)
    d.columns = [norm(c) for c in d.columns]
    val = next(c for c in d.columns if c.startswith(value_hint))
    d["v"] = num(d[val])
    d["date"] = [pd.Timestamp(int(y), PT_MONTHS[norm(m)[:3]], 1) for y, m in zip(d["ano"], d["mes"])]
    s = d.groupby("date")["v"].sum(min_count=1)
    return s / s.index.days_in_month / 1000.0           # mil m3/month -> mcm/d


def anp_supply(rec):
    avail = monthly_csv(find_link(ANP_PROD_PAGE, r"gn-disponivel", ANP_AVAILABLE), "disponivel")
    imports = monthly_csv(find_link(ANP_IMP_PAGE, r"importacao-gas-natural", ANP_IMPORTS), "importado")
    sup = pd.DataFrame({"Domestic_Available": avail, "Imports_Total": imports})
    sup = sup.join(rec[["Bolivia", "Argentina"]].rename(columns=lambda c: f"{c}_Pipeline"), how="outer")
    sup["Grid_Domestic_Receipts"] = rec["Domestic"]
    sup["LNG_Implied"] = (sup["Imports_Total"] - sup["Bolivia_Pipeline"].fillna(0) - sup["Argentina_Pipeline"].fillna(0)).clip(lower=0)
    sup["Total_Supply"] = sup["Domestic_Available"] + sup["Imports_Total"]
    sup.index.name = "date"
    log(f"ANP gas available: to {avail.dropna().index.max():%Y-%m}; imports: to {imports.dropna().index.max():%Y-%m}")
    cols = ["Domestic_Available", "Bolivia_Pipeline", "Argentina_Pipeline", "LNG_Implied", "Imports_Total",
            "Total_Supply", "Grid_Domestic_Receipts"]
    return sup[cols].round(3)


# ------------------------------------------------------------------ overlap / source watch / main

def overlap(bal, seg, dem, sup):
    nonpower = ["Industrial", "Automotive", "Residential", "Commercial", "Cogeneration", "Other_incl_CNG"]
    o = pd.DataFrame({
        "MME_Total_Demand": seg["Total_Demand"],
        "ANP_Grid_Deliveries": dem["Total_Grid_Deliveries"],
        "MME_Power_Generation": seg["Power_Generation"],
        "ANP_Grid_Power_Generation": dem["Power_Generation"],
        "MME_NonPower": seg[nonpower].sum(axis=1),
        "ANP_Grid_NonPower": dem[["Distributors", "Refineries", "Fertiliser", "Other"]].sum(axis=1),
        "MME_Oferta_Nacional": mme_row(bal, "oferta nacional"),
        "ANP_Domestic_Available": sup["Domestic_Available"],
        "MME_Import_Bolivia": mme_row(bal, "importacao bolivia"),
        "ANP_Bolivia_Pipeline": sup["Bolivia_Pipeline"],
        "MME_LNG_Regas": mme_row(bal, "regaseificacao"),
        "ANP_LNG_Implied": sup["LNG_Implied"],
        "MME_Oferta_Importada": mme_row(bal, "oferta importada"),
        "ANP_Imports_Total": sup["Imports_Total"],
    })
    o = o[o["MME_Total_Demand"].notna() & o["ANP_Grid_Deliveries"].notna()]
    o.index.name = "date"
    lines = []
    for a in range(0, len(o.columns), 2):
        m, n = o.columns[a], o.columns[a + 1]
        d = (o[n] - o[m]).dropna()
        if len(d):
            lines.append(f"{n} - {m}: mean {d.mean():+.2f}, mean abs {d.abs().mean():.2f}, "
                         f"min {d.min():+.2f}, max {d.max():+.2f} mcm/d over {len(d)} months "
                         f"({d.index.min():%Y-%m}..{d.index.max():%Y-%m}); ratio of means {o[n].mean() / o[m].mean():.2f}")
    log("\nOVERLAP vs MME\n" + "\n".join(lines))
    return o.round(3), lines


BULLETIN_BASE = ("https://www.gov.br/mme/pt-br/assuntos/secretarias/petroleo-gas-natural-e-biocombustiveis/publicacoes-1/"
                 "boletim-mensal-de-acompanhamento-da-industria-de-gas-natural")
OBSERVATORIO = ("https://www.gov.br/mme/pt-br/assuntos/observatorio-de-minas-e-energia/petroleo-gas-e-biocombustiveis/"
                "boletim-do-gas")


def source_watch(path, mme_last, grid_last, sup):
    """Signals that would let the segment split resume (MME annex month, a 2026 bulletin page, the Observatorio
    dashboard opening up) plus the ANP months, written to a small file so a change shows as a git diff."""
    lines = [f"MME annex latest month: {mme_last:%Y-%m}" if mme_last is not None else "MME annex: unavailable"]
    for year in (2025, 2026, 2027):
        try:
            r = S.get(f"{BULLETIN_BASE}/{year}", timeout=TIMEOUT)
            pdfs = sorted(set(re.findall(r"/(\d{2})-boletim[^\"/]*?\.pdf", r.text)))
            lines.append(f"MME bulletin page {year}: HTTP {r.status_code}, monthly PDFs listed: {','.join(pdfs) or 'none'}")
        except requests.RequestException as e:
            lines.append(f"MME bulletin page {year}: ERR {type(e).__name__}")
    try:
        r = S.get(OBSERVATORIO, timeout=TIMEOUT)
        gated = "require_login" in r.text or "credentials_cookie_auth" in r.text
        lines.append(f"MME Observatorio boletim do gas: HTTP {r.status_code}, login-gated={gated}")
    except requests.RequestException as e:
        lines.append(f"MME Observatorio boletim do gas: ERR {type(e).__name__}")
    lines.append(f"ANP pipeline movement latest month: {grid_last:%Y-%m}")
    for c in ("Domestic_Available", "Imports_Total"):
        s = sup[c].dropna()
        lines.append(f"ANP {c} latest month: {s.index.max():%Y-%m}" if len(s) else f"ANP {c}: none")
    text = "\n".join(lines) + "\n"
    log("\nSOURCE WATCH\n" + text)
    with open(path, "w") as f:
        f.write(text)


def load_sheet(path, sheet):
    try:
        d = pd.read_excel(path, sheet_name=sheet)
        return d.drop(columns=[c for c in d.columns if str(c).startswith("Unnamed")])
    except Exception:
        return None


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=OUT_DEFAULT)
    parser.add_argument("--start-date", default=DATA_START)
    parser.add_argument("--watch-out", default=None, help="write the source-watch status file here")
    args = parser.parse_args()
    start = pd.Timestamp(args.start_date)

    # MME (segments): re-read the annex; if MME pulls it, keep the stored copy
    try:
        bal = fetch_mme()
    except Exception as e:
        log(f"MME annex unavailable ({type(e).__name__}: {e}) - keeping the stored balance")
        bal = load_sheet(args.out, "Balance (all rows)")
        if bal is None:
            raise
        bal = bal.set_index(pd.to_datetime(bal.iloc[:, 0])).iloc[:, 1:]
        bal.index.name = "date"
    mme_last = bal.dropna(how="all").index.max()
    bal = bal[bal.index >= start]
    seg = mme_segments(bal)
    log(seg.tail(3).round(2).to_string())

    # ANP grid (incremental)
    existing = load_sheet(args.out, "Grid points (ANP)")
    if existing is not None:
        existing["month"] = pd.to_datetime(existing["month"])
        existing["point_code"] = existing["point_code"].astype(str).str.replace(r"\.0$", "", regex=True)
    pts = update_points(existing, start)
    pts, dem, rec, imb = grid_tables(pts, pbi_units())
    log("\ngrid deliveries by consumer type (mcm/d), last 6 months\n" + dem.tail(6).round(2).to_string())
    log("\ngrid receipts by source (mcm/d), last 6 months\n" + rec.tail(6).round(2).to_string())
    sup = anp_supply(rec)
    sup = sup[sup.index >= start]
    log("\nsupply (mcm/d), last 6 months\n" + sup.tail(6).round(2).to_string())
    ov, ov_lines = overlap(bal, seg, dem, sup)
    if args.watch_out:
        source_watch(args.watch_out, mme_last, dem.index.max(), sup)

    dem_out = dem.assign(Receipts_minus_Deliveries=imb, Source=SRC_GRID)
    sup_out = sup.dropna(how="all").assign(Source=SRC_SUPPLY)
    seg_out = seg.assign(Source=SRC_MME)
    grid_last = dem.index.max()
    notes = [
        "UNITS",
        "Million m3 per day (monthly average) on every sheet.",
        "",
        "DEMAND BY SEGMENT (MME)",
        f"End-use segments from MME's monthly gas bulletin annex, {seg.index.min():%Y-%m} to {seg.index.max():%Y-%m}. "
        "This is the only public national split by segment; MME has published nothing newer (2026 bulletin pages "
        "404, the Observatorio 'Boletim do Gas' needs a gov.br login, ABEGAS publishes annual totals only). The "
        "series is not extended with another publisher's data: no current source uses the same definitions.",
        "Industrial includes refineries and fertiliser plants (MME footnote); Automotive = vehicle CNG; "
        "Other_incl_CNG = other uses incl. compressed gas; Total_Demand is MME's own total.",
        "",
        "GRID DEMAND (ANP)",
        f"Monthly mean of daily 'Volume Realizado' at every delivery point of the transport pipelines, "
        f"{dem.index.min():%Y-%m} to {grid_last:%Y-%m}, by ANP's own consumer type: Distributors (city gates of the "
        "state distribution companies, i.e. industrial + vehicle + residential + commercial + cogeneration and "
        "distributor-supplied power plants together), Power_Generation (thermal plants connected to the transport "
        "grid), Refineries, Fertiliser (FAFEN/UFN), Other. Pipeline-to-pipeline interconnections are excluded.",
        "Grid flows leave out gas that never enters a transport pipeline: isolated systems (Parnaiba basin "
        "reservoir-to-wire plants), LNG terminals feeding power plants or distributors directly, gas sold straight "
        "from processing plants. So grid power generation is well below MME's national power figure.",
        "Receipts_minus_Deliveries = grid receipts (processing plants, Bolivia, Argentina, LNG terminals) minus "
        "deliveries: system use, losses and line pack.",
        "",
        "SUPPLY (ANP)",
        "Domestic_Available = ANP 'gas natural disponivel' (production less reinjection, flaring/losses and E&P "
        "own use), all states. Imports_Total = ANP gas imports (all origins). Bolivia_Pipeline / Argentina_Pipeline "
        "= grid receipts at Corumba + Caceres / Uruguaiana. LNG_Implied = Imports_Total minus the pipeline imports "
        "(derived, not published as such). Grid_Domestic_Receipts = domestic gas entering the transport grid.",
        "",
        "OVERLAP vs MME",
        "Both publishers side by side for the months they share. Mean differences (ANP minus MME, mcm/d):",
        *ov_lines,
        "",
        "SOURCE",
        f"MME: {URL}",
        f"ANP pipeline movement: {ANP_MOV_PAGE}",
        f"ANP production (gas disponivel): {ANP_PROD_PAGE}",
        f"ANP imports: {ANP_IMP_PAGE}",
        f"ANP point register (consumer type): Power BI panel at {ANP_PBI_PAGE}",
        "Each data row carries a Source column. Grid months are fetched incrementally (stored months kept; "
        f"missing, revised and the latest {REFRESH_LATEST} months re-read).",
    ]
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    pts_out = pts.assign(month=pts["month"].dt.strftime("%Y-%m-%d"))
    xlsx_notes.write_workbook(args.out, {"Demand by segment": seg_out, "Grid demand (ANP)": dem_out,
                                         "Supply (ANP)": sup_out, "Overlap vs MME": ov, "Balance (all rows)": bal,
                                         "Grid points (ANP)": pts_out.set_index("month")}, notes,
                              {"UNITS", "DEMAND BY SEGMENT (MME)", "GRID DEMAND (ANP)", "SUPPLY (ANP)", "OVERLAP vs MME",
                               "SOURCE"})
    log(f"Saved {args.out}")


if __name__ == "__main__":
    main()
