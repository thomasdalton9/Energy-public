"""
Round 2 of BRAZIL_GAS_2026_DISCOVERY.py. Round 1 found:
  - MME bulletin index links its "2025" year to a new folder
    .../petroleo-gas-natural-e-biocombustiveis/dgn/2025 (maybe dgn/2026 too);
  - MME Observatorio "Outras informacoes analiticas" embeds several public
    Power BI (publish-to-web) reports; "Boletim do Gas" itself is login-gated;
  - ANP open data has "Movimentacao de Gas Natural em Gasodutos de Transporte"
    (consolidated), "Comercializacao de gas natural", "Importacoes e
    exportacoes", production by state; ANP has a "Painel Dinamico de
    Movimentacao de Gas Natural" (Power BI);
  - EPE has a "Consumo de Gas Natural" area page.
This round lists the files on each, dumps CSV columns/categories/latest
month, and reads each public Power BI report's page names and data model
(entities/columns) to see whether any holds national demand by segment.
"""
import base64
import io
import json
import re
from urllib.parse import urljoin

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8"}
T = (10, 90)
S = requests.Session()
S.headers.update(H)


def out(*a):
    print(*a, flush=True)


def get(u, quiet=False, **kw):
    try:
        r = S.get(u, timeout=T, allow_redirects=True, **kw)
        if not quiet:
            out(f"GET {u}\n    -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)}B"
                f"{' REDIRECT ' + r.url if r.url != u else ''}")
        return r
    except Exception as e:
        out(f"GET {u}\n    -> ERR {type(e).__name__} {str(e)[:150]}")
        return None


def page_links(r, pat, limit=60, show=True):
    if r is None or r.status_code != 200 or "html" not in r.headers.get("content-type", ""):
        return []
    found = []
    for href, txt in re.findall(r'<a[^>]+href="([^"#]+)"[^>]*>(.*?)</a>', r.text, re.S | re.I):
        txt = re.sub(r"<[^>]+>|\s+", " ", txt).strip()
        full = urljoin(r.url, href)
        if re.search(pat, full, re.I) or re.search(pat, txt, re.I):
            found.append((full, txt[:110]))
    found = list(dict.fromkeys(found))
    if show:
        for f, t in found[:limit]:
            out(f"      {f} | {t}")
        if len(found) > limit:
            out(f"      ... {len(found) - limit} more")
    return found


def embeds(r):
    return sorted(set(re.findall(r"app\.powerbi\.com/view\?r=([A-Za-z0-9=_%-]+)", r.text))) if r is not None else []


def dump_csv(u, maxrows=None):
    import pandas as pd
    r = get(u)
    if r is None or r.status_code != 200:
        return None
    c = r.content
    try:
        if c[:2] == b"PK" and not u.lower().endswith(".csv"):
            xl = pd.ExcelFile(io.BytesIO(c))
            out(f"    sheets {xl.sheet_names[:20]}")
            df = xl.parse(xl.sheet_names[0])
        else:
            txt = c.decode("utf-8-sig", errors="replace") if b"\xc3" in c[:20000] else c.decode("latin-1")
            sep = ";" if txt[:3000].count(";") > txt[:3000].count(",") else ","
            df = pd.read_csv(io.StringIO(txt), sep=sep, engine="python", on_bad_lines="skip", dtype=str)
        out(f"    shape {df.shape}; cols {list(df.columns)[:40]}")
        out("      " + df.head(3).to_string(max_colwidth=28, max_cols=20)[:1500].replace("\n", "\n      "))
        out("      " + df.tail(3).to_string(max_colwidth=28, max_cols=20)[:1500].replace("\n", "\n      "))
        for col in df.columns:
            n = df[col].nunique()
            if 1 < n <= 60 and not re.search(r"volume|valor|quantidade|m3|m³", str(col), re.I):
                out(f"    distinct {col} ({n}): {sorted(map(str, df[col].dropna().unique()))[:60]}")
            if re.search(r"^ano|m[eê]s|data|per[ií]odo|date|refer", str(col), re.I):
                v = df[col].dropna().astype(str)
                out(f"    {col}: min {v.min()} max {v.max()} (last rows: {list(v.tail(3))})")
        return df
    except Exception as e:
        out(f"    parse ERR {type(e).__name__} {str(e)[:200]}")
        return None


def powerbi(token, label):
    """Page names and model entities of a publish-to-web Power BI report."""
    try:
        cfg = json.loads(base64.b64decode(token + "=" * (-len(token) % 4)).decode())
    except Exception as e:
        out(f"  PBI {label}: token decode ERR {e}")
        return
    k, t = cfg.get("k"), cfg.get("t")
    try:
        cl = S.get(f"https://api.powerbi.com/public/routing/cluster/{t}", timeout=T).json()
        base = cl["FixedClusterUri"].replace("-redirect", "-api").rstrip("/")
    except Exception as e:
        base = "https://wabi-brazil-south-api.analysis.windows.net"
        out(f"  PBI cluster lookup failed ({e}); trying {base}")
    hdr = {"X-PowerBI-ResourceKey": k, "Content-Type": "application/json;charset=UTF-8",
           "Origin": "https://app.powerbi.com", "Referer": "https://app.powerbi.com/"}
    try:
        m = S.get(f"{base}/public/reports/{k}/modelsAndExploration?preferReadOnlySession=true", headers=hdr, timeout=T)
        j = m.json()
    except Exception as e:
        out(f"  PBI {label}: modelsAndExploration ERR {type(e).__name__} {str(e)[:120]}")
        return
    exp = j.get("exploration", {})
    out(f"\n  PBI {label}: report '{exp.get('report', {}).get('displayName', '?')}' "
        f"pages={[s.get('displayName') for s in exp.get('sections', [])]}")
    for s in exp.get("sections", [])[:12]:
        titles = []
        for vc in s.get("visualContainers", [])[:40]:
            try:
                c = json.loads(vc.get("config", "{}"))
                sv = c.get("singleVisual", {})
                ttl = sv.get("vcObjects", {}).get("title", [{}])[0].get("properties", {}).get("text", {}) \
                    .get("expr", {}).get("Literal", {}).get("Value")
                proj = [p.get("queryRef") for v in sv.get("projections", {}).values() for p in v][:6]
                if ttl or proj:
                    titles.append(f"{sv.get('visualType')}:{ttl}:{proj}")
            except Exception:
                pass
        out(f"    page '{s.get('displayName')}': " + " || ".join(titles)[:1500])
    models = j.get("models", [])
    if not models:
        return
    try:
        cs = S.post(f"{base}/public/reports/conceptualschema", headers=hdr, timeout=T,
                    data=json.dumps({"ModelIds": [models[0]["id"]], "UserPreferredLocale": "pt-BR"})).json()
        for sch in cs.get("schemas", []):
            for e in sch.get("schema", {}).get("Entities", []):
                props = [p.get("Name") for p in e.get("Properties", [])]
                out(f"    entity {e.get('Name')}: {props[:40]}")
    except Exception as e:
        out(f"    conceptualschema ERR {type(e).__name__} {str(e)[:150]}")


MME = "https://www.gov.br/mme/pt-br"
out("\n######## 1. MME DGN folders ########")
DGN = f"{MME}/assuntos/secretarias/petroleo-gas-natural-e-biocombustiveis/dgn"
for u in [DGN, f"{DGN}/2025", f"{DGN}/2026", f"{DGN}/boletim-mensal-de-acompanhamento-da-industria-de-gas-natural",
          f"{DGN}/2026/boletim-mensal-de-acompanhamento-da-industria-de-gas-natural"]:
    r = get(u)
    found = page_links(r, r"dgn/|\.pdf|\.xlsx|boletim", 70)
    for f, t in found:
        if re.search(r"dgn/20\d\d/?$|dgn/[^/]*boletim[^/]*/?$", f) and f.rstrip("/") not in (u.rstrip("/"),):
            out("    -> sub folder", f)

out("\n######## 2. MME Observatorio public Power BI reports ########")
OBS = f"{MME}/assuntos/observatorio-de-minas-e-energia/petroleo-gas-e-biocombustiveis"
for u in [f"{OBS}/outras-informacoes-analiticas/outras-informacoes-analiticas", f"{OBS}/mercado-brasileiro-de-combustiveis",
          f"{MME}/assuntos/observatorio-de-minas-e-energia/energia-eletrica",
          f"{MME}/assuntos/observatorio-de-minas-e-energia/transicao-energetica"]:
    r = get(u)
    page_links(r, r"observatorio-de-minas-e-energia/.*(gas|g%C3%A1s|termel|boletim)", 30)
    for i, tok in enumerate(embeds(r)):
        powerbi(tok, f"{u.rsplit('/', 1)[-1]}#{i}")

out("\n######## 3. ANP gas datasets ########")
ANP_DA = "https://www.gov.br/anp/pt-br/centrais-de-conteudo/dados-abertos"
for ds in ["dados-consolidados-movimentacao-de-gas-natural-em-gasodutos-de-transporte", "comercializacao-de-gas-natural",
           "importacoes-e-exportacoes", "producao-de-petroleo-e-gas-natural-por-estado-e-localizacao"]:
    out(f"\n--- {ds}")
    r = get(f"{ANP_DA}/{ds}")
    files = page_links(r, r"\.(csv|xlsx?|zip|json)$", 60)
    # newest / most relevant files only
    pick = [f for f, t in files if re.search(r"gas|g%C3%A1s|gn|import|2026|2025", f + t, re.I)] or [f for f, t in files]
    for f in pick[-3:] if ds != "comercializacao-de-gas-natural" else pick:
        if f.lower().endswith((".csv", ".xlsx", ".xls")):
            dump_csv(f)

out("\n######## 4. ANP painel dinamico movimentacao gas ########")
for u in ["https://www.gov.br/anp/pt-br/centrais-de-conteudo/paineis-dinamicos-da-anp/painel-dinamico-de-movimentacao-de-gas-natural-em-gasodutos-de-transporte",
          "https://www.gov.br/anp/pt-br/assuntos/movimentacao-estocagem-e-comercializacao-de-gas-natural/acompanhamento-do-mercado-de-gas-natural"]:
    r = get(u)
    page_links(r, r"\.(csv|xlsx?|pdf)$|boletim|painel-dinamico", 30)
    for i, tok in enumerate(embeds(r)):
        powerbi(tok, f"anp-{u.rsplit('/', 1)[-1][:30]}#{i}")

out("\n######## 5. EPE consumo de gas natural ########")
r = get("https://www.epe.gov.br/pt/areas-de-atuacao/petroleo-gas-e-biocombustiveis/consumo-de-g%c3%a1s-natural")
page_links(r, r"g.s|gas|consumo|painel|xlsx|pdf|dashboard", 40)
for i, tok in enumerate(embeds(r)):
    powerbi(tok, f"epe#{i}")

out("\n######## 6. MME boletins mensais de energia ########")
r = get(f"{MME}/assuntos/secretarias/sntep/publicacoes/boletins-mensais-de-energia")
fl = page_links(r, r"boletim|boletins-mensais", 40)
out("\ndone")
