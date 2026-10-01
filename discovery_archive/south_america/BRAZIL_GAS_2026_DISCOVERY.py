"""
Brazil gas demand by segment after MME's annex froze (2025-06) and its PDF
bulletins stopped (Sep-2025). Looks for anything public that is current to
Aug/Sep 2026, in priority order:
  A. MME gas bulletin: index/year pages, candidate 2026 paths, site search,
     annex folder (newer xlsx names), Last-Modified of the old annex.
  B. MME/SNTEP "Boletim Mensal de Energia" (has a gas section).
  C. ANP open data: dataset index, gas-related dataset files (comercializacao,
     importacoes, producao, movimentacao), ANP site search for gas panels.
  D. ABEGAS: newest consumption posts.
  E. EPE: gas market publications / open data.
  F. MME Observatorio pages: which are public, which redirect to login.
Prints compact, greppable lines. Run from GitHub Actions (sandbox is blocked).
"""
import io
import re
import sys
from urllib.parse import urljoin, quote_plus

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
     "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8"}
T = (10, 60)
S = requests.Session()
S.headers.update(H)


def out(*a):
    print(*a, flush=True)


def get(u, **kw):
    try:
        r = S.get(u, timeout=T, allow_redirects=True, **kw)
        gated = "require_login" in r.url or "credentials_cookie_auth" in r.url
        out(f"GET {u}\n    -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)}B"
            f"{' REDIRECT ' + r.url if r.url != u else ''}{' LOGIN-GATED' if gated else ''}"
            f" lastmod={r.headers.get('last-modified', '-')}")
        return r
    except Exception as e:
        out(f"GET {u}\n    -> ERR {type(e).__name__} {str(e)[:150]}")
        return None


def links(r, pat=None, limit=80):
    if r is None or r.status_code != 200 or "html" not in r.headers.get("content-type", ""):
        return []
    found = []
    for href, txt in re.findall(r'<a[^>]+href="([^"#]+)"[^>]*>(.*?)</a>', r.text, re.S | re.I):
        txt = re.sub(r"<[^>]+>|\s+", " ", txt).strip()
        full = urljoin(r.url, href)
        if pat is None or re.search(pat, full + " " + txt, re.I):
            found.append((full, txt[:110]))
    found = list(dict.fromkeys(found))
    for f, t in found[:limit]:
        out(f"      {f} | {t}")
    if len(found) > limit:
        out(f"      ... {len(found) - limit} more")
    return found


def text_of(r):
    t = re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", r.text, flags=re.S | re.I)
    return re.sub(r"\s+", " ", t)


def plone_search(site, q, pat=None):
    for path in ("@@search", "search"):
        u = f"{site}/{path}?SearchableText={quote_plus(q)}&sort_on=Date&sort_order=reverse"
        r = get(u)
        if r is not None and r.status_code == 200:
            links(r, pat or r"\.(pdf|xlsx?|csv)|boletim|painel|gas|g%C3%A1s", 40)
            return


def dump_table(content, name):
    """Print sheets/columns/tail of a CSV or xlsx."""
    import pandas as pd
    try:
        if content[:2] == b"PK" or name.lower().endswith((".xlsx", ".xls")):
            xl = pd.ExcelFile(io.BytesIO(content))
            out(f"    sheets: {xl.sheet_names[:30]}")
            for s in xl.sheet_names[:12]:
                df = xl.parse(s, header=None, nrows=400)
                out(f"    --- sheet {s!r} shape {df.shape}")
                out("      " + df.head(8).to_string(max_colwidth=22, max_cols=16)[:1800].replace("\n", "\n      "))
                hits = [i for i in range(len(df)) if re.search(r"industri|residen|comerc|automot|gera..o|cogera|termel|demanda|consumo",
                                                                " ".join(map(str, df.iloc[i, :4])), re.I)]
                for i in hits[:20]:
                    out("      ROW", i, " | ".join(map(str, df.iloc[i, :6]))[:200])
            return
        txt = None
        for enc in ("utf-8-sig", "latin-1"):
            try:
                txt = content.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        sep = ";" if txt[:2000].count(";") > txt[:2000].count(",") else ","
        df = pd.read_csv(io.StringIO(txt), sep=sep, engine="python", on_bad_lines="skip")
        out(f"    shape {df.shape}; cols {list(df.columns)[:40]}")
        out("      " + df.tail(5).to_string(max_colwidth=25, max_cols=20)[:2000].replace("\n", "\n      "))
        for c in df.columns:
            if df[c].dtype == object and 1 < df[c].nunique() <= 40:
                out(f"    distinct {c}: {sorted(map(str, df[c].dropna().unique()))[:40]}")
            if re.search(r"ano|mes|m.s|data|per.odo|date", str(c), re.I):
                v = df[c].dropna().astype(str)
                if len(v):
                    out(f"    {c}: min {v.min()} max {v.max()}")
    except Exception as e:
        out(f"    parse ERR {type(e).__name__} {str(e)[:200]}")


def fetch_and_dump(u):
    r = get(u)
    if r is not None and r.status_code == 200 and len(r.content) > 100:
        dump_table(r.content, u.split("?")[0])
    return r


# --------------------------------------------------------------------------------------------
out("\n################ A. MME gas bulletin ################")
MME = "https://www.gov.br/mme/pt-br"
BASE = (f"{MME}/assuntos/secretarias/petroleo-gas-natural-e-biocombustiveis/publicacoes-1/"
        "boletim-mensal-de-acompanhamento-da-industria-de-gas-natural")
r = get(BASE)
links(r, r"boletim|anexo|\.pdf|\.xlsx|20(2[5-7])")
r = get(BASE + "/2025")
links(r, r"\.pdf|\.xlsx")
for cand in ["/2026", "/boletins-2026", "/boletim-2026", "/2026-1", "/anexos", "/anexos/historico-balanco-boletim.xlsx",
             "/anexos/historico-balanco-boletim-2026.xlsx", "/anexos/historico-balanco-boletim-1.xlsx"]:
    r = get(BASE + cand)
    if r is not None and r.status_code == 200 and "html" in r.headers.get("content-type", ""):
        links(r, r"\.pdf|\.xlsx|boletim")
# the parent publications folder may hold a renamed 2026 bulletin series
PUB = f"{MME}/assuntos/secretarias/petroleo-gas-natural-e-biocombustiveis/publicacoes-1"
r = get(PUB)
links(r, r"gas|g%C3%A1s|boletim")
plone_search(MME, "boletim mensal de acompanhamento da industria de gas natural 2026")
plone_search(MME, "historico balanco boletim gas natural")
plone_search(MME, "demanda de gás natural por segmento 2026")

out("\n################ B. MME Boletim Mensal de Energia ################")
plone_search(MME, "boletim mensal de energia 2026", r"\.pdf|\.xlsx|boletim")
for u in [f"{MME}/assuntos/secretarias/sntep/publicacoes/boletins-mensais-de-energia",
          f"{MME}/assuntos/secretarias/spte/publicacoes/boletim-mensal-de-energia",
          f"{MME}/assuntos/secretarias/sntep/publicacoes/boletim-mensal-de-energia"]:
    r = get(u)
    found = links(r, r"2026|\.xlsx|\.pdf")
    xl = [f for f, _ in found if f.lower().endswith((".xlsx", ".xls"))]
    if xl:
        fetch_and_dump(xl[0])
        break
    sub = [f for f, t in found if "2026" in f]
    if sub:
        r2 = get(sub[0])
        f2 = links(r2, r"\.xlsx|\.pdf")
        xl = [f for f, _ in f2 if f.lower().endswith((".xlsx", ".xls"))]
        if xl:
            fetch_and_dump(xl[-1])
        break

out("\n################ C. ANP ################")
ANP = "https://www.gov.br/anp/pt-br"
r = get(f"{ANP}/centrais-de-conteudo/dados-abertos")
idx = links(r, r"dados-abertos/", 150)
gas_pages = [f for f, t in idx if re.search(r"gas|g%C3%A1s|gás|importa|produ|moviment|comerci", f + t, re.I)]
for p in gas_pages[:14]:
    r = get(p)
    files = links(r, r"\.(csv|xlsx?|zip|json)(\?|$)", 25)
    for f, t in files[-2:]:
        if f.lower().endswith((".csv", ".xlsx", ".xls")):
            fetch_and_dump(f)
plone_search(ANP, "painel dinâmico gás natural", r"painel|powerbi|gas|g%C3%A1s")
plone_search(ANP, "boletim mensal gás natural consumo segmento", r"\.pdf|\.xlsx|boletim|painel")
for u in [f"{ANP}/assuntos/movimentacao-estocagem-e-comercializacao-de-gas-natural/acompanhamento-do-mercado-de-gas-natural",
          f"{ANP}/centrais-de-conteudo/paineis-dinamicos-da-anp",
          f"{ANP}/centrais-de-conteudo/publicacoes/boletins-anp/boletins/boletim-mensal-da-producao-de-petroleo-e-gas-natural"]:
    r = get(u)
    links(r, r"painel|powerbi|\.xlsx|\.csv|\.pdf|gas|g%C3%A1s|2026", 60)
    if r is not None and r.status_code == 200:
        for m in sorted(set(re.findall(r"app\.powerbi\.com/view\?r=[A-Za-z0-9=_-]+", r.text)))[:10]:
            out("      POWERBI", m)

out("\n################ D. ABEGAS ################")
for u in ["https://www.abegas.org.br/estatisticas-de-consumo", "https://www.abegas.org.br/?s=consumo+de+g%C3%A1s",
          "https://www.abegas.org.br/arquivos/category/releases"]:
    r = get(u)
    links(r, r"consumo|estat|\.pdf|\.xlsx", 30)

out("\n################ E. EPE ################")
for u in ["https://www.epe.gov.br/pt/publicacoes-dados-abertos/dados-abertos",
          "https://www.epe.gov.br/pt/publicacoes-dados-abertos/publicacoes?search=g%C3%A1s%20natural",
          "https://dashboard.epe.gov.br/apps/gas-natural/", "https://www.epe.gov.br/pt/areas-de-atuacao/petroleo-gas-e-biocombustiveis"]:
    r = get(u)
    links(r, r"g.s|gas|natural|mensal|painel|dashboard", 40)

out("\n################ F. MME Observatorio ################")
OBS = f"{MME}/assuntos/observatorio-de-minas-e-energia"
for u in [OBS, f"{OBS}/petroleo-gas-e-biocombustiveis"]:
    r = get(u)
    for f, t in links(r, r"observatorio-de-minas-e-energia/", 60)[:40]:
        if re.search(r"gas|g%C3%A1s|petroleo", f, re.I) and f != u:
            r2 = get(f)
            if r2 is not None and r2.status_code == 200:
                emb = sorted(set(re.findall(r"(app\.powerbi\.com/view\?r=[^\"'&]+|[a-z0-9.-]*(?:qlik|tableau|looker)[a-z0-9./-]*)", r2.text)))
                out(f"      embeds: {emb or 'none'}")
out("\ndone")
