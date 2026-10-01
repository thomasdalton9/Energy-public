"""
Where can gas demand BY SECTOR (power, industrial, residential,
vehicle, etc.) be pulled for Colombia and Brazil, from 2021? Probes:
  Colombia: datos.gov.co catalog (Socrata), the Gestor del Mercado de
            Gas (Bolsa Mercantil, bmcbec.com.co) pages - links + XHR.
  Brazil:   ABEGAS statistics pages, MME's monthly gas bulletin page,
            ANP open data, dados.gov.br catalog.
Prints every data-looking link (csv/xlsx/xls/json/api/pdf) and every
JSON/CSV network response a page makes when loaded in a browser.
Not reachable from the editing sandbox.
"""
import re

import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
TIMEOUT = (10, 45)
LINK_RE = re.compile(r'(?:href|src|data-url)=["\']([^"\']+\.(?:csv|xlsx?|json|pdf|zip)(?:\?[^"\']*)?)["\']', re.I)
KEYWORDS = re.compile(r"gas|gás|demanda|consumo|segment|sector|setor|boletim|informe|estad", re.I)


def out(*a):
    print(*a, flush=True)


def get(url, **kw):
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kw)
        out(f"GET {r.url} -> {r.status_code} {r.headers.get('content-type', '')[:40]} {len(r.content)}B")
        return r
    except Exception as e:
        out(f"GET {url} -> ERR {type(e).__name__}: {e}")
        return None


def links(url, keep=KEYWORDS):
    r = get(url)
    if r is None or r.status_code != 200:
        return
    found = sorted(set(LINK_RE.findall(r.text)))
    anchors = sorted(set(m for m in re.findall(r'href=["\']([^"\'#]+)["\']', r.text) if keep.search(m)))
    for f in found:
        out("   file:", f)
    for a in anchors[:60]:
        out("   page:", a)


out("################ COLOMBIA ################")
out("\n== datos.gov.co catalog search ==")
for q in ["gas natural demanda", "gas natural consumo sector", "demanda gas natural termoelectrico"]:
    r = get("https://api.us.socrata.com/api/catalog/v1", params={"domains": "www.datos.gov.co", "q": q, "limit": 15})
    if r is not None and r.ok:
        for res in r.json().get("results", []):
            x = res["resource"]
            out(f"   [{x['id']}] {x['name']} | updated {x.get('data_updated_at', '')[:10]} | cols: {x.get('columns_name', [])[:12]}")

out("\n== Gestor del Mercado de Gas (BMC) ==")
for u in ["https://www.bmcbec.com.co/", "https://www.bmcbec.com.co/informes/informes-diarios",
          "https://www.bmcbec.com.co/informes/informes-mensuales", "https://www.bmcbec.com.co/informes",
          "https://www.bmcbec.com.co/informes/declaracion-de-produccion", "https://www.bmcbec.com.co/estadisticas"]:
    out(f"\n-- {u}")
    links(u)

out("\n################ BRAZIL ################")
for u in ["https://www.abegas.org.br/", "https://www.abegas.org.br/estatisticas",
          "https://www.abegas.org.br/arquivos/estatisticas",
          "https://www.gov.br/mme/pt-br/assuntos/secretarias/petroleo-gas-natural-e-biocombustiveis/publicacoes-1/boletim-mensal-de-acompanhamento-da-industria-de-gas-natural",
          "https://www.gov.br/anp/pt-br/centrais-de-conteudo/dados-abertos"]:
    out(f"\n-- {u}")
    links(u)

out("\n== dados.gov.br CKAN search ==")
for q in ["gas natural consumo", "consumo gas natural segmento"]:
    r = get("https://dados.gov.br/api/publico/conjuntos-dados", params={"nomeConjuntoDados": q, "pagina": 1})
    if r is not None and r.ok:
        try:
            for d in (r.json() if isinstance(r.json(), list) else r.json().get("registros", []))[:15]:
                out("   ", d.get("titulo") or d.get("title"), "|", d.get("nome") or d.get("name"))
        except Exception as e:
            out("   parse err", e, r.text[:300])

out("\n################ BROWSER CAPTURE ################")
from playwright.sync_api import sync_playwright

PAGES = ["https://www.bmcbec.com.co/", "https://www.bmcbec.com.co/informes",
         "https://www.abegas.org.br/estatisticas"]
with sync_playwright() as p:
    b = p.chromium.launch()
    for url in PAGES:
        ctx = b.new_context(user_agent=HEADERS["User-Agent"])
        page = ctx.new_page()
        seen = []
        page.on("response", lambda resp: seen.append((resp.status, resp.headers.get("content-type", ""), resp.url))
                if any(k in resp.headers.get("content-type", "") for k in ("json", "csv", "excel", "spreadsheet", "pdf"))
                or re.search(r"\.(csv|xlsx?|pdf)(\?|$)|/api/", resp.url, re.I) else None)
        out(f"\n== {url}")
        try:
            page.goto(url, timeout=60000, wait_until="load")
            page.wait_for_timeout(6000)
            out("   title:", page.title())
            for a in page.locator("a").all()[:400]:
                try:
                    t = (a.inner_text() or "").strip().replace("\n", " ")
                    h = a.get_attribute("href") or ""
                    if KEYWORDS.search(t + " " + h):
                        out(f"   a: {t[:70]!r} -> {h}")
                except Exception:
                    pass
        except Exception as e:
            out("   goto error:", type(e).__name__, str(e)[:200])
        for st, ct, u in seen:
            out(f"   RESP {st} {ct[:35]} {u}")
        ctx.close()
    b.close()
