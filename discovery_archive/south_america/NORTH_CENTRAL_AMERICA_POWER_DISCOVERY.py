"""
One-off probe: where do the northern Central American grid operators (and
the regional operator EOR) serve generation by fuel/technology in a form a
script can download?

  Guatemala   AMM  https://www.amm.org.gt/
  El Salvador UT   https://www.ut.com.sv/  (+ estadistico.ut.com.sv)
  Honduras    ODS  https://www.ods.org.hn/
  Belize      BEL  https://www.bel.com.bz/
  Regional    EOR  https://www.enteoperador.org/

For each site: fetch a few start pages, follow links whose text/URL looks
like operations/statistics/reports up to a small page budget, and print
every data-file link (xls/xlsx/csv/json/zip/pdf), every form, iframe and
any script URL that looks like an API. Also tries a few known endpoints
directly (AMM GraficaPW, UT OperacionDiaria).

Usage: python3 CENTRAL_AMERICA_POWER_DISCOVERY.py [site ...]
"""

print("STARTING", flush=True)

import re
import sys
from urllib.parse import urljoin, urlparse

import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36",
           "Accept-Language": "es-ES,es;q=0.9,en;q=0.8"}

SITES = {
    "amm": ["https://www.amm.org.gt/", "https://www.amm.org.gt/portal/",
            "https://wl12.amm.org.gt/GraficaPW/graficaCombustible",
            "https://wl12.amm.org.gt/GraficaPW/"],
    "ut": ["https://www.ut.com.sv/", "https://estadistico.ut.com.sv/",
           "https://estadistico.ut.com.sv/OperacionDiaria.aspx"],
    "ods": ["https://www.ods.org.hn/", "https://ods.org.hn/"],
    "bel": ["https://www.bel.com.bz/"],
    "eor": ["https://www.enteoperador.org/", "https://enteoperador.org/"],
}
FOLLOW = re.compile(r"informe|report|estad|operac|despacho|posdespacho|generac|produc|datos|data|energ|"
                    r"mercado|publicac|boletin|transparen|resultado|tecnolog|recurso|diari|mensual|"
                    r"anual|histor|graf|power|statistic|annual|dashboard|indicador|mapa", re.I)
SKIP = re.compile(r"facebook|twitter|linkedin|youtube|instagram|whatsapp|mailto:|tel:|javascript:|#$|"
                  r"wp-json|xmlrpc|\.(jpg|jpeg|png|gif|svg|webp|css|ico|mp4)(\?|$)", re.I)
DATA = re.compile(r"\.(xlsx?|xlsm|csv|json|zip|pdf|ods)(\?|$)|download|descarga|getfile|attachment", re.I)
API = re.compile(r"""["'](https?://[^"']*(?:api|json|ashx|asmx|svc|Grafica|service|odata|powerbi|"""
                 r"""tableau|arcgis|reportes)[^"']*)["']""", re.I)
BUDGET = 40


def get(url, **kw):
    try:
        r = requests.get(url, headers=HEADERS, timeout=40, **kw)
        return r
    except requests.RequestException as e:
        print(f"  ERR {url}: {type(e).__name__}: {str(e)[:200]}", flush=True)
        return None


def crawl(name, starts):
    print(f"\n######## {name.upper()} ########", flush=True)
    hosts = {urlparse(u).netloc.replace("www.", "") for u in starts}
    queue, seen, data_links, apis = list(starts), set(), {}, set()
    pages = 0
    while queue and pages < BUDGET:
        url = queue.pop(0)
        if url in seen:
            continue
        seen.add(url)
        r = get(url)
        pages += 1
        if r is None:
            continue
        ctype = r.headers.get("content-type", "")
        print(f"\n== [{r.status_code}] {url} -> {r.url} ({ctype}, {len(r.content):,} B)", flush=True)
        if "html" not in ctype and "text" not in ctype and "json" not in ctype:
            continue
        html = r.text
        title = re.search(r"<title[^>]*>(.*?)</title>", html, re.S | re.I)
        print(f"   title: {title.group(1).strip()[:120] if title else ''}", flush=True)
        if "json" in ctype:
            print("   JSON:", html[:1500], flush=True)
            continue
        for m in re.finditer(r"<form[^>]*>", html, re.I):
            print("   FORM", m.group(0)[:250], flush=True)
        for m in re.finditer(r"<iframe[^>]*src=[\"']([^\"']+)", html, re.I):
            src = urljoin(r.url, m.group(1))
            print("   IFRAME", src, flush=True)
            if src not in seen:
                queue.append(src)
        for m in API.finditer(html):
            apis.add(m.group(1))
        for m in re.finditer(r"<script[^>]*src=[\"']([^\"']+)", html, re.I):
            s = m.group(1)
            if not re.search(r"jquery|bootstrap|wp-includes|google|gtag|elementor|wp-content/plugins", s, re.I):
                print("   SCRIPT", urljoin(r.url, s), flush=True)
        for m in re.finditer(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", html, re.I | re.S):
            href, text = m.group(1).strip(), re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()
            if SKIP.search(href):
                continue
            link = urljoin(r.url, href)
            if DATA.search(link):
                if link not in data_links:
                    data_links[link] = text[:80]
                continue
            host = urlparse(link).netloc.replace("www.", "")
            if (host in hosts or any(host.endswith(h) for h in hosts)) and \
                    (FOLLOW.search(link) or FOLLOW.search(text)) and link not in seen:
                queue.append(link)
        # pages that mention generation by technology
        text = re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", html, flags=re.S))
        for m in re.finditer(r".{0,80}(generaci[oó]n por|por tipo de|por recurso|por tecnolog|"
                             r"posdespacho|hidroel[eé]ctric|geot[eé]rm|bunker|carb[oó]n).{0,80}", text, re.I):
            print("   TEXT", m.group(0)[:200], flush=True)
            break
    print(f"\n-- {name}: {len(data_links)} data links", flush=True)
    for link, text in list(data_links.items())[:150]:
        print(f"   DATA {link}  [{text}]", flush=True)
    print(f"-- {name}: API-like URLs in pages", flush=True)
    for a in sorted(apis)[:60]:
        print("   API", a, flush=True)
    print(f"-- {name}: unvisited queue ({len(queue)}):", flush=True)
    for q in queue[:60]:
        print("   Q", q, flush=True)


def amm_known():
    print("\n######## AMM known endpoints ########", flush=True)
    for url, params in [
        ("https://wl12.amm.org.gt/GraficaPW/graficaCombustible", {}),
        ("https://wl12.amm.org.gt/GraficaPW/graficaCombustible", {"dt": "15/01/2024"}),
        ("https://wl12.amm.org.gt/GraficaPW/graficaAreaScada", {}),
    ]:
        r = get(url, params=params)
        if r is not None:
            print(f"[{r.status_code}] {r.url} {r.headers.get('content-type')} {len(r.content):,} B", flush=True)
            print(r.text[:3000], flush=True)


def ut_known():
    print("\n######## UT known endpoints ########", flush=True)
    r = get("https://estadistico.ut.com.sv/OperacionDiaria.aspx")
    if r is not None:
        print(f"[{r.status_code}] {r.url} {len(r.content):,} B", flush=True)
        html = r.text
        for m in re.finditer(r"<(input|select)[^>]*>", html, re.I):
            print("  ", m.group(0)[:200], flush=True)
        for m in re.finditer(r".{0,200}(Hidr|Geot|T[eé]rmic|Biomas|Solar|E[oó]lic).{0,300}", html):
            print("   SNIP", m.group(0)[:500], flush=True)
            break


if __name__ == "__main__":
    want = sys.argv[1:] or list(SITES)
    if "amm" in want:
        amm_known()
    if "ut" in want:
        ut_known()
    for name in want:
        crawl(name, SITES[name])
    print("DONE", flush=True)
