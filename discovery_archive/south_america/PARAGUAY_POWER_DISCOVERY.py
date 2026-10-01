"""
One-off probe for Paraguay power data (generation by plant, own consumption,
exports to Brazil and Argentina, capacity) before writing PARAGUAY_POWER.py.

  crawl     - ANDE (ande.gov.py), Itaipu Binacional (itaipu.gov.br / .gov.py),
              Entidad Binacional Yacyreta (eby.gov.py / eby.org.ar), VMME
              (ssme.gov.py / mopc.gov.py): fetch each seed page, follow
              same-site links whose URL or text looks statistical (two levels,
              capped), and print every document link (pdf/xls/xlsx/csv/zip/json)
              and keyword link found.
  datos     - datos.gov.py CKAN search (ANDE, Itaipu, Yacyreta, energia).
  cammesa   - CAMMESA post-operation MDB for a few days: every GENERADORES row
              for Yacyreta (YACY*) and every INTERCAMBIO='S' (import) node, with
              that day's MWh - how CAMMESA counts Yacyreta and Paraguay.

Usage: python3 PARAGUAY_POWER_DISCOVERY.py [crawl] [datos] [cammesa] [fetch URL ...]
       (manual workflow: discovery_archive/workflows/paraguay_power_discovery.yml)
"""
import csv
import io
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile
from collections import defaultdict
from urllib.parse import urljoin, urlparse

import requests

HERE = os.path.dirname(os.path.abspath(__file__))
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0 Safari/537.36", "Accept-Language": "es-PY,es;q=0.9,pt;q=0.8,en;q=0.7"}
DOC = re.compile(r"\.(pdf|xlsx?|csv|zip|json|ods|xlsm)(\?|$)", re.I)
KEY = re.compile(r"estad|compil|generac|gera[cç]|energ|produ|demanda|operac|opera[cç]|datos|dados|informe|"
                 r"reporte|relat|balance|memoria|anuario|boletin|bolet|despacho|export|cesi|intercamb|"
                 r"transparen|indicador|potencia|capacidad|cota|nivel|embalse|reservat|series|historic", re.I)
SEEDS = [
    "https://www.ande.gov.py/",
    "https://www.ande.gov.py/estadisticas.php",
    "https://www.ande.gov.py/compilacion_estadistica.php",
    "https://www.ande.gov.py/datos_abiertos.php",
    "https://www.ande.gov.py/interna.php?id=1196",
    "https://www.itaipu.gov.br/energia/geracao",
    "https://www.itaipu.gov.br/energia/integracao-ao-sistema-brasileiro",
    "https://www.itaipu.gov.br/",
    "https://www.itaipu.gov.py/es/energia/generacion",
    "https://www.itaipu.gov.py/",
    "https://www.eby.gov.py/",
    "https://www.eby.org.ar/",
    "https://www.ssme.gov.py/vmme/",
    "https://www.ssme.gov.py/",
    "https://www.mopc.gov.py/",
    "https://sien.ssme.gov.py/",
]


def get(s, url, **kw):
    try:
        r = s.get(url, timeout=kw.pop("timeout", 40), allow_redirects=True, **kw)
        return r
    except requests.RequestException as e:
        print(f"  ERR {url}: {type(e).__name__}: {str(e)[:150]}", flush=True)
        return None


def links(html, base):
    out = []
    for m in re.finditer(r"<a\b[^>]*href=[\"']([^\"'#]+)[\"'][^>]*>(.*?)</a>", html, re.I | re.S):
        href = urljoin(base, m.group(1).strip())
        text = re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()[:90]
        out.append((href, text))
    # also bare document URLs in scripts/json
    for m in re.finditer(r"[\"'](https?://[^\"'\s]+\.(?:pdf|xlsx?|csv|zip|json))[\"']", html, re.I):
        out.append((m.group(1), "(script)"))
    return out


QUIET = False
ANDE_SEEDS = ["https://www.ande.gov.py/", "https://www.ande.gov.py/estadisticas.php",
              "https://www.ande.gov.py/transparencia.php", "https://www.ande.gov.py/datos_abiertos.php"]


def crawl(seeds=None, cap=30):
    s = requests.Session()
    s.headers.update(UA)
    seen, docs = set(), {}
    per_site = defaultdict(int)
    queue = [(u, 0) for u in (seeds or SEEDS)]
    while queue:
        url, depth = queue.pop(0)
        if url in seen:
            continue
        seen.add(url)
        host = urlparse(url).netloc
        if per_site[host] >= cap:
            continue
        per_site[host] += 1
        r = get(s, url)
        if r is None:
            continue
        ctype = r.headers.get("content-type", "")
        title = re.search(r"<title[^>]*>(.*?)</title>", r.text, re.I | re.S) if "html" in ctype else None
        print(f"\n[{depth}] {r.status_code} {url} -> {r.url} ({ctype[:30]}, {len(r.content):,} B) "
              f"{(title.group(1).strip()[:80] if title else '')}", flush=True)
        if r.status_code != 200 or "html" not in ctype:
            continue
        for href, text in links(r.text, r.url):
            h = urlparse(href).netloc
            if SKIP.search(href):
                continue
            if DOC.search(href):
                if href not in docs:
                    docs[href] = text
                    print(f"    DOC {href}  [{text[:70]}]")
            elif KEY.search(href) or KEY.search(text):
                if href not in seen and not QUIET:
                    print(f"    key {href}  [{text[:70]}]")
                if depth < 2 and (h == host or h.endswith(host.replace("www.", ""))):
                    queue.append((href, depth + 1))
    print(f"\n==== {len(docs)} document links")
    for d, t in sorted(docs.items()):
        print(f"  {d}  [{t}]")


def datos():
    s = requests.Session()
    s.headers.update(UA)
    for q in ("ANDE", "energia electrica", "Itaipu", "Yacyreta", "generacion", "VMME"):
        for base in ("https://www.datos.gov.py/api/3/action/package_search", "https://datos.gov.py/api/3/action/package_search"):
            r = get(s, base, params={"q": q, "rows": 30})
            if r is None:
                continue
            print(f"\n== {base} q={q}: {r.status_code} {r.headers.get('content-type','')[:30]}")
            try:
                res = r.json()["result"]["results"]
            except Exception:  # noqa: BLE001
                print("   ", r.text[:300])
                continue
            for p in res:
                print(f"  * {p.get('name')}: {p.get('title')} ({p.get('organization', {}) and p['organization'].get('title')})")
                for rs in p.get("resources", [])[:8]:
                    print(f"      - {rs.get('format')} {rs.get('url')} [{(rs.get('name') or '')[:60]}] {rs.get('last_modified') or rs.get('created')}")
            break


def mdb_rows(path, table):
    out = subprocess.run(["mdb-export", path, table], capture_output=True, text=True, timeout=300)
    if out.returncode != 0:
        raise RuntimeError(out.stderr[:200])
    return list(csv.DictReader(io.StringIO(out.stdout)))


def cammesa():
    sys.path.insert(0, os.path.join(HERE, "..", "..", "south_america"))
    import ARGENTINA_POWER_DAILY as AR  # noqa: E402
    import datetime as dt
    s = AR.A.make_session()
    for day in (dt.date(2021, 6, 15), dt.date(2024, 1, 15), dt.date(2025, 6, 15), dt.date(2026, 8, 15)):
        docs = AR.month_docs(s, day.replace(day=1))
        hit = docs.get(day.strftime("PO%y%m%d.zip"))
        if not hit:
            print(day, "no PO file")
            continue
        doc, att = hit
        r = s.get(AR.A.ATTACHMENT_URL, params={"attachmentId": att["id"], "docId": doc["id"], "nemo": doc.get("nemo") or AR.NEMO}, timeout=180)
        zf = zipfile.ZipFile(io.BytesIO(r.content))
        name = next(n for n in zf.namelist() if n.lower().endswith(".mdb"))
        with tempfile.NamedTemporaryFile(suffix=".mdb", delete=False) as f:
            f.write(zf.read(name))
            path = f.name
        gens = {g["GRUPO"]: g for g in mdb_rows(path, "GENERADORES")}
        e = defaultdict(float)
        for v in mdb_rows(path, "VALORES_GENERADORES"):
            try:
                e[v["GRUPO"]] += float(v["ENERGIA"] or 0)
            except ValueError:
                pass
        print(f"\n== {day}: {len(gens)} generators; GENERADORES columns: {list(next(iter(gens.values())).keys())}")
        for g, row in sorted(gens.items()):
            if "YACY" in g.upper() or row.get("INTERCAMBIO") == "S" or "PY" in g.upper() or "ITAI" in g.upper():
                print(f"  {g:12s} {e.get(g, 0):10.1f} MWh  {json.dumps(row, ensure_ascii=False)[:300]}")
        # tables list, in case one holds interchange with Paraguay
        out = subprocess.run(["mdb-tables", "-1", path], capture_output=True, text=True)
        print("  tables:", out.stdout.replace("\n", " ")[:1500])
        os.unlink(path)


SKIP = re.compile(r"facebook|twitter|instagram|youtube|linkedin|flickr|whatsapp|mailto|FontSize|login|"
                  r"register|/page/\d|cota-\d|noticias/|/tag/|/author/|wp-json|xmlrpc|\.jpg|\.png", re.I)


PDF_PAT = r"(?i)(GWh|MWh).{0,400}(ANDE|Eletrobras|ENBPar|cedid|SINP|SADI|mensual|enero|janeiro)"
PDF_MAX = 8


def fetch(urls, text_chars=3500):
    """Compact view of each page: unique doc/keyword links (filtered) + stripped text + any html tables.
    A PDF is read with pdfplumber (first pages' text); an xlsx/csv prints its sheets' heads."""
    import pandas as pd
    s = requests.Session()
    s.headers.update(UA)
    for u in urls:
        r = get(s, u, timeout=90)
        if r is None:
            continue
        ct = r.headers.get("content-type", "")
        print(f"\n== {r.status_code} {u} -> {r.url} ({ct[:40]}, {len(r.content):,} B)")
        if r.status_code != 200:
            continue
        if "pdf" in ct or u.lower().endswith(".pdf"):
            try:
                import pdfplumber
                with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                    print(f"   pdf pages: {len(pdf.pages)}")
                    hits = 0
                    for i, pg in enumerate(pdf.pages):
                        t = pg.extract_text() or ""
                        if re.search(PDF_PAT, t, re.S):
                            print(f"   --- page {i + 1}:\n" + t[:1500])
                            hits += 1
                        if hits >= PDF_MAX:
                            break
            except Exception as e:  # noqa: BLE001
                print("   pdf error", e)
            continue
        if re.search(r"sheet|excel|csv|octet", ct) or DOC.search(u):
            try:
                xl = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None)
                for name, df in xl.items():
                    print(f"   sheet {name} {df.shape}\n{df.head(25).to_string()[:2500]}")
            except Exception as e:  # noqa: BLE001
                print("   not a workbook:", e, r.text[:500])
            continue
        seen = set()
        for href, text in links(r.text, r.url):
            if href in seen or SKIP.search(href):
                continue
            seen.add(href)
            if DOC.search(href) or KEY.search(href) or KEY.search(text):
                print(f"    link {href}  [{text[:70]}]")
        try:
            for i, t in enumerate(pd.read_html(io.StringIO(r.text))[:6]):
                print(f"   table {i} {t.shape}\n{t.head(20).to_string()[:2000]}")
        except Exception:  # noqa: BLE001
            pass
        txt = re.sub(r"<script.*?</script>|<style.*?</style>|<nav.*?</nav>|<header.*?</header>|<footer.*?</footer>",
                     " ", r.text, flags=re.S | re.I)
        print("   TEXT:", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", txt))[:text_chars])


def grep(url, pattern, n=40, width=220):
    """Contexts of a regex in a page's raw HTML/JS (find embedded data, API endpoints)."""
    s = requests.Session()
    s.headers.update(UA)
    r = get(s, url, timeout=90)
    if r is None:
        return
    print(f"\n== grep {pattern!r} in {url}: {r.status_code}, {len(r.content):,} B")
    for i, m in enumerate(re.finditer(pattern, r.text, re.I)):
        if i >= n:
            break
        a = max(0, m.start() - width // 2)
        print("   ..." + re.sub(r"\s+", " ", r.text[a:m.end() + width // 2]) + "...")


def wpsearch(base, query, pages=10):
    """WordPress REST search: every post matching `query` (title, date, link, text start)."""
    s = requests.Session()
    s.headers.update(UA)
    print(f"\n== wp search {base} q={query!r}")
    for page in range(1, pages + 1):
        r = get(s, f"{base.rstrip('/')}/wp-json/wp/v2/posts", params={"search": query, "per_page": 100, "page": page,
                                                                      "_fields": "date,link,title,content"})
        if r is None or r.status_code != 200:
            print("   stop:", r.status_code if r is not None else None, (r.text[:200] if r is not None else ""))
            break
        items = r.json()
        for it in items:
            body = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", it["content"]["rendered"]))
            print(f"  {it['date'][:10]} {it['title']['rendered'][:80]} | {it['link']}\n      {body[:400]}")
        if len(items) < 100:
            break


if __name__ == "__main__":
    args = sys.argv[1:] or ["crawl", "datos", "cammesa"]
    for a in [x for x in args if x.startswith("grep=")]:   # grep=URL|REGEX
        url, pat = a[5:].split("|", 1)
        grep(url, pat)
    for a in [x for x in args if x.startswith("wp=")]:     # wp=BASE|QUERY
        base, q = a[3:].split("|", 1)
        wpsearch(base, q)
    args = [x for x in args if not x.startswith(("grep=", "wp="))]
    if "fetch" in args:
        fetch(args[args.index("fetch") + 1:])
        args = args[:args.index("fetch")]
    if "quiet" in args:
        QUIET = True
    if "crawl" in args:
        crawl()
    if "ande" in args:
        crawl(ANDE_SEEDS, cap=60)
    if "datos" in args:
        datos()
    if "cammesa" in args:
        cammesa()
