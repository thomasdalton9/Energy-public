"""
One-off probe (not in MASTER_SOUTH_AMERICA.py): find where CENACE,
Ecuador's grid operator, serves generation by source and reservoir data
(Mazar, Amaluza/Paute - levels, 'cota') in a form a script can download,
so ECUADOR_CENACE.py can be written against real responses.

Crawls CENACE's public operation pages (Informacion Operativa, the real
dispatch page, the daily programmed dispatch page) up to a small page
budget, printing every form, iframe and data-file / report link, and any
text mentioning cota / embalse / Mazar / produccion. Saves data files it
finds (xls/xlsx/csv/json/pdf) as ec_discovery_* next to this script.
"""

print("STARTING", flush=True)

import re
from urllib.parse import urljoin, urlparse

import requests
import urllib3

# CENACE's server doesn't send its intermediate certificate, so Python's
# check fails ("unable to get local issuer certificate") where a browser
# fetches the missing piece itself. These are public, read-only pages, so
# the check is skipped for cenace.gob.ec only.
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36"}
START = [
    "https://www.cenace.gob.ec/info-operativa/InformacionOperativa.htm",
    "https://www.cenace.gob.ec/elaboracion-del-despacho-real/",
    "https://www.cenace.gob.ec/emision-del-despacho-economico-diario-programado-del-sistema-nacional-interconectado/",
    "https://www.cenace.gob.ec/informes-mensuales-de-gestion/",
]
FOLLOW = re.compile(r"info-operativa|operativ|despacho|produc|generac|embals|cota|hidro|reporte|diari|bosni|simem|"
                    r"\.aspx|\.htm|\.php", re.I)
DATA = re.compile(r"\.(xlsx?|csv|json|pdf|zip)(\?|$)", re.I)
MENTIONS = re.compile(r"cota|embalse|mazar|amaluza|paute|producci[oó]n|hidr[aá]ulic|t[eé]rmic", re.I)
BUDGET = 30

# Round 2 (after round 1, Sep-2026): InformacionOperativa.htm carries the
# numbers itself - daily production by source as Plotly charts - so save
# it whole, plus each Plotly trace's name and a few values, and any text on
# reservoirs (cota / embalse / Mazar / Amaluza).
import json
page = requests.get(START[0], headers=HEADERS, timeout=60, verify=False)
open("ec_discovery_informacion_operativa.html", "w", encoding="utf-8").write(page.text)
print(f"saved InformacionOperativa.htm ({len(page.text):,} chars)", flush=True)
for m in re.finditer(r'Plotly\.newPlot\(\s*["\']([^"\']+)["\']\s*,\s*(\[.*?\])\s*,', page.text, re.S):
    try:
        traces = json.loads(m.group(2))
    except ValueError:
        print(f"  PLOT {m.group(1)}: (traces not plain JSON, {len(m.group(2)):,} chars)", flush=True)
        continue
    for t in traces:
        x, y = t.get("x") or [], t.get("y") or []
        print(f"  PLOT {m.group(1)} trace={t.get('name')!r} type={t.get('type')} n={len(y)} "
              f"x0={x[:2]} y0={y[:3]}", flush=True)
text = re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", page.text, flags=re.S))
for m in list(re.finditer(r"cota|embalse|mazar|amaluza|paute|nivel", text, re.I))[:12]:
    print(f"  RESERVOIR-TEXT ...{text[max(0, m.start() - 120):m.start() + 220]}...", flush=True)
for tab in re.findall(r'(?:href|data-url|src)=["\']([^"\']*(?:Diari|Mensual|Anual|Embals|Cota|Hidro)[^"\']*)["\']', page.text, re.I)[:20]:
    print(f"  TAB-LINK {tab}", flush=True)
START = []  # round 1's crawl is done

seen, queue, saved = set(), list(START), 0
while queue and len(seen) < BUDGET:
    url = queue.pop(0)
    if url in seen:
        continue
    seen.add(url)
    try:
        r = requests.get(url, headers=HEADERS, timeout=60, verify="cenace.gob.ec" not in urlparse(url).netloc)
    except requests.RequestException as e:
        print(f"\n==== {url}: FAILED {type(e).__name__}: {e}", flush=True)
        continue
    ctype = r.headers.get("Content-Type", "")
    print(f"\n==== {url}: {r.status_code} {ctype} {len(r.content):,} bytes", flush=True)
    if not r.ok:
        continue
    if "html" not in ctype:
        if saved < 12:
            name = "ec_discovery_" + re.sub(r"[^A-Za-z0-9.]+", "_", urlparse(url).path.split("/")[-1])[-60:]
            open(name, "wb").write(r.content)
            saved += 1
            print(f"  saved -> {name}", flush=True)
        continue
    html = r.text
    for form in re.findall(r"<form[^>]*>", html, re.I):
        print(f"  FORM {form[:200]}", flush=True)
    for src in re.findall(r'<iframe[^>]+src=["\']([^"\']+)', html, re.I):
        full = urljoin(url, src)
        print(f"  IFRAME {full}", flush=True)
        queue.append(full)
    text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html))
    for m in list(MENTIONS.finditer(text))[:5]:
        print(f"  TEXT ...{text[max(0, m.start() - 100):m.start() + 180]}...", flush=True)
    for href, label in re.findall(r'<a[^>]+href=["\']([^"\'#]+)["\'][^>]*>(.*?)</a>', html, re.S | re.I):
        label = re.sub(r"<[^>]+>|\s+", " ", label).strip()
        full = urljoin(url, href)
        if DATA.search(full):
            print(f"  DATA {label[:60]!r} -> {full}", flush=True)
            if saved < 12 and ("cenace" in full) and not full.lower().endswith(".pdf"):
                queue.insert(0, full)
        elif "cenace" in full and FOLLOW.search(full + " " + label):
            print(f"  LINK {label[:60]!r} -> {full}", flush=True)
            queue.append(full)

print("\nDONE", flush=True)
