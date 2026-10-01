"""
One-off probe (not in MASTER_SOUTH_AMERICA.py): find CAMMESA data on
Argentina's hydro reservoirs (levels / stored volume / inflows) for the
Comahue basin (El Chocon, Piedra del Aguila, Alicura...), Yacyreta and
Salto Grande, to add a hydro summary next to argentina_generation_mix.py.

Checks, all on CAMMESA's public document API (no key):
  1. Every file inside one day's PROGRAMACION_DIARIA ZIP (the report
     argentina_generation_mix.py already downloads) - name, size, header
     and first rows - looking for reservoir/level/flow tables.
  2. Other daily report codes ('nemo') that may carry reservoir data -
     PARTE_POST_OPERATIVO and similar - listing what each returns.
Outputs go to ar_discovery_* files next to this script.
"""

print("STARTING", flush=True)

import datetime as dt
import io
import json
import zipfile

import requests

LOOKUP_URL = "https://api.cammesa.com/pub-svc/public/findDocumentosByNemoRango"
ATTACHMENT_URL = "https://api.cammesa.com/pub-svc/public/findAttachmentByNemoId"
TIME_FMT = "%Y-%m-%dT%H:%M:%S.000Z"
HEADERS = {"User-Agent": "gas-demand-scripts/1.0"}
DAY = dt.date.today() - dt.timedelta(days=2)
NEMOS = ["PROGRAMACION_DIARIA", "PARTE_POST_OPERATIVO", "POST_OPERATIVO_DIARIO", "PARTE_DIARIO",
         "INFORME_SINTETICO_DIARIO", "RESULTADOS_OPERACION_DIARIA", "HIDRAULICIDAD", "COTAS_CAUDALES"]


def find(nemo, day):
    params = {"fechadesde": day.strftime(TIME_FMT), "fechahasta": (day + dt.timedelta(days=1)).strftime(TIME_FMT),
              "nemo": nemo}
    r = requests.get(LOOKUP_URL, params=params, headers=HEADERS, timeout=60)
    print(f"\n==== {nemo}: status {r.status_code}, {len(r.content):,} bytes", flush=True)
    if not r.ok:
        return []
    try:
        docs = r.json()
    except ValueError:
        print("  not JSON:", r.text[:300], flush=True)
        return []
    if not isinstance(docs, list):
        print("  response:", str(docs)[:300], flush=True)
        return []
    for d in docs[:5]:
        print(f"  doc id={d.get('id')} titulo={d.get('titulo')!r} fecha={d.get('fecha')} "
              f"adjuntos={[a.get('id') for a in d.get('adjuntos', [])][:12]}", flush=True)
    return docs


def show_zip(content, tag):
    try:
        zf = zipfile.ZipFile(io.BytesIO(content))
    except zipfile.BadZipFile:
        print(f"  {tag}: not a zip ({content[:80]!r})", flush=True)
        return
    for info in zf.infolist():
        data = zf.read(info.filename)
        text = data.decode("utf-8-sig", errors="replace")
        lines = text.splitlines()
        print(f"\n  --- {info.filename} ({info.file_size:,} bytes, {len(lines)} lines)", flush=True)
        for line in lines[:6]:
            print(f"    {line[:300]}", flush=True)
        low = text.lower()
        if any(k in low for k in ("cota", "embalse", "caudal", "volumen", "chocon", "piedra del aguila", "yacyreta")):
            print("    ^^ mentions cota/embalse/caudal/volumen - saved", flush=True)
            with open(f"ar_discovery_{tag}_{info.filename.replace('/', '_')}", "wb") as f:
                f.write(data)


# Round 1 (Sep-2026): PROGRAMACION_DIARIA has no reservoir tables, and
# PARTE_POST_OPERATIVO is HTML pages plus an Access .mdb of dispatch and
# reserves - no lake levels either. So round 2 skips CAMMESA and looks at
# AIC (aic.gob.ar), which runs the Comahue lakes (Chocon, Piedra del
# Aguila, Alicura, ...) and publishes their operation.
ROUND1 = False
import re
from urllib.parse import urljoin

AIC_PAGES = ["https://www.aic.gob.ar/", "https://www.aic.gob.ar/sitio/hidrologia",
             "https://www.aic.gob.ar/sitio/embalses", "https://www.aic.gob.ar/sitio/operacion-embalses"]
seen = set()
queue = []  # round 2 done - see round 3 below
while queue and len(seen) < 25:
    url = queue.pop(0)
    if url in seen:
        continue
    seen.add(url)
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=60)
    except requests.RequestException as e:
        print(f"\n==== {url}: FAILED {type(e).__name__}: {e}", flush=True)
        continue
    ctype = r.headers.get("Content-Type", "")
    print(f"\n==== {url}: {r.status_code} {ctype} {len(r.content):,} bytes", flush=True)
    if not r.ok:
        continue
    if "html" not in ctype:
        name = "ar_discovery_aic_" + re.sub(r"[^A-Za-z0-9]+", "_", url.split("/")[-1])[:50]
        open(name, "wb").write(r.content)
        print(f"  saved -> {name}", flush=True)
        continue
    text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
    for m in list(re.finditer(r"cota|volumen|embalse|hm3|hm³", text, re.I))[:4]:
        print(f"  TEXT ...{text[max(0, m.start() - 120):m.start() + 200]}...", flush=True)
    for href, label in re.findall(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', r.text, re.S | re.I):
        label = re.sub(r"<[^>]+>|\s+", " ", label).strip()
        full = urljoin(url, href)
        if "aic.gob.ar" not in full:
            continue
        if re.search(r"embals|cota|hidrol|operac|caudal|parte|diari|datos|xls|csv|pdf", full + " " + label, re.I):
            print(f"  LINK {label[:60]!r} -> {full}", flush=True)
            if re.search(r"embals|cota|operac|parte|diari|xls|csv", full + " " + label, re.I) and not full.lower().endswith(".pdf"):
                queue.append(full)

# Round 3: round 2 found AIC's reservoir list (aic.gob.ar/embalses) with a
# 'Ver Detalle' page per lake (embalses-detalle?a=N&z=...). Open each and
# print what it holds (level, volume, flows) plus any data links/scripts.
listing = requests.get("https://www.aic.gob.ar/embalses", headers={"User-Agent": "Mozilla/5.0"}, timeout=60)
details = sorted(set(re.findall(r'href=["\']([^"\']*embalses-detalle\?a=\d+[^"\'#]*)', listing.text)))
print(f"\n#### {len(details)} AIC reservoir detail pages", flush=True)
for k, href in enumerate(details):
    url = urljoin("https://www.aic.gob.ar/embalses", href)
    r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=60)
    text = re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", r.text, flags=re.S))
    start = text.find("Aprovechamientos")
    print(f"\n==== {url}: {r.status_code} {len(r.content):,} bytes", flush=True)
    print(f"  TEXT {text[start:start + 1800] if start >= 0 else text[:1800]}", flush=True)
    for src in re.findall(r'<script[^>]+src=["\']([^"\']+)', r.text)[:10]:
        print(f"  SCRIPT {src}", flush=True)
    for m in re.findall(r'(https?://[^"\'\s]+(?:json|api|csv|xls)[^"\'\s]*)', r.text)[:10]:
        print(f"  DATAURL {m}", flush=True)
    if k == 0:
        open("ar_discovery_aic_detail_sample.html", "w", encoding="utf-8").write(r.text)

for nemo in NEMOS if ROUND1 else []:
    docs = find(nemo, DAY)
    for doc in docs[:1]:
        for att in doc.get("adjuntos", [])[:6]:
            r = requests.get(ATTACHMENT_URL, params={"attachmentId": att["id"], "docId": doc["id"],
                                                     "nemo": doc.get("nemo", nemo)}, headers=HEADERS, timeout=120)
            print(f"  attachment {att['id']}: status {r.status_code}, {r.headers.get('Content-Type')}, "
                  f"{len(r.content):,} bytes", flush=True)
            if r.ok and r.content[:2] == b"PK":
                show_zip(r.content, nemo.lower())
            elif r.ok:
                name = f"ar_discovery_{nemo.lower()}_{att['id']}".replace("/", "_")
                with open(name, "wb") as f:
                    f.write(r.content)
                print(f"  saved -> {name}", flush=True)

print("\nDONE", flush=True)
