"""
One-off probe (not in MASTER_SOUTH_AMERICA.py): find where ADME, Uruguay's
electricity market administrator, serves generation by source and
reservoir data in a form a script can download, so URUGUAY_ADME.py can be
written against real responses rather than guesses.

Pages checked (all public, no key):
  - pronos.adme.com.uy/gpf.php   "Generacion por fuente" - SCADA generation
    by source, 10-minute and hourly files, with a date-range form
  - adme.com.uy/datosabiertos.html, controlpanel.php, imasd/adme_data.html
    - open-data and operation-data pages

For each page: status, content type, any <form> (action, method, fields)
and every link to a data file (xls/xlsx/ods/csv/json/zip). Then submits
the gpf form for two recent days and saves whatever comes back.
Outputs go to uy_discovery_* files next to this script.

Round 1 (Sep-2026) found it: gpf.php takes fecha_ini/fecha_fin as
dd/mm/yyyy and links /cache/gpf_<a>_<b>_horario.ods (and _diezminutal):
hourly MW per plant group (Salto Grande, Bonete, Baygorria, Palmar, wind,
solar, thermal, biomass, imports, demand) plus per-plant sheets and
exports. Round 2 looks for reservoir series and history depth.
"""

print("STARTING", flush=True)

import datetime as dt
import re
from html.parser import HTMLParser
from urllib.parse import urljoin

import requests

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/124.0 Safari/537.36"}
PAGES = [
    "https://pronos.adme.com.uy/gpf.php",
    "https://adme.com.uy/datosabiertos.html",
    "https://adme.com.uy/controlpanel.php",
    "https://adme.com.uy/imasd/adme_data.html",
    "https://adme.com.uy/pronos2.php",
]
DATA_EXT = re.compile(r"\.(xlsx?|ods|csv|json|zip|txt)(\?|$)", re.I)


class Scan(HTMLParser):
    def __init__(self):
        super().__init__()
        self.forms, self.links, self._form = [], [], None

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "form":
            self._form = {"action": a.get("action"), "method": a.get("method", "get"), "fields": []}
            self.forms.append(self._form)
        elif tag in ("input", "select", "textarea") and self._form is not None:
            self._form["fields"].append({k: a.get(k) for k in ("name", "type", "value", "id") if a.get(k)})
        elif tag == "a" and a.get("href"):
            self.links.append(a["href"])
        elif tag in ("iframe", "script") and a.get("src"):
            self.links.append(a["src"])

    def handle_endtag(self, tag):
        if tag == "form":
            self._form = None


def fetch(label, url, **kw):
    print(f"\n{'=' * 70}\n{label}: {url} {kw.get('params') or kw.get('data') or ''}\n{'=' * 70}", flush=True)
    try:
        r = requests.request(kw.pop("method", "GET"), url, headers=HEADERS, timeout=90, **kw)
    except requests.RequestException as e:
        print(f"  FAILED: {type(e).__name__}: {e}", flush=True)
        return None
    print(f"  status {r.status_code}, {r.headers.get('Content-Type')}, {len(r.content):,} bytes", flush=True)
    return r


def save(r, name):
    ctype = (r.headers.get("Content-Type") or "").lower()
    ext = "html" if "html" in ctype else "json" if "json" in ctype else "xls" if "excel" in ctype or "spreadsheet" in ctype else "bin"
    path = f"uy_discovery_{name}.{ext}"
    with open(path, "wb") as f:
        f.write(r.content)
    print(f"  saved -> {path}", flush=True)
    return path


def scan(r, base):
    s = Scan()
    try:
        s.feed(r.text)
    except Exception as e:
        print(f"  (html parse: {e})", flush=True)
    for f in s.forms:
        print(f"  FORM action={f['action']} method={f['method']} fields={f['fields']}", flush=True)
    data_links = sorted({urljoin(base, l) for l in s.links if DATA_EXT.search(l) or "descarg" in l.lower() or ".php" in l})
    for l in data_links[:80]:
        print(f"  LINK {l}", flush=True)
    return s


ROUND1 = False  # done Sep-2026 - see the header
for i, url in enumerate(PAGES if ROUND1 else []):
    r = fetch(f"page {i}", url)
    if r is not None and r.ok:
        save(r, f"page{i}")
        scan(r, url)

# gpf.php's own form, two recent days, in the date formats it's likely to take
end = dt.date.today() - dt.timedelta(days=1)
start = end - dt.timedelta(days=1)
for label, params in [] if not ROUND1 else [
    ("gpf dd/mm/yyyy", {"fecha_ini": start.strftime("%d/%m/%Y"), "fecha_fin": end.strftime("%d/%m/%Y"), "send": "MOSTRAR"}),
    ("gpf yyyy-mm-dd", {"fecha_ini": start.isoformat(), "fecha_fin": end.isoformat(), "send": "MOSTRAR"}),
]:
    r = fetch(label, "https://pronos.adme.com.uy/gpf.php", params=params)
    if r is not None and r.ok:
        path = save(r, label.split()[1].replace("/", ""))
        s = scan(r, "https://pronos.adme.com.uy/gpf.php")
        # follow the first few data-file links it offers
        for link in [urljoin("https://pronos.adme.com.uy/gpf.php", l) for l in s.links if DATA_EXT.search(l)][:4]:
            f = fetch("  follow", link)
            if f is not None and f.ok:
                save(f, "file_" + re.sub(r"[^A-Za-z0-9]+", "_", link.split("/")[-1])[:40])

# ----------------------------------------------------------
# Round 2 (after round 1 found the gpf .ods files): reservoir series,
# older generation history, and how long a range gpf.php serves.
# ----------------------------------------------------------
for label, url, params in [] if True else [
    ("seriesbonete (Rio Negro operation)", "https://pronos.adme.com.uy/seriesbonete.php", None),
    ("saltogrande_xls (Salto Grande flows)", "https://pronos.adme.com.uy/scripts/saltogrande_xls.php", None),
    ("gpf_historico (pre-2019)", "https://adme.com.uy/gpf_historico.php", None),
    ("gpf one month back in 2019", "https://pronos.adme.com.uy/gpf.php",
     {"fecha_ini": "01/01/2019", "fecha_fin": "31/01/2019", "send": "MOSTRAR"}),
    ("gpf a full year", "https://pronos.adme.com.uy/gpf.php",
     {"fecha_ini": "01/01/2025", "fecha_fin": "31/12/2025", "send": "MOSTRAR"}),
]:
    r = fetch(label, url, params=params) if params else fetch(label, url)
    if r is None or not r.ok:
        continue
    name = re.sub(r"[^a-z0-9]+", "_", label.lower())[:30]
    save(r, "r2_" + name)
    if "html" in (r.headers.get("Content-Type") or ""):
        s = scan(r, url)
        text = re.sub(r"<[^>]+>", " ", r.text)
        print("  TEXT", re.sub(r"\s+", " ", text)[:1500], flush=True)
        for link in [urljoin(url, l) for l in s.links if DATA_EXT.search(l) and "horario" in l][:1]:
            f = fetch("  follow", link)
            if f is not None and f.ok:
                save(f, "r2_file_" + name)

# ----------------------------------------------------------
# Round 3: the Rio Negro plants' SCADA series (seriesbonete.php links
# /cgi-bin/seriescentralhidro.cgi?idCentral=bon&ts=ods&dtIni=<serial>&dtFin=<serial>,
# serials = days since 1899-12-30). Print each .ods sheet's header and a
# few rows, for a short and a long range, per plant.
# ----------------------------------------------------------
import zipfile, io, xml.etree.ElementTree as ET
_T = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
_O = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"


def ods_rows(content):
    root = ET.fromstring(zipfile.ZipFile(io.BytesIO(content)).read("content.xml"))
    out = {}
    for t in root.iter(_T + "table"):
        rows = []
        for r in t.iter(_T + "table-row"):
            row = []
            for c in r:
                if c.tag != _T + "table-cell":
                    continue
                v = c.get(_O + "value") or c.get(_O + "date-value") or ("".join(c.itertext()) or None)
                row += [v] * min(int(c.get(_T + "number-columns-repeated", "1")), 50)
            while row and row[-1] is None:
                row.pop()
            if row:
                rows.append(row)
        out[t.get(_T + "name")] = rows
    return out


def serial(d):
    return (d - dt.date(1899, 12, 30)).days


today = dt.date.today()
for plant in ("bon", "pal", "bay"):
    for label, a, b in [("3 days", today - dt.timedelta(days=3), today),
                        ("since 2019", dt.date(2019, 1, 1), today)]:
        url = (f"https://pronos.adme.com.uy/cgi-bin/seriescentralhidro.cgi?idCentral={plant}&ts=ods"
               f"&dtIni={serial(a)}&dtFin={serial(b)}")
        r = fetch(f"{plant} {label}", url)
        if r is None or not r.ok:
            continue
        try:
            sheets = ods_rows(r.content)
        except Exception as e:
            print(f"  not an ods ({type(e).__name__}): {r.content[:200]!r}", flush=True)
            continue
        if label == "3 days":
            save(r, f"r3_{plant}")
        for name, rows in sheets.items():
            print(f"  SHEET {name}: {len(rows)} rows", flush=True)
            for row in rows[:6] + rows[-2:]:
                print(f"    {row[:14]}", flush=True)

print("\nDONE", flush=True)
