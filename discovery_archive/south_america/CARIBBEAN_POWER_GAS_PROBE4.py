"""
Round 4 (Oct-2026): monthly gas use for the Dominican Republic and Jamaica,
and Jamaica generation granularity.
  do4   datos.gob.do (CKAN) search, DGA customs statistics, ONE, the CNE data
        repository (datacne.gob.do Dropbox listing API)
  jm4   MSET "Jamaica Energy Statistics 2025" (monthly / quarterly tables?),
        STATIN production statistics, JPS 2024 annual report generation pages
Usage: python3 CARIBBEAN_POWER_GAS_PROBE4.py do4|jm4
"""

print("STARTING", flush=True)

import io
import json
import os
import re
import sys
from urllib.parse import urljoin

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36", "Accept-Language": "es-ES,es;q=0.9,en;q=0.8"}
S = requests.Session()
S.headers.update(H)
OUT = "probe_files"
os.makedirs(OUT, exist_ok=True)


def get(url, save=None, method="GET", **kw):
    try:
        r = S.request(method, url, timeout=90, **kw)
        print(f"[{r.status_code}] {method} {r.url[:200]} {r.headers.get('content-type')} {len(r.content):,} B", flush=True)
        if save:
            with open(os.path.join(OUT, save), "wb") as f:
                f.write(r.content)
        return r
    except requests.RequestException as e:
        print(f"ERR {url}: {type(e).__name__}: {str(e)[:200]}", flush=True)
        return None


def links(r, pat, limit=80):
    seen = set()
    for m in re.finditer(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", r.text, re.I | re.S):
        link = urljoin(r.url, m.group(1))
        txt = re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()
        if link not in seen and re.search(pat, link + " " + txt, re.I):
            seen.add(link)
            if len(seen) <= limit:
                print("   A", link[:220], "|", txt[:90])
    return seen


def pdf_pages(content, name, pattern, maxchars=3000, maxpages=25):
    import pdfplumber
    shown = 0
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        print(f"  {name}: {len(pdf.pages)} pages", flush=True)
        with open(os.path.join(OUT, name + ".txt"), "w") as f:
            for i, p in enumerate(pdf.pages):
                t = p.extract_text() or ""
                f.write(f"\n===== page {i + 1}\n{t}")
                if re.search(pattern, t, re.I) and shown < maxpages:
                    shown += 1
                    print(f"  --- page {i + 1}\n{t[:maxchars]}", flush=True)


def do4():
    for q in ["gas natural", "combustibles", "importacion combustibles", "generacion electrica", "GNL"]:
        r = get("https://datos.gob.do/api/3/action/package_search", params={"q": q, "rows": 20})
        if r is not None and r.ok:
            try:
                for p in r.json()["result"]["results"]:
                    print(f"   PKG {p['name']} | {p.get('title')} | {p.get('organization', {}).get('title')} | "
                          f"modified {p.get('metadata_modified', '')[:10]}")
                    for res in p.get("resources", [])[:6]:
                        print(f"       RES {res.get('format')} {res.get('url', '')[:180]}")
            except Exception as e:
                print("   parse", e, r.text[:300])
    for u in ["https://www.aduanas.gob.do/", "https://www.aduanas.gob.do/estadisticas/",
              "https://www.aduanas.gob.do/datos-abiertos/"]:
        r = get(u, save="dga_" + re.sub(r"\W", "_", u[23:])[:40] + ".html")
        if r is not None and "html" in (r.headers.get("content-type") or ""):
            links(r, r"estad|datos|comercio|import|\.xlsx?|\.csv")
    for u in ["https://www.one.gob.do/datos-y-estadisticas/", "https://www.one.gob.do/"]:
        r = get(u)
        if r is not None and "html" in (r.headers.get("content-type") or ""):
            links(r, r"energ|combust|electric|import|comercio", 40)
    j = get("https://datacne.gob.do/_next/static/chunks/02z4ygpzb.lqf.js")
    if j is not None and j.ok:
        for m in re.finditer(r"/api/dropbox/(list|preview|pdf-proxy)", j.text):
            print("---- JS context\n", j.text[max(0, m.start() - 700): m.start() + 500], flush=True)
    for method, kw in [("GET", {}), ("GET", {"params": {"path": ""}}), ("POST", {"json": {"path": ""}}),
                       ("POST", {"json": {}})]:
        r = get("https://datacne.gob.do/api/dropbox/list", method=method, save=f"datacne_list_{method}.json", **kw)
        if r is not None:
            print("   ", r.text[:3000])


def jm4():
    r = get("https://www.mset.gov.jm/wp-content/uploads/2021/07/JAMAICA-ENERGY-STATISTICS-2025.pdf")
    if r is not None and r.ok:
        pdf_pages(r.content, "jm_energy_stats_2025",
                  r"January|Jan\b|Quarter|Q1|monthly|natural gas|LNG|generation by|net generation", maxpages=30)
    r = get("https://statinja.gov.jm/productionstats.aspx", save="statin_production.html")
    if r is not None:
        t = re.sub(r"<[^>]+>", " ", r.text)
        for m in re.finditer(r"electric", t, re.I):
            print("   ...", re.sub(r"\s+", " ", t[max(0, m.start() - 300): m.start() + 300]))
        links(r, r"electric|energy|datazoa|\.xlsx?|\.pdf", 40)
    r = get("https://www.jpsco.com/wp-content/uploads/2025/04/JPS-AR-2024-MM-FINAL-300425-SPREADS.pdf")
    if r is not None and r.ok:
        pdf_pages(r.content, "jps_ar_2024", r"(LNG|natural gas).*(MWh|GWh|generation)|net generation|fuel mix|"
                                           r"generation mix|MMBtu", maxpages=12)


if __name__ == "__main__":
    {"do4": do4, "jm4": jm4}[sys.argv[1]]()
    print("DONE", flush=True)
