"""
Discovery probe (manual, GitHub Actions only - the editing sandbox blocks these hosts): is there raw data on gas
use by sector (industrial / GNV / residential / commercial) for Jamaica, Dominican Republic, Puerto Rico, Ecuador,
Panama and El Salvador?  It only PRINTS what it finds; nothing is written to the data workbooks.

1. Jamaica: newest 'Jamaica Energy Statistics' PDF (MSET). Prints every table page that mentions natural gas,
   bauxite or alumina, so the sector table (a web search suggests bauxite/alumina and cement natural gas rows
   exist in the 2023/2024 editions) can be read and a parser written from the real layout.
2. Dominican Republic: datos.gob.do CKAN catalogue and MEM statistics pages, searched for gas natural / GNV /
   industrial datasets.
3. Puerto Rico: EIA natural gas consumption by end use for Puerto Rico.
4. Ecuador / Panama / El Salvador: ministry statistics pages (links only).
"""
import re
import sys

import requests

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}


def get(url, **kw):
    try:
        r = requests.get(url, headers=UA, timeout=60, **kw)
        print(f"{r.status_code} {url} ({len(r.content):,} bytes)", flush=True)
        return r
    except requests.RequestException as e:
        print(f"FAIL {url}: {type(e).__name__}", flush=True)


def links(url, pat):
    r = get(url)
    if r is not None and r.ok:
        for h in sorted(set(re.findall(r"href=[\"']([^\"']+)[\"']", r.text))):
            if re.search(pat, h, re.I):
                print("   link:", h, flush=True)


def jamaica():
    print("\n== Jamaica ==", flush=True)
    r = get("https://www.mset.gov.jm/document-category/statistics-data/")
    pdfs = re.findall(r"href=[\"']([^\"']+ENERGY-STATISTICS-\d{4}[^\"']*\.pdf)[\"']", r.text, re.I) if r is not None else []
    pdfs = pdfs or ["https://www.mset.gov.jm/wp-content/uploads/2021/07/JAMAICA-ENERGY-STATISTICS-2024.pdf"]
    pdf = get(pdfs[0])
    if pdf is None or not pdf.ok:
        return
    import pymupdf
    doc = pymupdf.open(stream=pdf.content, filetype="pdf")
    for i, page in enumerate(doc):
        text = page.get_text()
        if re.search(r"natural gas|bauxite|alumina", text, re.I) and re.search(r"TABLE", text):
            print(f"--- page {i + 1} ---\n" + "\n".join(t for t in text.splitlines() if t.strip())[:3500], flush=True)


def dominican():
    print("\n== Dominican Republic ==", flush=True)
    for q in ("gas natural", "GNV", "hidrocarburos consumo"):
        r = get("https://datos.gob.do/api/3/action/package_search", params={"q": q, "rows": 20})
        if r is not None and r.ok:
            for p in r.json()["result"]["results"]:
                print("   dataset:", p["name"], "|", p["title"][:80], "|", [x.get("format") for x in p["resources"]][:6])
    for u in ("https://mem.gob.do/transparencia/estadisticas-del-sector-energetico/",
              "https://mem.gob.do/hidrocarburos/", "https://mem.gob.do/estadisticas/"):
        links(u, r"xlsx|xls|pdf|estadist|balance|gas")


def puerto_rico():
    print("\n== Puerto Rico (EIA) ==", flush=True)
    for s in ("ng_cons_sum_dcu_spr_m", "ng_cons_sum_dcu_spr_a"):
        r = get(f"https://www.eia.gov/dnav/ng/hist/{s}.htm")
        if r is not None and r.ok:
            print(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))[:1500])


def others():
    print("\n== Ecuador / Panama / El Salvador ==", flush=True)
    for u, pat in (("https://www.controlrecursosyenergia.gob.ec/", r"balance|estadist|gas"),
                   ("https://www.recursosyenergia.gob.ec/", r"balance|estadist|gas"),
                   ("https://www.energia.gob.pa/", r"estadist|balance|hidrocarb"),
                   ("https://www.energia.gob.pa/estadisticas/", r"xls|pdf|gas|hidrocarb"),
                   ("https://www.dgehm.gob.sv/", r"estadist|hidrocarb|gas")):
        links(u, pat)


if __name__ == "__main__":
    for fn in (jamaica, dominican, puerto_rico, others):
        try:
            fn()
        except Exception as e:  # noqa: BLE001 - discovery: keep going
            print("ERROR", fn.__name__, type(e).__name__, e, file=sys.stderr)
