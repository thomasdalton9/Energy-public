"""
Round 14 (Oct-2026), Belize. PROBE12 found that PUC Belize (puc.bz) posts
BEL's monthly 'Application to PUC for COPA Tariff Review' (Cost of Power
Adjustment) and the PUC decision, e.g.
https://www.puc.bz/wp-content/uploads/2026/08/BEL-Application-to-PUC-for-COPA-Tariff-Review-September-2026.pdf
A COPA filing should list the month's energy purchases by supplier (BECOL
hydro, BELCOGEN bagasse, CFE imports, solar, BEL diesel/gas turbine...).
Here: the text of one application, and how far back the monthly filings go
(WordPress search + media library JSON).

Usage: python3 NORTH_CENTRAL_AMERICA_POWER_PROBE14.py puc
"""

print("STARTING", flush=True)

import io
import json
import re
import sys

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36"}
S = requests.Session()
S.headers.update(H)


def get(url, **kw):
    try:
        r = S.get(url, timeout=90, **kw)
        print(f"[{r.status_code}] {r.url[:200]} {r.headers.get('content-type')} {len(r.content):,} B", flush=True)
        return r
    except requests.RequestException as e:
        print(f"ERR {url}: {e}", flush=True)
        return None


def pdf_text(content, max_pages=12):
    import pdfplumber
    out = []
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for i, p in enumerate(pdf.pages[:max_pages]):
            out.append((i + 1, p.extract_text() or "", p.extract_tables()))
    return out


def puc():
    for url in ["https://www.puc.bz/wp-content/uploads/2026/08/BEL-Application-to-PUC-for-COPA-Tariff-Review-September-2026.pdf",
                "https://www.puc.bz/wp-content/uploads/2026/08/Decision-and-Order-COPA-Tariff-for-BEL-September-2026-w-Summary.pdf"]:
        r = get(url)
        if r is None or r.content[:4] != b"%PDF":
            continue
        for n, txt, tables in pdf_text(r.content):
            if re.search(r"MWh|GWh|kWh|BECOL|BELCOGEN|CFE|Hydro|Solar|purchase", txt, re.I):
                print(f"  --- page {n}\n{txt[:2500]}", flush=True)
                for t in tables[:3]:
                    for row in t[:25]:
                        print("     T", row, flush=True)
    # history of COPA filings: WordPress media search
    found = {}
    for page in range(1, 8):
        r = get("https://www.puc.bz/wp-json/wp/v2/media", params={"search": "COPA", "per_page": 100, "page": page})
        if r is None or not r.ok:
            break
        try:
            items = r.json()
        except ValueError:
            break
        if not items:
            break
        for it in items:
            found[it.get("source_url")] = it.get("date")
    for u, d in sorted(found.items(), key=lambda x: x[1] or ""):
        print(f"  MEDIA {d} {u}", flush=True)
    for q in ["COPA", "cost of power", "energy purchases", "fuel and purchased power"]:
        r = get("https://www.puc.bz/", params={"s": q})
        if r is not None:
            for m in re.finditer(r"href=\"(https://www\.puc\.bz/[^\"]+)\"[^>]*>([^<]{5,120})<", r.text):
                if re.search(r"copa|power|tariff|purchase", m.group(1) + m.group(2), re.I):
                    print(f"   SEARCH[{q}] {m.group(1)[:160]} | {m.group(2).strip()[:90]}", flush=True)


if __name__ == "__main__":
    {"puc": puc}[sys.argv[1]]()
    print("DONE", flush=True)
