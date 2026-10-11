"""Cameroon discovery: ARSEL / ENEO / SONATREL / SNH pages and PDFs. Prints links and PDF text."""
import re, io, sys, requests
from urllib.parse import urljoin
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120 Safari/537.36"}
PAGES = [
 "https://arsel-cm.org/", "https://arsel-cm.org/tableau-de-bord_production/",
 "https://arsel-cm.org/bilan-energetique-mensuel/", "https://arsel-cm.org/energies-produites/",
 "https://arsel-cm.org/tableau-de-bord/", "https://arsel-cm.org/publications/",
 "https://arsel-cm.org/statistiques/", "https://arsel-cm.org/wp-json/wp/v2/pages?per_page=100&search=bord",
 "https://arsel-cm.org/wp-json/wp/v2/media?per_page=100&mime_type=application/pdf",
 "https://www.eneocameroon.cm/", "https://www.eneocameroon.cm/index.php/fr/rapports-annuels",
 "https://www.sonatrel.cm/", "https://www.snh.cm/", "https://www.snh.cm/index.php/en/",
 "https://www.snh.cm/index.php/fr/production",
]
pdfs = []
for u in PAGES:
    try:
        r = requests.get(u, headers=H, timeout=(8,25))
        print("\n###", u, r.status_code, len(r.text))
        links = set(re.findall(r'https?://[^"\'\s<>\\]+\.(?:pdf|xlsx?|csv)', r.text.replace("\\/", "/"), re.I))
        for m in re.findall(r'href=["\']([^"\']+)["\']', r.text):
            if re.search(r'\.(pdf|xlsx?|csv)$', m, re.I): links.add(urljoin(u, m))
        for l in sorted(links): print("  FILE", l); pdfs.append(l)
        if "wp-json" in u: print(r.text[:1500])
        if not links:
            print("  hrefs:", sorted(set(urljoin(u, m) for m in re.findall(r'href=["\']([^"\']+)["\']', r.text)))[:80])
    except Exception as e:
        print("\n###", u, "ERR", e)
try:
    import pdfplumber
except ImportError:
    pdfplumber = None
seen = set()
for l in pdfs:
    if l in seen: continue
    seen.add(l)
    if not re.search(r'arsel|snh|sonatrel|eneo', l, re.I): continue
    try:
        r = requests.get(l, headers=H, timeout=(8,60))
        print("\n=== PDF", l, r.status_code, len(r.content), r.headers.get("Last-Modified"))
        with pdfplumber.open(io.BytesIO(r.content)) as p:
            print("pages", len(p.pages))
            for i, pg in enumerate(p.pages[:6]):
                print(f"--- page {i+1}"); print((pg.extract_text() or "")[:2500])
                for t in pg.extract_tables()[:2]: print("TABLE", t[:15])
    except Exception as e:
        print("ERR", l, e)
