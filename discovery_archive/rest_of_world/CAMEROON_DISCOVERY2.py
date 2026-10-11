"""Cameroon discovery stage 2: open ARSEL production PDFs, PV bilan, SNH key-data pages."""
import re, io, requests, pdfplumber
from urllib.parse import urljoin
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120 Safari/537.36"}
B = "https://arsel-cm.org/wp-content/uploads/2026/"
def get(u, t=(8,60)):
    return requests.get(u, headers=H, timeout=t)
def dump(u, maxpages=40, maxchars=1800):
    try:
        r = get(u); print("\n=== PDF", u, r.status_code, len(r.content), r.headers.get("Last-Modified"), flush=True)
        with pdfplumber.open(io.BytesIO(r.content)) as p:
            print("pages", len(p.pages))
            for i, pg in enumerate(p.pages[:maxpages]):
                print(f"--- page {i+1}"); print((pg.extract_text() or "[no text]")[:maxchars], flush=True)
    except Exception as e:
        print("ERR", u, e, flush=True)
dump(B+"01/PRODUCTION-Tableau-de-Bord.pdf")
dump(B+"01/PRODUCTION-Energies-produites.pdf")
dump(B+"01/Transport-TB.pdf", maxpages=6)
dump(B+"02/PV-de-validation-du-bilan-energetique-Mai-2025.pdf", maxpages=4)
dump(B+"02/PV-de-validation-du-bilan-en-ergetique-Decembre-2025.pdf", maxpages=4)
for u in ["https://snh.cm/index.php/chiffres-cles/", "https://snh.cm/index.php/en/key-data/", "https://www.snh.cm/envira/production/", "https://snh.cm/index.php/activites/exploration-production/"]:
    try:
        r = get(u, (8,30)); print("\n###", u, r.status_code, len(r.text))
        for m in sorted(set(re.findall(r'https?:[^"\'\s<>]+\.(?:pdf|xlsx?|jpg|png)', r.text.replace("\\/","/"), re.I))):
            if 'wp-content' in m and not re.search(r'logo|icon|-\d+x\d+', m): print("  FILE", m)
        txt = re.sub(r'<[^>]+>', ' ', re.sub(r'<(script|style).*?</\1>', '', r.text, flags=re.S)); txt = re.sub(r'\s+', ' ', txt)
        i = txt.lower().find('production'); print(txt[max(0,i-200):i+1500])
    except Exception as e: print("ERR", u, e, flush=True)
