"""Cameroon discovery stage 4: PV bilan pages 5-12 (peak/sales), other months layout check, SNH media."""
import re, io, json, requests, pdfplumber
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120 Safari/537.36"}
B = "https://arsel-cm.org/wp-content/uploads/2026/02/"
def pdf(u):
    return pdfplumber.open(io.BytesIO(requests.get(u, headers=H, timeout=(8,60)).content))
with pdf(B+"PV-de-validation-du-bilan-energetique-Mai-2025.pdf") as p:
    for i in range(4, 13):
        t = p.pages[i].extract_text() or "[none]"
        print(f"--- MAI page {i+1}"); print(t[:1700], flush=True)
for f in ["PV-de-validation-du-bilan-e-nergetique-Janvier-2025.pdf", "PV-de-validation-du-bilan-en-ergetique-Decembre-2025.pdf"]:
    with pdf(B+f) as p:
        print("\n=== ", f, len(p.pages))
        for i, pg in enumerate(p.pages):
            t = pg.extract_text() or ""
            if "INJECTIONS DES CENTRALES" in t or "BILAN DES INJECTIONS DES CENTRALES" in t or "TOTAL INJECTION" in t:
                print(f"--- page {i+1}"); print(t[:1800], flush=True)
for u in ["https://snh.cm/wp-json/wp/v2/media?per_page=100&search=production", "https://snh.cm/wp-json/wp/v2/media?per_page=100&mime_type=application/pdf",
          "https://snh.cm/index.php/rapports-annuels/"]:
    try:
        r = requests.get(u, headers=H, timeout=(8,40)); print("\n###", u, r.status_code, len(r.text))
        for m in sorted(set(re.findall(r'https?:[^"\'\s<>]+\.(?:pdf|xlsx?)', r.text.replace("\\/","/"), re.I))): print("  FILE", m)
    except Exception as e: print("ERR", u, e)
