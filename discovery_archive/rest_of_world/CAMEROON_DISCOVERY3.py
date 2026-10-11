"""Cameroon discovery stage 3: PV bilan energetique pages 5+, SNH annual reports."""
import re, io, requests, pdfplumber
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120 Safari/537.36"}
B = "https://arsel-cm.org/wp-content/uploads/2026/02/"
r = requests.get(B+"PV-de-validation-du-bilan-energetique-Mai-2025.pdf", headers=H, timeout=(8,60))
with pdfplumber.open(io.BytesIO(r.content)) as p:
    for i, pg in enumerate(p.pages):
        if i < 4: continue
        t = pg.extract_text() or "[no text]"
        print(f"--- page {i+1} chars={len(t)} imgs={len(pg.images)}"); print(t[:1500], flush=True)
        for tb in pg.extract_tables()[:2]: print("TABLE", tb[:12], flush=True)
for u in ["https://snh.cm/index.php/rapports-annuels/", "https://www.snh.cm/rapports-annuels/", "https://snh.cm/index.php/en/annual-reports/", "https://snh.cm/index.php/snh-infos/", "https://www.snh.cm/envira/production/"]:
    try:
        r = requests.get(u, headers=H, timeout=(8,30)); print("\n###", u, r.status_code, len(r.text))
        for m in sorted(set(re.findall(r'https?:[^"\'\s<>]+\.(?:pdf|xlsx?)', r.text.replace("\\/","/"), re.I))): print("  FILE", m)
    except Exception as e: print("ERR", u, e)
