"""Cameroon discovery stage 5: SNH quarterly petroleum data PDFs."""
import io, requests, pdfplumber
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120 Safari/537.36"}
for u in ["https://www.snh.cm/wp-content/uploads/2026/10/Donnees_Petrolieres_2eme_trimestre_2026_fr.pdf",
          "https://snh.cm/wp-content/uploads/2026/06/Donnees-petrolieres-2025.pdf",
          "https://snh.cm/wp-content/uploads/2026/06/Graphiques_SNH_VA__Production.pdf"]:
    try:
        r = requests.get(u, headers=H, timeout=(8,60)); print("\n=== PDF", u, r.status_code, len(r.content), r.headers.get("Last-Modified"), flush=True)
        with pdfplumber.open(io.BytesIO(r.content)) as p:
            print("pages", len(p.pages))
            for i, pg in enumerate(p.pages[:8]):
                t = pg.extract_text() or "[no text]"; print(f"--- page {i+1} imgs={len(pg.images)}"); print(t[:2200], flush=True)
                for tb in pg.extract_tables()[:2]: print("TABLE", tb[:14], flush=True)
    except Exception as e: print("ERR", u, e, flush=True)
