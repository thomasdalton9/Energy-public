"""Cameroon discovery stage 6: PV bilan pages 1-8 text (peak load / sales / generation summary)."""
import io, requests, pdfplumber
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120 Safari/537.36"}
B = "https://arsel-cm.org/wp-content/uploads/2026/02/"
for f in ["PV-de-validation-du-bilan-energetique-Mai-2025.pdf", "PV-de-validation-du-bilan-e-nergetique-Janvier-2025.pdf"]:
    r = requests.get(B+f, headers=H, timeout=(8,60))
    with pdfplumber.open(io.BytesIO(r.content)) as p:
        print("\n=== ", f, len(p.pages), flush=True)
        for i in range(4, 9):
            print(f"--- page {i+1}"); print((p.pages[i].extract_text() or "[none]")[:1800], flush=True)
