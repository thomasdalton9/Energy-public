"""Dump the sector tables of recent Gestor del Mercado reports, whose
layout changed around Aug 2025 (industrial drops, compressors jump), to
fix COLOMBIA_GAS.py's parser. Not reachable from the editing sandbox."""
import sys, os, io, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import requests
import pdfplumber
import COLOMBIA_GAS as C

URLS = [
    "https://www.bmcbec.com.co/sites/default/files/2025-08/Informe%20Mensual%202025%20Julio.pdf",
    "https://www.bmcbec.com.co/sites/default/files/2025-09/Informe%20Mensual%202025%20Agosto.pdf",
    "https://www.bmcbec.com.co/sites/default/files/2026-01/Informe%20Mensual%202025%20diciembre.pdf",
    "https://www.bmcbec.com.co/sites/default/files/2026-09/Informe%20Mensual%202026%20Agosto.pdf",
]
for u in URLS:
    print("\n====================", u, flush=True)
    c = requests.get(u, headers=C.H, timeout=C.T).content
    with pdfplumber.open(io.BytesIO(c)) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if not re.search(r"Segmento|Residencial|demanda promedio", text, re.I):
                continue
            print(f"--- page {i + 1}", flush=True)
            for ln in text.splitlines()[:60]:
                print("   ", ln[:200])
            for tb in page.extract_tables()[:4]:
                print(f"   TABLE {len(tb)} rows")
                for row in tb[:20]:
                    print("     ", [str(x)[:14].replace("\n", "|") if x else "" for x in row][:16])
