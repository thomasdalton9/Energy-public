"""
One-off probe (prints only): find the LNG IMPORTS by country and REGASIFICATION capacity by country tables in
GIIGNL's annual report (2025 edition = 2024 data), for splitting the coal-to-gas switching chart into gas existing
import terminals can land vs gas needing new regas.
Round 1 found: imports by country on page 16. Round 2: the regas section (page 46 on) - page heads, then the full
text of pages that look like per-country / per-terminal capacity tables.
"""
import os, re, sys
from io import BytesIO
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "americas"))
from GIIGNL_CONTRACTED_VS_SPOT import fetch_pdf_bytes  # noqa: E402
import pdfplumber

DEC = re.compile(r"\b\d+[.,]\d\b")

content = fetch_pdf_bytes(2025)
with pdfplumber.open(BytesIO(content)) as pdf:
    texts = [p.extract_text() or "" for p in pdf.pages]
    for i in range(44, len(texts)):
        t = texts[i]
        n_mtpa, n_dec = len(re.findall(r"MTPA", t)), len(re.findall(DEC, t))
        print(f"p{i + 1}: {len(t)} chars, {n_mtpa} MTPA, {n_dec} decimals | " + " | ".join(l.strip() for l in t.splitlines()[:3])[:150])
    print("\n----- page 46 -----\n" + texts[45][:3000])
    tables = [i for i in range(44, len(texts)) if re.search(r"nominal|capacity \(mtpa\)|send-?out|number of (tanks|terminals)|storage capacity", texts[i], re.I)
              and len(DEC.findall(texts[i])) > 25]
    for i in tables[:5]:
        print(f"\n----- page {i + 1} -----\n" + texts[i][:5000])
