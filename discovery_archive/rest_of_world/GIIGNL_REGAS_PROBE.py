"""
One-off probe (prints only): LNG imports and regasification capacity by country from GIIGNL's annual report
(2025 edition = end-2024 data), for splitting the coal-to-gas switching chart into gas that existing import
terminals can land vs gas needing new regas.
Found: imports by country (net of re-exports) on page 16; the per-terminal regasification table (pages ~52-66)
prints each market's total nominal send-out capacity as "<n> MTPA" under the market name. Round 3 prints every
line of the table pages carrying "MTPA" with the two lines before it, plus page 46 (global summary).
"""
import os, re, sys
from io import BytesIO
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "americas"))
from GIIGNL_CONTRACTED_VS_SPOT import fetch_pdf_bytes  # noqa: E402
import pdfplumber

content = fetch_pdf_bytes(2025)
with pdfplumber.open(BytesIO(content)) as pdf:
    texts = [p.extract_text() or "" for p in pdf.pages]
print("----- page 46 -----\n" + texts[45][:1500])
for i, t in enumerate(texts):
    if "Send-out" not in t or "Nominal" not in t:
        continue
    lines = t.splitlines()
    for j, line in enumerate(lines):
        if re.search(r"\d\s*MTPA", line):
            ctx = " || ".join(l.strip() for l in lines[max(0, j - 2):j + 1])
            print(f"p{i + 1}: {ctx[:220]}")
