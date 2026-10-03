"""
One-off probe (prints only): find the LNG IMPORTS by country and REGASIFICATION capacity by country tables in
GIIGNL's annual report, for splitting the coal-to-gas switching chart into gas that existing import terminals can
land vs gas needing new regas. Prints page numbers and text of pages that look like those tables.
"""
import os, re, sys
from io import BytesIO
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "americas"))
from GIIGNL_CONTRACTED_VS_SPOT import fetch_pdf_bytes  # noqa: E402
import pdfplumber

PAT = re.compile(r"regasification capacity|nominal capacity|send-?out capacity|receiving terminals|imports by (country|market)|LNG imports in 20\d\d|importing (countries|markets)", re.I)
for year in (2025, 2024):
    content = fetch_pdf_bytes(year)
    if not content:
        print(f"== {year}: no report"); continue
    with pdfplumber.open(BytesIO(content)) as pdf:
        print(f"== {year}: {len(pdf.pages)} pages")
        hits = []
        for i, p in enumerate(pdf.pages):
            t = p.extract_text() or ""
            if PAT.search(t):
                hits.append(i)
                head = " | ".join(l.strip() for l in t.splitlines()[:4])
                print(f"  p{i + 1}: {len(t)} chars; matches {sorted(set(m.group(0).lower() for m in PAT.finditer(t)))[:5]}; {head[:160]}")
        # full text of the pages most likely to be country tables: many country names + numbers
        countries = re.compile(r"\b(Japan|Korea|China|India|Taiwan|Thailand|Pakistan|Bangladesh|Spain|France|Italy|Germany|Netherlands|Brazil|Chile|Argentina|Colombia|Turkey|Philippines|Vietnam|Singapore|Malaysia|Indonesia)\b")
        scored = sorted(hits, key=lambda i: -len(countries.findall(pdf.pages[i].extract_text() or "")))[:6]
        for i in sorted(scored):
            print(f"\n----- {year} page {i + 1} -----")
            print((pdf.pages[i].extract_text() or "")[:4500])
    if year == 2025 and content:
        break
