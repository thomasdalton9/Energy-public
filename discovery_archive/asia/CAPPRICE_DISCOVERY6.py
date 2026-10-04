"""
Capacity round 6: run asia/PHILIPPINES_DOE_CAPACITY.py twice to a temp file (the second run should find no new
release) and print the sheets; print the head of every PDF page (titles / scopes).
"""
import io
import os
import subprocess
import sys
import tempfile

import pandas as pd
import pdfplumber
import requests

pd.set_option("display.width", 250)
p = os.path.join(tempfile.mkdtemp(), "philippines_power_capacity.xlsx")
for _ in range(2):
    subprocess.run([sys.executable, "asia/PHILIPPINES_DOE_CAPACITY.py", "--out", p], timeout=900)
if os.path.exists(p):
    xl = pd.ExcelFile(p)
    for s in ("Monthly", "Dependable", "Release"):
        if s in xl.sheet_names:
            print(f"--- {s}", flush=True)
            print(xl.parse(s).to_string(), flush=True)
    g = xl.parse("By grid")
    print(g.groupby(["measure", "scope", "grid"]).size().to_string(), flush=True)
PDF = ("https://d24qbtp4vooyzi.cloudfront.net/api/media/file/2%20%20Installed%20and%20Dependable%20Capacity%20per%20"
       "Grid%20and%20per%20technology%202003%202025.pdf?prefix=dev%2Fmedia")
r = requests.get(PDF, timeout=(20, 120), headers={"User-Agent": "Mozilla/5.0"})
print("HEAD-like headers:", {k: v for k, v in r.headers.items() if k.lower() in ("last-modified", "etag")})
with pdfplumber.open(io.BytesIO(r.content)) as pdf:
    for i, pg in enumerate(pdf.pages):
        print(f"=== page {i + 1}: {(pg.extract_text() or '')[:260]!r}", flush=True)
