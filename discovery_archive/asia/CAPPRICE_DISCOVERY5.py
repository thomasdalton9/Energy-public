"""
Capacity round 5: text / tables of DOE Philippines '2 Installed and Dependable Capacity per Grid and per
technology 2003-2025' PDF (linked from doe.gov.ph/.../electric-power-industry/2025-power-statistics), and the
energy-statistics index page links for other years.
"""
import io
import re

import pdfplumber
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
PDF = ("https://d24qbtp4vooyzi.cloudfront.net/api/media/file/2%20%20Installed%20and%20Dependable%20Capacity%20per%20"
       "Grid%20and%20per%20technology%202003%202025.pdf?prefix=dev%2Fmedia")


def out(*a):
    print(*a, flush=True)


r = requests.get(PDF, headers=H, timeout=(20, 120))
out(f"PDF {r.status_code} {r.headers.get('content-type')} {len(r.content)} LM={r.headers.get('last-modified')} "
    f"ETag={r.headers.get('etag')}")
with pdfplumber.open(io.BytesIO(r.content)) as pdf:
    out(f"{len(pdf.pages)} pages")
    for i, p in enumerate(pdf.pages[:6]):
        out(f"=== page {i + 1} ({p.width}x{p.height})")
        out(p.extract_text()[:6000])
        for t in p.extract_tables()[:2]:
            out("--- table", len(t), "rows")
            for row in t[:45]:
                out(row)
r = requests.get("https://doe.gov.ph/data-and-prices/energy-statistics/electric-power-industry", headers=H, timeout=(20, 90))
out(sorted(set(re.findall(r'electric-power-industry/[a-z0-9\-]+', r.text))))
