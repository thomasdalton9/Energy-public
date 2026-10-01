"""
One-off: download Gas Networks Ireland's transmission tariff PDFs (annual
tariffs + short-term capacity examples) and print their text, to read the
daily (D-1) capacity multiplier and monthly seasonal factors at Moffat.
Run in GitHub Actions (gasnetworks.ie is blocked from the sandbox).
"""
import io
import re

import requests
from pypdf import PdfReader

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
PAGE = "https://www.gasnetworks.ie/about/regulation/tariffs/transmission-tariffs"
KNOWN = [
    "https://www.gasnetworks.ie/sites/default/files/2025-08/Tariffs-Short-Term-Capacity-Examples-2025-26.pdf",
    "https://www.gasnetworks.ie/sites/default/files/2025-08/Tariffs-Tx-2025-26.pdf",
]


def links():
    found = []
    try:
        html = requests.get(PAGE, headers=H, timeout=60).text
        found = sorted({u if u.startswith("http") else "https://www.gasnetworks.ie" + u
                        for u in re.findall(r'href="([^"]+\.pdf)"', html)
                        if re.search(r"(Short-Term|Tariffs-Tx-20)", u, re.I)})
        print("from page:", *found, sep="\n  ")
    except requests.RequestException as e:
        print("page failed:", e)
    newest = [u for u in found if re.search(r"2026-27|2026", u)]
    return list(dict.fromkeys(newest + KNOWN + found))[:6]


for url in links():
    print("=" * 100, "\n", url, flush=True)
    try:
        r = requests.get(url, headers=H, timeout=120)
        r.raise_for_status()
        for i, page in enumerate(PdfReader(io.BytesIO(r.content)).pages):
            print(f"--- page {i + 1}")
            print(page.extract_text())
    except Exception as e:
        print("failed:", type(e).__name__, e)
