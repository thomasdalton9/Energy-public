"""
Indonesia raw power data, discovery round 6 (last). Round 5: the Statistik Ketenagalistrikan 2024 book
(80 pages) has no monthly tables - only annual PLN production by plant type and province (Tabel 35) - and
Tableau guest calls still return 401 (no public view names). This round prints the book's generation
pages for the report, checks the HEESI landing page for xlsx attachments, and the ESDM highlight page.
"""
import io
import re
import sys

import requests
import urllib3

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from INDONESIA_DISCOVERY1 import H, T, out, probe  # noqa: E402

urllib3.disable_warnings()
STAT = "https://gatrik.esdm.go.id/gatrik/gatrik-api//storage/migrasi/download_index/files/"


def book():
    import pdfplumber
    r = requests.get(STAT + "91fa8-buku-statistik-ketenagalistrikan-2024.pdf", headers=H, timeout=180, verify=False)
    pdf = pdfplumber.open(io.BytesIO(r.content))
    for n in (9, 14, 15, 56, 57, 58):
        t = pdf.pages[n - 1].extract_text() or ""
        out(f"\n----- p{n} -----\n" + t[:3000])


if __name__ == "__main__":
    try:
        book()
    except Exception as e:
        out(f"!! book {e!r}")
    r = probe("HEESI", "https://www.esdm.go.id/en/publikasi/handbook-of-energy-economic-statistics-of-indonesia-heesi",
              show=0, links=0, kw=False)
    if r is not None:
        for m in sorted(set(re.findall(r"""href=["']([^"']+\.(?:xlsx?|pdf|zip))["']""", r.text, re.I))):
            out("  HEESI file " + m)
    probe("ESDM highlight", "https://www.esdm.go.id/id/highlight", show=0, links=40)
    probe("ESDM download", "https://www.esdm.go.id/id/download", show=0, links=0)
