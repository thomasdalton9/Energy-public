"""
Follow-up to TRINIDAD_GAS.py (PDF-based, one hardcoded bulletin URL,
documented as needing manual updates). A web search found:
  1. A real listing/index page (energy.gov.tt/?p=11949) that might give
     a stable way to find the CURRENT bulletin without guessing dates.
  2. The bulletins are published in BOTH PDF and Excel formats - Excel
     would be far more reliable to parse than PDF table extraction
     (which needed a real workaround this session for a table-merge
     quirk).
  3. The current (as of this search) bulletin:
     https://www.energy.gov.tt/wp-content/uploads/2026/07/MEEI-Consolidated-Monthly-Bulletins_January-March-2026-04-05-2026.pdf

Checking the listing page for an Excel download link alongside the PDF,
and confirming the direct PDF URL still resolves.
"""
import re
import sys

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 45)


def try_get(label, url, **kwargs):
    print(f"\n{'=' * 70}\n{label}: {url}\n{'=' * 70}", file=sys.stderr)
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kwargs)
        print(f"  status={r.status_code} bytes={len(r.content)} content-type={r.headers.get('content-type')}",
              file=sys.stderr)
        return r
    except requests.RequestException as e:
        print(f"  ERROR: {type(e).__name__}: {e}", file=sys.stderr)
        return None


r = try_get("MEEI bulletin listing page", "https://www.energy.gov.tt/?p=11949")
if r is not None and r.status_code == 200:
    text = r.text
    pdf_links = sorted(set(re.findall(r'href="([^"]+\.pdf)"', text, re.I)))
    xlsx_links = sorted(set(re.findall(r'href="([^"]+\.xlsx?)"', text, re.I)))
    print(f"  PDF links on page: {pdf_links[:10]}", file=sys.stderr)
    print(f"  Excel links on page: {xlsx_links[:10]}", file=sys.stderr)

try_get("Current bulletin PDF (confirming it resolves)",
        "https://www.energy.gov.tt/wp-content/uploads/2026/07/"
        "MEEI-Consolidated-Monthly-Bulletins_January-March-2026-04-05-2026.pdf")

# Guessing the Excel version shares the same filename stem, common on
# WordPress media uploads when both formats are published together.
try_get("Guessed Excel version of the current bulletin",
        "https://www.energy.gov.tt/wp-content/uploads/2026/07/"
        "MEEI-Consolidated-Monthly-Bulletins_January-March-2026-04-05-2026.xlsx")
