"""
Follow-up to TRINIDAD_GAS_EXCEL_DISCOVERY.py, which confirmed MEEI's
bulletin listing page (energy.gov.tt/?p=11949) always links the CURRENT
bulletin directly (solving TRINIDAD_GAS.py's "needs manual URL updates"
limitation) and that an Excel version exists alongside the PDF.
Checking the Excel's real sheet names and the Production-by-Company /
Utilization-by-Sector tables' actual layout before rewriting the
production script to use it instead of PDF table extraction.
"""
import io
import sys

import pandas as pd
import requests

HEADERS = {"User-Agent": "Mozilla/5.0"}
TIMEOUT = (10, 60)
URL = ("https://www.energy.gov.tt/wp-content/uploads/2026/09/"
       "MEEI-Consolidated-Monthly-Bulletins_January-May-2026-16-06-2026.xlsx")

r = requests.get(URL, headers=HEADERS, timeout=TIMEOUT)
print(f"status={r.status_code} bytes={len(r.content)}", file=sys.stderr)

xl = pd.ExcelFile(io.BytesIO(r.content))
print(f"\n{len(xl.sheet_names)} sheets: {xl.sheet_names}", file=sys.stderr)

for name in xl.sheet_names:
    if "3" in name or "gas" in name.lower() or "GAS" in name:
        print(f"\n=== sheet {name!r} (candidate) ===", file=sys.stderr)
        df = xl.parse(name, header=None)
        print(f"  shape: {df.shape}", file=sys.stderr)
        print(df.head(30).to_string(), file=sys.stderr)
