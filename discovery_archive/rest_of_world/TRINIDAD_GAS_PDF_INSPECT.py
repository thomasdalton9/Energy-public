"""
Check parseability of Trinidad & Tobago's Ministry of Energy and Energy
Industries (MEEI) Consolidated Monthly Bulletin - a real, live,
recently-published PDF (confirmed via web search) that includes
"natural gas utilization by sector" and "natural gas production and
utilization" sections. data.gov.tt (the structured CSV alternative) is
down for maintenance, so this PDF is the current live lead.

Just extracts and prints raw text/tables from the relevant pages to see
whether the sector-level gas data is in a machine-parseable table or
free-form prose, before committing to building a real parser.
"""
import io
import sys

import requests

URL = ("https://www.energy.gov.tt/wp-content/uploads/2025/09/"
       "MEEI-Consolidated-Monthly-Bulletins_January-April-2025-14-07-2025.pdf")
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
TIMEOUT = (10, 120)


def main():
    r = requests.get(URL, headers=HEADERS, timeout=TIMEOUT)
    print(f"status={r.status_code} bytes={len(r.content)} content-type={r.headers.get('content-type')}",
          file=sys.stderr)
    if r.status_code != 200:
        print(r.text[:1000], file=sys.stderr)
        return

    import pdfplumber
    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
        print(f"  {len(pdf.pages)} pages", file=sys.stderr)
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if "natural gas" in text.lower() and ("sector" in text.lower() or "utilization" in text.lower()
                                                    or "utilisation" in text.lower()):
                print(f"\n=== page {i + 1}: text ===", file=sys.stderr)
                print(text[:2000], file=sys.stderr)
                tables = page.extract_tables()
                print(f"  {len(tables)} table(s) extracted on this page", file=sys.stderr)
                for t_idx, table in enumerate(tables):
                    print(f"  --- table {t_idx} ({len(table)} rows) ---", file=sys.stderr)
                    for row in table[:15]:
                        print(f"    {row}", file=sys.stderr)


if __name__ == "__main__":
    main()
