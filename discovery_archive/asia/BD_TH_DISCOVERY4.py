"""
Bangladesh discovery, round 4: full text of Petrobangla daily gas reports in the older (2021 - Feb 2022) layout,
whose sub-total / R-LNG lines BANGLADESH_PETROBANGLA_GAS.py misread. Prints the text of the reports dated
19 Jan 2021, 9 Jan 2021 and 15 Jun 2021, plus one from Mar 2022 (new layout) for comparison.
"""
import io
import os
import sys
from datetime import date

import pdfplumber

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "asia"))
import BANGLADESH_PETROBANGLA_GAS as P  # noqa: E402

WANT = {date(2021, 1, 19), date(2021, 1, 9), date(2021, 6, 15), date(2022, 3, 15)}


def main():
    found = {}
    for page in range(150, 240):
        rows = P.listing_page(page)
        for d, u in rows:
            if d in WANT and d not in found:
                found[d] = u
        ds = [d for d, _ in rows if d]
        print(f"page {page}: {min(ds) if ds else None}..{max(ds) if ds else None}; found {sorted(found)}", flush=True)
        if len(found) == len(WANT) or (ds and max(ds) < date(2021, 1, 1)):
            break
    for d, u in sorted(found.items()):
        r = P.get(u)
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            text = "\n".join((p.extract_text() or "") for p in pdf.pages)
        print(f"\n######## {d} {u}\n{text}\nPARSED: {P.parse(text)}", flush=True)


if __name__ == "__main__":
    main()
