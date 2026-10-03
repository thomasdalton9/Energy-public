"""
Bangladesh discovery, round 11: why bangladesh_gas.xlsx lacks the 29th (and 28 Feb) of nearly every month.
For the listing pages covering late Sep 2026 and late Feb 2026, print each row's raw label text, the label date the
pull derives, the PDF url and the PDF's own 'Date :' line plus asia/BANGLADESH_PETROBANGLA_GAS.report_date().
"""
import io
import os
import re
import sys

import pdfplumber

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "asia"))
import BANGLADESH_PETROBANGLA_GAS as pb  # noqa: E402


def raw_rows(page):
    r = pb.get(pb.LIST, params={"filters": pb.FILTER, "page": page})
    html = re.sub(r"<!--.*?-->", " ", r.text, flags=re.S)
    for chunk in re.split(r"<tr[\s>]", html)[1:]:
        m = re.search(r'https://objectstorage[^"\']+\.pdf', chunk)
        if m:
            yield re.sub(r"<[^>]+>|\s+", " ", chunk).translate(pb.BN).strip()[:160], m.group(0)


def main():
    for page in (1, 2, 3, 23, 24, 25):
        print(f"==== page {page}", flush=True)
        labels = dict((u, d) for d, u in pb.listing_page(page))
        for text, url in raw_rows(page):
            try:
                c = pb.get(url).content
                with pdfplumber.open(io.BytesIO(c)) as pdf:
                    t = "\n".join((p.extract_text() or "") for p in pdf.pages[:2])
                line = re.search(r"Date[^\n]{0,80}", t)
                line = line.group(0) if line else t[:200].replace("\n", " | ")
                rd = pb.report_date(t)
                tot = pb.parse(t).get("Total_supply")
            except Exception as e:  # noqa: BLE001
                line, rd, tot = f"ERR {e}", None, None
            print(f"  label={labels.get(url)} | text={text!r}\n    url=...{url[-60:]}\n    pdf: {line!r} -> {rd} total {tot}",
                  flush=True)


if __name__ == "__main__":
    main()
