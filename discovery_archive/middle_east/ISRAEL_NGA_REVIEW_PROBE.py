"""
Probe (Oct 2026): which Israel Natural Gas Authority annual 'Natural Gas Sector Review' PDFs exist on gov.il BlobFolder,
and what their text layer looks like. Downloads every candidate URL with curl_cffi (impersonate chrome), saves the PDFs
to probe_out/ (uploaded as a workflow artifact) and prints a page/char summary. Used to build middle_east/ISRAEL_NGA_ANNUAL.py.
"""
import os
import sys

from curl_cffi import requests as creq

OUT = "probe_out"
os.makedirs(OUT, exist_ok=True)
pats = ["ng_{y}/he/ng_{y}.pdf", "ng-{y}/he/ng-{y}.pdf", "ng_{y}/he/ng-{y}.pdf", "ng-{y}/he/ng_{y}.pdf",
        "ng{y}/he/ng{y}.pdf", "natural_gas_{y}/he/natural_gas_{y}.pdf", "ng_{y}/en/ng_{y}.pdf", "ng-{y}/en/ng-{y}.pdf",
        "ng_report_{y}/he/ng_report_{y}.pdf", "ng_{y}_summary/he/ng_{y}_summary.pdf"]
for y in range(2008, 2027):
    for p in pats:
        u = "https://www.gov.il/BlobFolder/reports/" + p.format(y=y)
        try:
            r = creq.get(u, impersonate="chrome", timeout=90)
        except Exception as e:  # noqa: BLE001
            print(y, u, "ERR", e)
            continue
        ok = r.status_code == 200 and r.content.startswith(b"%PDF")
        print(y, r.status_code, len(r.content), "PDF" if ok else "", u, r.headers.get("Last-Modified"), flush=True)
        if ok:
            fn = os.path.join(OUT, p.format(y=y).replace("/", "__"))
            open(fn, "wb").write(r.content)
            try:
                import pdfplumber
                with pdfplumber.open(fn) as pdf:
                    n = len(pdf.pages)
                    chars = sum(len(pg.extract_text() or "") for pg in pdf.pages)
                print("    pages", n, "chars", chars)
            except Exception as e:  # noqa: BLE001
                print("    pdfplumber err", e)
sys.exit(0)
