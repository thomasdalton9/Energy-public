"""Probe 2 (manual, Actions): FERC 'U.S. LNG Export Terminals - Existing, Approved not Yet Built, and Proposed' page ->
linked PDF/xlsx -> text lines for the Texas projects (status, in-service / expected dates). Also the DOE/FERC
construction-status pages if linked."""
import io
import re

import requests

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120 Safari/537.36"}
KW = re.compile(r"rio grande|port arthur|golden pass|corpus christi|texas lng|freeport|brownsville|sabine", re.I)
u = "https://www.ferc.gov/media/north-american-lng-export-terminals-existing-approved-not-yet-built-and-proposed-8"
r = requests.get(u, headers=H, timeout=60)
print(u, r.status_code)
links = sorted(set(re.findall(r'href="([^"]+\.(?:pdf|xlsx|xls|csv)[^"]*)"', r.text, re.I)))
print(links[:30])
for l in links[:6]:
    if l.startswith("/"):
        l = "https://www.ferc.gov" + l
    try:
        rr = requests.get(l, headers=H, timeout=90)
        print("==", l, rr.status_code, len(rr.content), rr.headers.get("content-type"))
        if l.lower().endswith(".pdf") or "pdf" in rr.headers.get("content-type", ""):
            from pypdf import PdfReader
            rd = PdfReader(io.BytesIO(rr.content))
            for pi, pg in enumerate(rd.pages):
                for ln in (pg.extract_text() or "").splitlines():
                    if KW.search(ln):
                        print(f"  p{pi + 1}: {ln[:300]}")
        else:
            import pandas as pd
            x = pd.read_excel(io.BytesIO(rr.content), sheet_name=None, header=None)
            for n, d in x.items():
                for _, row in d.iterrows():
                    t = " | ".join(str(v) for v in row if str(v) != "nan")
                    if KW.search(t):
                        print(f"  {n}: {t[:400]}")
    except Exception as e:  # noqa: BLE001
        print("fail", l, type(e).__name__, str(e)[:150])
