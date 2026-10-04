"""
Hydro reservoir discovery, round 9: why only 3 of the Internet Archive's IRSA captures in Dec 2024 - Sep 2026 gave a row
(asia/PAKISTAN_IRSA_RESERVOIRS.py wayback()): status / type / first bytes / parse result per capture.
"""
import io
import os
import re
import sys
import time

import pandas as pd
import pdfplumber
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "asia"))
import PAKISTAN_IRSA_RESERVOIRS as P  # noqa: E402

r = requests.get(P.CDX, params={"url": "pakirsa.gov.pk/Doc/Data*", "output": "json", "collapse": "original",
                                "fl": "timestamp,original,statuscode,mimetype", "limit": 20000}, timeout=(20, 180))
caps = r.json()[1:]
gap = []
for c in caps:
    m = re.search(r"Data(\d\d)-(\d\d)-(\d{4})", c[1])
    if m and pd.Timestamp(int(m.group(3)), int(m.group(2)), int(m.group(1))) > pd.Timestamp("2024-12-19"):
        gap.append(c)
print(len(caps), "captures;", len(gap), "in the gap; statuses", pd.Series([c[2] for c in gap]).value_counts().to_dict(),
      "mimetypes", pd.Series([c[3] for c in gap]).value_counts().to_dict(), flush=True)
for ts, orig, st, mt in gap[:25]:
    t = time.time()
    try:
        p = requests.get(f"http://web.archive.org/web/{ts}id_/{orig}", headers=P.H, timeout=(20, 60))
        head = p.content[:8]
        info = f"{p.status_code} {p.headers.get('content-type')} {len(p.content)}b {head!r} {time.time() - t:.1f}s"
        if p.content[:4] == b"%PDF":
            try:
                row, dead = P.parse_pdf(p.content)
                info += f" -> T {row.get('Tarbela_level_ft')} M {row.get('Mangla_level_ft')}"
                if row.get("Tarbela_level_ft") is None and row.get("Mangla_level_ft") is None:
                    with pdfplumber.open(io.BytesIO(p.content)) as pdf:
                        info += "\n   TEXT: " + (pdf.pages[0].extract_text() or "")[:700].replace("\n", " | ")
            except Exception as e:  # noqa: BLE001
                info += f" parse error {e}"
    except requests.RequestException as e:
        info = f"ERR {e} {time.time() - t:.1f}s"
    print(ts, orig[-20:], st, mt, info, flush=True)
