"""
Probe 18: run the Enagas bulletin parser (europe/GAS_TSO_SOUTHEAST_DAILY.py parse_enagas_bulletin) on every listed bulletin 2021-2026;
for those that fail print the text of the pages around the demand table. Prints only (short).
"""
import os
import re
import sys
from datetime import datetime, timezone

import requests

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "europe"))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
import GAS_TSO_SOUTHEAST_DAILY as g  # noqa: E402

H = {"User-Agent": g.UA}
files = {}
today = datetime.now(timezone.utc).date()
for y in range(2020, today.year + 1):
    for m in range(1, 13):
        try:
            r = requests.get(g.ENAGAS_PAGE, params={"category": "", "month": m, "year": y}, headers=H, timeout=60)
            for x in re.findall(r'href="(/content/dam[^"]+\.pdf)"', r.text):
                files.setdefault(x, (y, m))
        except Exception:  # noqa: BLE001
            pass
print(len(files), "pdfs")
import io
import pdfplumber
bad = 0
for f, ym in sorted(files.items(), key=lambda kv: kv[1]):
    name = f.split("/")[-1]
    try:
        content = requests.get("https://www.enagas.es" + f, headers=H, timeout=120).content
        res = g.parse_enagas_bulletin(content)
    except Exception as e:  # noqa: BLE001
        print("ERR", name, type(e).__name__)
        continue
    if res:
        print("OK ", ym, name[:50], res[0].strftime("%Y-%m"), res[1:])
    else:
        bad += 1
        print("BAD", ym, name)
        if bad <= 5:
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                for i, pg in enumerate(pdf.pages[:8]):
                    t = pg.extract_text() or ""
                    if re.search(r"(?i)demand", t) and "GWh" in t:
                        print(f"   --- page {i + 1}: {re.sub(chr(10), ' | ', t[:500])}")
                        break
sys.exit(0)
