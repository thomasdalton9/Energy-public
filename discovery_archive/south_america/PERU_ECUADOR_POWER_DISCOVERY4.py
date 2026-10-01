"""
Round 4 (Ecuador only): what the Wayback Machine's copies of CENACE's
InformacionOperativa.htm can give for an Ember cross-check.

Round 3 found only 30 distinct snapshot days since Nov-2020 (2020-2023
copies are an old Excel-frames layout), so there is no usable daily
history anywhere - CENACE's daily series starts with the first run of
ECUADOR_CENACE.py (2026-09-26). This round parses the 2024+ snapshots'
daily, month-to-date and year-to-date sections (MWh by source, plus the
thermal oil/gas split) so a same-period comparison with Ember's monthly
data is possible (month-to-date daily averages vs Ember's month).
"""

print("STARTING", flush=True)

import os
import re
import sys
import time

import requests

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                                "south_america"))
import ECUADOR_CENACE as EC  # noqa: E402

S = requests.Session()
S.headers["User-Agent"] = "Mozilla/5.0 (energy-data research; github thomasdalton9/Energy)"


def get(url, **kw):
    for i in range(5):
        try:
            r = S.get(url, timeout=180, **kw)
            if r.status_code in (429, 502, 503, 504):
                time.sleep(20 * (i + 1))
                continue
            return r
        except requests.RequestException as e:
            print("  failed", e, flush=True)
            time.sleep(20 * (i + 1))
    return None


r = get("https://web.archive.org/cdx/search/cdx",
        params={"url": "cenace.gob.ec/info-operativa/InformacionOperativa.htm", "from": "2024",
                "fl": "timestamp,original,statuscode", "filter": "statuscode:200", "collapse": "timestamp:8"})
snaps = [l.split() for l in (r.text.splitlines() if r is not None and r.ok else []) if l.strip()]
print(f"{len(snaps)} snapshot days since 2024", flush=True)
for ts, orig, _ in snaps:
    rr = get(f"https://web.archive.org/web/{ts}id_/{orig}")
    if rr is None or not rr.ok:
        print(ts, "fetch failed", flush=True)
        continue
    h = rr.text
    print(f"\n== {ts} ({len(h):,} chars, {h.count('Plotly.newPlot')} plotly)", flush=True)
    marks = [(m.start(), m.group(1)) for m in re.finditer(r"INFORMACIÓN OPERATIVA (DIARIA|MENSUAL|ANUAL)", h)]
    for i, (pos, kind) in enumerate(marks):
        end = marks[i + 1][0] if i + 1 < len(marks) else len(h)
        text = EC.plain(h[pos:end])
        head = re.search(r"OPERATIVA \w+ (.{0,80}?) PRODUCCI", text)
        vals = {}
        for label, key in EC.TOTALS:
            m = re.search(re.escape(label) + r"\s+([\d\s., \xa0]+?)(?=\s+[A-ZÁÉÍÓÚ(]|$)", text)
            if m:
                try:
                    vals[key] = EC.number(m.group(1))
                except ValueError:
                    pass
        split = {}
        for fig in EC.plots(h, pos, end):
            names = {t[0] for t in fig}
            if "Gas Natural" in names:
                split = {t[0]: (t[3][0] if t[3] else None) for t in fig}
        print(f"  {kind:8s} {head.group(1) if head else '?'!s:60.60s} {vals} split={split}", flush=True)

print("\nDONE", flush=True)
