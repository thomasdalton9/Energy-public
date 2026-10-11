"""
Probe (Oct 2026) for middle_east/ISRAEL_GAS_FIELDS.py: how do Energean's announcements (investegate.co.uk, ENOG) word
Israel / Karish gas volumes? Prints, per announcement since 2022, every table row or sentence with 'bcm' or an
Israel production row, so the Karish year-to-date parser can cover Q1 updates and 2022-23 wording.

RUN 1 (2026-10-11): Q1 trading updates put the gas in the table ('Israel | 59.0 (inc. 0.7 bcm of sales gas)'),
Q1 2026 only in a sentence ('0.96 bcm of gas in Q1 2026 and 1.19 bcm in Q1 2025'); Energean Israel half/full-year
accounts state '... six months ended 30 June 2025 totaled approximately 2.3 bcm'. Parser updated accordingly.

Workflow (manual): discovery_archive/workflows/energean_karish_text_probe.yml.
"""
import importlib.util
import os
import re

import requests

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
spec = importlib.util.spec_from_file_location("g", os.path.join(ROOT, "middle_east", "ISRAEL_GAS_FIELDS.py"))
g = importlib.util.module_from_spec(spec)
spec.loader.exec_module(g)

for url in g.energean_announcements(set()):
    t = g.rns_text(requests.get(url, headers=g.UA, timeout=90).text)
    date = re.search(r"\b(\d{1,2} (?:January|February|March|April|May|June|July|August|September|October|November|December) 20\d\d)\b", t)
    hits = [m.group(0) for m in re.finditer(r"[^.|]{0,200}\bbcm\b[^.|]{0,120}", t)]
    rows = [t[max(0, m.start() - 350): m.end() + 250] for m in re.finditer(r"Israel[^|]{0,60}\| \d+", t)]
    if not hits and not rows:
        continue
    print(f"\n=== {url.split('energean--enog/')[1]}  {date.group(1) if date else ''}  parsed={g.parse_karish(t)}")
    for h in hits[:8]:
        print("   BCM |", h.strip()[:330])
    for r in rows[:3]:
        print("   ROW |", r[:620])
