"""
Probe (Oct 2026) for middle_east/ISRAEL_GAS_FIELDS.py: Karish year-to-date gas for FY 2022 (first gas 26 Oct 2022),
9M 2023 and FY 2023 was not parsed. Prints, for every Energean announcement dated Oct 2022 - Apr 2024, the full
context (no sentence splitting) around every 'bcm' and 'totaled approximately'.

Workflow (manual): discovery_archive/workflows/energean_karish_gaps_probe.yml.
"""
import importlib.util
import os
import re

import requests

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
spec = importlib.util.spec_from_file_location("g", os.path.join(ROOT, "middle_east", "ISRAEL_GAS_FIELDS.py"))
g = importlib.util.module_from_spec(spec)
spec.loader.exec_module(g)

MONTHS = r"(?:January|February|March|April|May|June|July|August|September|October|November|December)"
for url in g.energean_announcements(set()):
    t = g.rns_text(requests.get(url, headers=g.UA, timeout=90).text)
    d = re.search(rf"\b(\d{{1,2}}) ({MONTHS}) (20\d\d)\b", t)
    if not d or not (("2022" == d.group(3) and d.group(2) in ("October", "November", "December")) or d.group(3) == "2023"
                     or (d.group(3) == "2024" and d.group(2) in ("January", "February", "March", "April"))):
        continue
    hits = [t[max(0, m.start() - 260): m.end() + 140] for m in re.finditer(r"\bbcm\b|totaled approximately|bcf\b", t)]
    hits = [h for h in hits if re.search(r"sold|sales|production|produced|Karish|Israel", h, re.I)
            and not re.search(r"bcm/y|bcm/yr|bcm/year|reserves|resources|contract", h, re.I)]
    print(f"\n=== {url.split('energean--enog/')[1]}  {d.group(0)}  parsed={g.parse_karish(t)}")
    for h in hits[:10]:
        print("   |", h.strip())
