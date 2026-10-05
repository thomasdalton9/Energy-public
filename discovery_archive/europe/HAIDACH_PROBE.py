"""
Probe: Haidach storage (Austria, fed from / delivering to the German grid). Eurostat shows German imports from Austria of 6-9 TWh a month in winter
and exports to Austria of 17-19 TWh a month in summer since Dec 2024 that ENTSOG's AT/DE border rows lack. Hypothesis: the Haidach injections and
withdrawals. Prints (1) AGSI+ Austria children (companies, facilities), (2) monthly net stock change (TWh) per child since Oct 2022, from AGSI+;
(3) ENTSOG storage / interconnection points mentioning Haidach (physical flow monthly TWh, 2025).
"""
import os
import sys
import time
from datetime import date

import pandas as pd
import requests

KEY = os.environ.get("GIE_API_KEY", "")
H = {"x-key": KEY, "User-Agent": "Mozilla/5.0"}


def agsi(params):
    out, page = [], 1
    while True:
        r = requests.get("https://agsi.gie.eu/api", params=dict(params, page=page, size=300), headers=H, timeout=120)
        if r.status_code != 200:
            print("AGSI", params, r.status_code, r.text[:120]); return out
        j = r.json(); out += j.get("data", [])
        if page >= int(j.get("last_page", 1) or 1):
            return out
        page += 1; time.sleep(0.3)


r = agsi({"country": "AT", "from": "2025-02-01", "to": "2025-02-01"})
for d in r:
    for c in d.get("children", []) or []:
        print("child", c.get("name"), c.get("eic"), c.get("gasInStorage"), c.get("workingGasVolume"), c.get("url"), flush=True)
        for f in c.get("children", []) or []:
            print("   facility", f.get("name"), f.get("eic"), f.get("gasInStorage"), f.get("workingGasVolume"), flush=True)
if r:
    print({k: v for k, v in r[0].items() if k != "children"})
# company-level history for every child
kids = [(c.get("name"), c.get("eic"), c.get("url")) for d in r for c in d.get("children", []) or []]
for name, eic, url in kids:
    rows = agsi({"country": "AT", "company": eic, "from": "2022-09-30", "to": date.today().isoformat()})
    if not rows:
        rows = agsi({"company": eic, "from": "2022-09-30", "to": date.today().isoformat()})
    if not rows:
        print("no rows for", name, eic); continue
    df = pd.DataFrame(rows)
    df["gasDayStart"] = pd.to_datetime(df["gasDayStart"])
    s = pd.to_numeric(df.set_index("gasDayStart")["gasInStorage"], errors="coerce").sort_index()
    m = s.resample("MS").last()
    ch = (m.diff()).round(2)   # TWh (stock change over the month; first month unknown)
    print(f"STOCKCHANGE {name} {eic} monthly TWh from {m.index[1].date()}: " + ",".join(f"{x:.1f}" for x in ch.iloc[1:].values), flush=True)
    print(f"STOCK {name}: " + ",".join(f"{x:.1f}" for x in m.values), flush=True)
sys.exit(0)
