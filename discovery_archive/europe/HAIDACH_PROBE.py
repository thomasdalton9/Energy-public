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


lst = requests.get("https://agsi.gie.eu/api/about", params={"show": "listing", "country": "AT"}, headers=H, timeout=120)
print("listing", lst.status_code, len(lst.content), flush=True)
kids = []
for comp in lst.json():
    print("company", comp.get("name"), comp.get("eic"), flush=True)
    for fac in comp.get("facilities", []) or []:
        print("   facility", fac.get("name"), fac.get("eic"), fac.get("operational_start_date"), fac.get("operational_end_date"), flush=True)
        kids.append(("fac", fac.get("name"), fac.get("eic"), comp.get("eic")))
for kind, name, eic, coeic in kids:
    params = {"country": "AT", "company": coeic, "facility": eic, "from": "2022-09-30", "to": date.today().isoformat()}
    rows = agsi(params)
    if not rows:
        print("no rows for", name, eic, flush=True); continue
    df = pd.DataFrame(rows)
    df["gasDayStart"] = pd.to_datetime(df["gasDayStart"])
    s = pd.to_numeric(df.set_index("gasDayStart")["gasInStorage"], errors="coerce").sort_index()
    wgv = pd.to_numeric(df["workingGasVolume"], errors="coerce").max()
    m = s.resample("MS").last()
    ch = m.diff().round(2)
    print(f"STOCKCHANGE {name} {eic} wgv={wgv} first={s.index.min().date()} last={s.index.max().date()} monthly TWh from {m.index[1].date() if len(m)>1 else ''}: " + ",".join(f"{x:.1f}" for x in ch.iloc[1:].values), flush=True)
sys.exit(0)
