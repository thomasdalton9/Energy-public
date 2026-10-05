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


lst = requests.get("https://agsi.gie.eu/api/about", params={"show": "listing"}, headers=H, timeout=120)
print("listing", lst.status_code, len(lst.content), lst.text[:300], flush=True)
for u, pr in (("https://agsi.gie.eu/api/about", {}), ("https://agsi.gie.eu/api/about", {"show": "listing", "country": "AT"}), ("https://agsi.gie.eu/api/about", {"show": "companies"}), ("https://agsi.gie.eu/api/about", {"show": "facilities"})):
    x = requests.get(u, params=pr, headers=H, timeout=120)
    print("ABOUT", pr, x.status_code, len(x.content), x.text[:1500].replace("\n", " "), flush=True)
kids = []
try:
    j = lst.json()
    at = j.get("Europe", j).get("AT", {}) if isinstance(j, dict) else {}
    if not at:
        for k, v in j.items():
            if isinstance(v, dict) and "AT" in v:
                at = v["AT"]
    print("AT listing keys:", list(at.keys()) if isinstance(at, dict) else type(at), flush=True)
    for comp in (at.values() if isinstance(at, dict) else at):
        if not isinstance(comp, dict):
            continue
        print("company", comp.get("name"), comp.get("eic"), flush=True)
        for fac in comp.get("facilities", []) or []:
            print("   facility", fac.get("name"), fac.get("eic"), flush=True)
            kids.append(("fac", fac.get("name"), fac.get("eic"), comp.get("eic")))
        kids.append(("co", comp.get("name"), comp.get("eic"), None))
except Exception as e:
    print("listing parse failed", type(e).__name__, e, lst.text[:300], flush=True)
for kind, name, eic, coeic in kids:
    params = {"country": "AT", "from": "2022-09-30", "to": date.today().isoformat()}
    params["company" if kind == "co" else "facility"] = eic
    if kind == "fac":
        params["company"] = coeic
    rows = agsi(params)
    if not rows:
        print("no rows for", kind, name, eic, flush=True); continue
    df = pd.DataFrame(rows)
    df["gasDayStart"] = pd.to_datetime(df["gasDayStart"])
    s = pd.to_numeric(df.set_index("gasDayStart")["gasInStorage"], errors="coerce").sort_index()
    m = s.resample("MS").last()
    ch = m.diff().round(2)
    print(f"STOCKCHANGE {kind} {name} {eic} monthly TWh from {m.index[1].date()}: " + ",".join(f"{x:.1f}" for x in ch.iloc[1:].values), flush=True)
    print(f"STOCK {name}: " + ",".join(f"{x:.1f}" for x in m.values), flush=True)
sys.exit(0)
