"""
Probe: National Gas Data Portal catalogue items about NTS own use / shrinkage / compressor fuel / total demand / physical flow / actuals
(to find the gas Great Britain's NTS offtake series leaves out), then the 2025 and 2024 sums (kWh -> GWh) of every match. Prints only.
"""
import re
import sys

import requests

BASE = "https://data.nationalgas.com"
H = {"User-Agent": "Mozilla/5.0", "Content-Type": "application/json", "Accept": "application/json"}
txt = requests.get(BASE + "/api/find-gas-data-folders", headers=H, timeout=90).text
items = {}
for m in re.finditer(r'"name":\s*"([^"]*)",\s*"description":\s*"((?:PUBOB?J?\d+)[^"]*)"', txt):
    items[re.match(r"(PUBOB?J?\d+)", m.group(2)).group(1)] = m.group(1)
print("catalogue items", len(items), flush=True)
pat = re.compile(r"shrink|own use|compress|fuel|usage|total demand|physical flow|actual|LNG|embedded|unaccounted|UIG|calorific|demand", re.I)
sel = {k: v for k, v in items.items() if pat.search(v)}
print("matching", len(sel), flush=True)
for k, v in sorted(sel.items(), key=lambda kv: kv[1]):
    print("  ", k, v)
sys.stdout.flush()
for yr in (2024, 2025):
    print("== sums", yr, flush=True)
    ids = list(sel)
    for i in range(0, len(ids), 25):
        body = {"latestFlag": "Y", "applicableFor": "Y", "dateFrom": f"{yr}-01-01", "dateTo": f"{yr}-12-31", "dateType": "GASDAY", "ids": ",".join(ids[i:i + 25])}
        try:
            r = requests.post(BASE + "/api/find-gas-data", json=body, headers=H, timeout=240)
            tot = {}
            for it in r.json().get("data", []):
                v = it.get("value")
                if v is None:
                    continue
                n = it.get("itemName")
                a = tot.setdefault(n, [0.0, 0, it.get("unit")])
                a[0] += float(v)
                a[1] += 1
            for n, (s, c, u) in sorted(tot.items()):
                print(f"   {n} | n={c} sum={s:.0f} unit={u}", flush=True)
        except Exception as e:  # noqa: BLE001
            print("   chunk failed", type(e).__name__, str(e)[:120], flush=True)
