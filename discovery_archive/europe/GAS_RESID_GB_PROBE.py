"""Probe (EU27+UK gas residual, Great Britain): National Gas Data Portal catalogue (all items) and monthly sums of every demand/supply/linepack/shrinkage/
fuel/actual/LNG item, 2025-01..2026-09. Writes results/residual/gb_catalogue.txt and gb_items_monthly.csv."""
import os, re, sys, time
import requests, pandas as pd
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results", "residual")
os.makedirs(OUT, exist_ok=True)
B = "https://data.nationalgas.com"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36", "Accept": "application/json, text/plain, */*",
     "Referer": "https://data.nationalgas.com/find-gas-data/view", "Content-Type": "application/json"}
txt = requests.get(B + "/api/find-gas-data-folders", headers=H, timeout=90).text
items = {}
for m in re.finditer(r'"name":\s*"([^"]*)",\s*"description":\s*"((?:PUBOB?J?\d+)[^"]*)"', txt):
    items[re.match(r"(PUBOB?J?\d+)", m.group(2)).group(1)] = m.group(1) + " || " + m.group(2)[:140]
print("items", len(items), flush=True)
open(os.path.join(OUT, "gb_catalogue.txt"), "w").write("\n".join(f"{k}\t{v}" for k, v in sorted(items.items())))
pat = re.compile(r"shrink|own use|compress|fuel|total|physical|actual|LNG|embedded|unaccounted|UIG|linepack|line pack|imbalance|offtake|offtaken|demand|supply|storage|interconnector|moffat|beach|terminal|biomethane|entry|NTS|DN|LDZ", re.I)
sel = [k for k, v in items.items() if pat.search(v)]
print("selected", len(sel), flush=True)
rows = []
months = pd.period_range("2025-01", "2026-09", freq="M")
for i in range(0, len(sel), 20):
    ids = sel[i:i + 20]
    for q0 in range(0, len(months), 6):
        ms = months[q0:q0 + 6]
        body = {"latestFlag": "Y", "applicableFor": "Y", "dateFrom": ms[0].start_time.date().isoformat(), "dateTo": ms[-1].end_time.date().isoformat(),
                "dateType": "GASDAY", "ids": ",".join(ids)}
        for attempt in range(3):
            try:
                r = requests.post(B + "/api/find-gas-data", json=body, headers=H, timeout=240)
                if not r.ok:
                    time.sleep(5); continue
                for it in r.json().get("data", []):
                    v = it.get("value")
                    if v is None: continue
                    try:
                        d = pd.to_datetime(it["applicableFor"], format="%d/%m/%Y")
                        rows.append((it.get("itemName"), d.strftime("%Y-%m"), float(v), it.get("unit") or it.get("unitOfMeasure")))
                    except Exception:
                        pass
                break
            except Exception as e:
                print("fail", type(e).__name__, flush=True); time.sleep(5)
    print("done", i, len(rows), flush=True)
df = pd.DataFrame(rows, columns=["item", "month", "value", "unit"])
g = df.groupby(["item", "month"]).agg(sum=("value", "sum"), n=("value", "size"), unit=("unit", "first")).reset_index()
g.to_csv(os.path.join(OUT, "gb_items_monthly.csv"), index=False)
print(g.item.nunique(), "items with data")
