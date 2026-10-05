"""Probe (gas residual, Great Britain): monthly sums, 2025-01..2026-09, of the National Gas portal's NTS allocation/balance items (input/output totals, shrinkage,
linepack, aggregate entry/exit, LNG/beach/interconnector daily flows, NTS demand actual, LDZ allocations and UIG). Saves results/residual/gb_ids_monthly.csv."""
import os, time
import requests, pandas as pd
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results", "residual")
B = "https://data.nationalgas.com"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json, text/plain, */*", "Referer": "https://data.nationalgas.com/find-gas-data/view", "Content-Type": "application/json"}
IDS = ["PUBOB578", "PUBOB579", "PUBOB580", "PUBOB581", "PUBOB582", "PUBOB583", "PUBOB584", "PUBOB585", "PUBOB3985", "PUBOB694", "PUBOBJ1621", "PUBOBJ1620", "PUBOB261", "PUBOB262",
       "PUBOB4422", "PUBOB637", "PUBOB652", "PUBOBJ523", "PUBOBJ524", "PUBOBJ515", "PUBOBJ519", "PUBOBJ518", "PUBOBJ516", "PUBOBJ517", "PUBOBJ521", "PUBOBJ1286", "PUBOBJ1287",
       "PUBOBJ1288", "PUBOBJ1289", "PUBOBJ1290", "PUBOBJ1305", "PUBOBJ1306", "PUBOBJ1014", "PUBOB2039", "PUBOB3", "PUBOBJ1023", "PUBOBJ1024", "PUBOBJ1025", "PUBOBJ1026", "PUBOBJ1028",
       "PUBOB3876", "PUBOB16241", "PUBOB3877", "PUBOB3881", "PUBOB3891", "PUBOB4423", "PUBOB4424", "PUBOB570", "PUBOBJ2471", "PUBOBJ2484"]
rows = []
months = pd.period_range("2025-01", "2026-09", freq="M")
path = os.path.join(OUT, "gb_ids_monthly.csv")
for i in range(0, len(IDS), 5):
    ids = IDS[i:i + 5]
    for q0 in range(0, len(months), 3):
        ms = months[q0:q0 + 3]
        body = {"latestFlag": "Y", "applicableFor": "Y", "dateFrom": ms[0].start_time.date().isoformat(), "dateTo": ms[-1].end_time.date().isoformat(), "dateType": "GASDAY", "ids": ",".join(ids)}
        for attempt in range(3):
            try:
                r = requests.post(B + "/api/find-gas-data", json=body, headers=H, timeout=120)
                if not r.ok:
                    time.sleep(5); continue
                for it in r.json().get("data", []):
                    v = it.get("value")
                    if v is None: continue
                    try:
                        d = pd.to_datetime(it["applicableFor"], format="%d/%m/%Y")
                        rows.append((it.get("itemName"), d.strftime("%Y-%m"), d.strftime("%Y-%m-%d"), float(v), it.get("unit") or it.get("unitOfMeasure")))
                    except Exception:
                        pass
                break
            except Exception as e:
                print("fail", type(e).__name__, flush=True); time.sleep(5)
    df = pd.DataFrame(rows, columns=["item", "month", "day", "value", "unit"])
    g = df.groupby(["item", "month"]).agg(sum=("value", "sum"), n=("day", "nunique"), unit=("unit", "first")).reset_index()
    g.to_csv(path, index=False)
    print("done", i, len(rows), flush=True)
