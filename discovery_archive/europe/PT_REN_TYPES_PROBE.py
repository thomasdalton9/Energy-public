"""Probe: every `type` REN's DataHub GasConsumptionSupplyDaily returns (supply as well as consumption items), sampled every 3rd day of 2025; prints
the mean GWh/d per type and month, plus two raw responses."""
import json
from datetime import date, timedelta
import pandas as pd
import requests
REN = "https://servicebus.ren.pt/datahubapi/gas/GasConsumptionSupplyDaily"
rec = []
d = date(2025, 1, 1)
shown = 0
while d <= date(2025, 12, 31):
    try:
        r = requests.get(REN, params={"culture": "en-US", "date": d.isoformat()}, timeout=(15, 45), headers={"User-Agent": "Mozilla/5.0"})
        rows = r.json() if r.ok else []
    except Exception as e:  # noqa: BLE001
        rows = []
    if shown < 2 and rows and d.month in (1, 7):
        print(d, json.dumps(rows)[:3000], flush=True)
        shown += 1
    for x in rows:
        try:
            rec.append((d.month, x.get("type"), x.get("group", ""), float(x["daily_Accumulation"])))
        except Exception:  # noqa: BLE001
            pass
    d += timedelta(days=3)
df = pd.DataFrame(rec, columns=["m", "type", "group", "v"])
print(df.groupby(["type", "m"]).v.mean().unstack().round(1).to_string())
