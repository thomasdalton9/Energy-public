"""Probe: Energinet Energi Data Service ProductionConsumptionSettlement annual totals (GWh) 2023-2025, DK1+DK2, to compare with ENTSO-E."""
import requests, pandas as pd
B = "https://api.energidataservice.dk/dataset/ProductionConsumptionSettlement"
for y in (2023, 2024, 2025):
    r = requests.get(B, params={"start": f"{y}-01-01T00:00", "end": f"{y + 1}-01-01T00:00", "limit": 0}, timeout=300)
    j = r.json()
    recs = j.get("records")
    if recs is None:
        print(y, r.status_code, str(j)[:500]); continue
    d = pd.DataFrame(recs)
    num = d.select_dtypes("number").columns
    print(y, len(d), "rows")
    print((d[num].sum() / 1000).round(0).to_string(), flush=True)
