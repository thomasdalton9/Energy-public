"""Probe: Energinet Energi Data Service electricity balance (Denmark) vs the ENTSO-E series used in the Europe master."""
import json, requests
B = "https://api.energidataservice.dk/dataset"
for ds in ("ElectricityBalanceNonv", "ElectricityBalance", "ProductionConsumptionSettlement"):
    try:
        r = requests.get(f"{B}/{ds}", params={"start": "2025-01-01T00:00", "end": "2025-01-01T03:00", "limit": 4}, timeout=60)
        print(ds, r.status_code)
        j = r.json()
        print(json.dumps(j.get("records", j)[:2] if isinstance(j.get("records", j), list) else j, indent=1)[:2500])
    except Exception as e:  # noqa: BLE001
        print(ds, "FAILED", e)
# monthly 2025 totals (GWh) from the hourly balance, summed over both price areas
for ds in ("ElectricityBalanceNonv",):
    try:
        r = requests.get(f"{B}/{ds}", params={"start": "2025-01-01T00:00", "end": "2026-01-01T00:00", "limit": 0, "sort": "HourUTC asc"}, timeout=300)
        j = r.json()
        recs = j["records"]
        import pandas as pd
        d = pd.DataFrame(recs)
        d["HourUTC"] = pd.to_datetime(d["HourUTC"])
        num = d.select_dtypes("number").columns
        print(ds, len(d), "rows", list(d.columns))
        print((d[num].sum() / 1000).round(0).to_string())
    except Exception as e:  # noqa: BLE001
        print(ds, "annual FAILED", e)
