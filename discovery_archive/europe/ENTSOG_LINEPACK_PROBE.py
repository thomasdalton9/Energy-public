"""Probe: does ENTSOG publish linepack (indicator 'Linepack') per operator, with what coverage and history?"""
import requests, pandas as pd, json
API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
for ind in ["Linepack", "Linepack Flow", "Linepack Change"]:
    for frm, to in [("2026-01-10", "2026-01-12"), ("2025-07-10", "2025-07-12")]:
        try:
            r = requests.get(f"{API}/operationaldatas", params={"indicator": ind, "periodType": "day", "from": frm, "to": to, "limit": -1, "timezone": "CET"}, headers=H, timeout=120)
            print(ind, frm, r.status_code, len(r.content))
            j = r.json().get("operationaldatas", [])
            print(" rows", len(j))
            if j:
                df = pd.DataFrame(j)
                print(df[["operatorKey", "operatorLabel", "pointKey", "pointLabel", "indicator", "periodFrom", "value", "unit"]].head(60).to_string())
                print(df.groupby("operatorKey").size().to_string())
        except Exception as e:
            print(ind, frm, "ERR", e)
r = requests.get(f"{API}/Indicators", headers=H, timeout=60) if False else None
