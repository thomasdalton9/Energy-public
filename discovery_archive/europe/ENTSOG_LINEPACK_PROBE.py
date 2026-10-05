"""Probe: does ENTSOG publish linepack per operator? Tries indicator spellings, period types and the point list."""
import requests, pandas as pd
API = "https://transparency.entsog.eu/api/v1"
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
def get(path, **p):
    r = requests.get(f"{API}/{path}", params=p, headers=H, timeout=120)
    print(path, p, r.status_code, len(r.content), r.text[:120].replace("\n", " "))
    try:
        return r.json()
    except Exception:
        return {}
for ind in ["Linepack", "linepack", "LinePack", "Line pack", "Physical Flow"]:
    for pt in ["day", "hour"]:
        j = get("operationaldatas", indicator=ind, periodType=pt, **{"from": "2026-01-10", "to": "2026-01-11"}, limit=3, timezone="CET")
        print("  rows", len(j.get("operationaldatas", [])))
j = get("operationaldatas", indicator="Linepack", periodType="day", operatorKey="FR-TSO-0001", **{"from": "2025-01-10", "to": "2025-01-12"}, limit=10)
for k in ("operators", "operatorpointdirections"):
    pass
j = get("operatorpointdirections", limit=-1)
rows = j.get("operatorpointdirections", [])
df = pd.DataFrame(rows)
print(len(df), list(df.columns)[:20])
m = df[df.apply(lambda r: "inepack" in str(r.to_dict()) or "LNP" in str(r.get("pointKey", "")), axis=1)]
print(len(m)); print(m.head(40).to_string())
