"""
GNI's open-data gas supply/demand CSVs end 2026-03-31 (confirmed via
IRELAND_GNI_RESOURCES_INSPECT.py: one file per dataset, republished
2026-08 with a ~4.5 month lag). For a series that runs to yesterday,
check the ENTSOG Transparency Platform: list every point GNI (and the
GB side of Moffat) operates, then pull daily Physical Flow for the
candidate entry points (Moffat import, Corrib/Bellanaboy and Inch
production) for 2026 to see what actually has data and in what unit.
"""
import sys

import pandas as pd
import requests

API = "https://transparency.entsog.eu/api/v1"
HEADERS = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
TIMEOUT = (15, 120)


def get(path, **params):
    r = requests.get(f"{API}/{path}", headers=HEADERS, params=params, timeout=TIMEOUT)
    print(f"GET {r.url} -> {r.status_code}", file=sys.stderr, flush=True)
    r.raise_for_status()
    return r.json()


print("=== operators mentioning Ireland / GNI ===", flush=True)
ops = get("operators", limit=-1)["operators"]
ie_ops = [o for o in ops if o.get("operatorCountryKey") in ("IE", "GB") and
          ("gas networks" in (o.get("operatorLabel") or "").lower() or o.get("operatorCountryKey") == "IE")]
for o in ie_ops:
    print(f"  {o['operatorKey']}  {o['operatorLabel']}  ({o['operatorCountryKey']}, {o.get('operatorTypeLabel')})", flush=True)

print("\n=== operator point directions for Irish operators ===", flush=True)
opd = get("operatorpointdirections", limit=-1)["operatorpointdirections"]
rows = [p for p in opd if p.get("tSOCountry") == "IE" or "moffat" in (p.get("pointLabel") or "").lower()
        or "corrib" in (p.get("pointLabel") or "").lower() or "inch" in (p.get("pointLabel") or "").lower()
        or "bellanaboy" in (p.get("pointLabel") or "").lower()]
df = pd.DataFrame(rows)
cols = [c for c in ["pointKey", "pointLabel", "operatorKey", "operatorLabel", "directionKey", "tSOCountry",
                    "adjacentCountry", "pointType", "isPipeInPipe", "hasData", "connectedOperators",
                    "adjacentOperatorKey", "isInvalid"] if c in df.columns]
print(df[cols].to_string(), flush=True)

print("\n=== daily Physical Flow, 2026-01-01 .. today, per candidate point/direction ===", flush=True)
for _, p in df.iterrows():
    label = (p.get("pointLabel") or "").lower()
    if not any(k in label for k in ("moffat", "corrib", "inch", "bellanaboy", "gormanston", "twynholm", "brighouse")):
        continue
    try:
        data = get("operationalData", periodType="day", indicator="Physical Flow", pointKey=p["pointKey"],
                   operatorKey=p["operatorKey"], directionKey=p["directionKey"], **{"from": "2026-01-01"},
                   to="2026-09-30", limit=-1)["operationalData"]
    except Exception as e:
        print(f"  {p['pointLabel']} {p['operatorKey']} {p['directionKey']}: ERR {type(e).__name__}: {e}", flush=True)
        continue
    if not data:
        print(f"  {p['pointLabel']} {p['operatorKey']} {p['directionKey']}: no rows", flush=True)
        continue
    d = pd.DataFrame(data)
    d["periodFrom"] = pd.to_datetime(d["periodFrom"])
    unit = d["unit"].iloc[0] if "unit" in d else "?"
    print(f"  {p['pointLabel']} {p['operatorKey']} {p['directionKey']}: {len(d)} days, "
          f"{d['periodFrom'].min().date()} to {d['periodFrom'].max().date()}, unit={unit}, "
          f"mean={d['value'].mean():,.0f}, last={d['value'].iloc[-1]:,.0f}", flush=True)
    print(d[["periodFrom", "value", "flowStatus"] if "flowStatus" in d else ["periodFrom", "value"]].tail(3).to_string(),
          flush=True)
