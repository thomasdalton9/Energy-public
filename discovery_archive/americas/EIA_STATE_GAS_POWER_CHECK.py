"""Sanity check for MISO gas burn: EIA natural gas consumption for electric power (monthly, by state, EIA API v2
natural-gas/cons/sum, process VEU) summed over the MISO states. NOT the same footprint (MISO covers parts of TX,
KY, MT and excludes non-MISO utilities inside the states); order of magnitude only. Prints Bcf/d per month."""
import os, requests, pandas as pd
KEY = os.environ["EIA_API_KEY"]
STATES = ["ND", "SD", "MN", "IA", "WI", "MI", "IL", "IN", "MO", "AR", "LA", "MS"]
r = requests.get("https://api.eia.gov/v2/natural-gas/cons/sum/data/", params={
    "api_key": KEY, "frequency": "monthly", "data[0]": "value", "facets[process][]": "VEU",
    "start": "2024-01", "length": 5000, "sort[0][column]": "period", "sort[0][direction]": "asc",
    **{f"facets[duoarea][{i}]": f"S{s}" for i, s in enumerate(STATES)}}, timeout=60)
print(r.status_code)
j = r.json()["response"]; d = pd.DataFrame(j["data"]); print(j.get("total"), d.columns.tolist())
d["value"] = pd.to_numeric(d["value"], errors="coerce")
d = d[d["duoarea"].isin([f"S{s}" for s in STATES])]
p = d.pivot_table(index="period", columns="duoarea", values="value", aggfunc="sum")
tot = p.sum(axis=1)   # MMcf per month
tot.index = pd.to_datetime(tot.index)
out = pd.DataFrame({"MMcf": tot, "Bcf_per_day": tot / tot.index.days_in_month / 1e3})
print(out.to_string())
print(p.tail(2).T.to_string())
