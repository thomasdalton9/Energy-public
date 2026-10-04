"""Probe: Elering dashboard OpenAPI gas endpoints (paths, parameters) and a sample response of each, for the Estonian/Latvian gas balance
(Karksi and Balticconnector flows, Estonian consumption). Prints only."""
import json
import requests
B = "https://dashboard.elering.ee"
spec = requests.get(B + "/v3/api-docs", timeout=30).json()
paths = {p: v for p, v in spec["paths"].items() if "gas" in p.lower()}
for p, v in paths.items():
    for m, d in v.items():
        print(m.upper(), p, [(x["name"], x.get("schema", {}).get("type")) for x in d.get("parameters", [])], d.get("summary", ""))
schemas = spec.get("components", {}).get("schemas", {})
for k in schemas:
    if k.startswith("Gas") and "Collection" not in k:
        print("SCHEMA", k, json.dumps({a: b.get("type") for a, b in schemas[k].get("properties", {}).items()}))
for p in paths:
    if "{" in p:
        continue
    for q in ({"start": "2025-03-01T00:00:00.000Z", "end": "2025-03-05T00:00:00.000Z"}, {}):
        try:
            r = requests.get(B + p, params=q, timeout=40)
            print("GET", p, q, r.status_code, r.text[:600].replace("\n", " "), flush=True)
            if r.ok:
                break
        except Exception as e:  # noqa: BLE001
            print("ERR", p, e)
