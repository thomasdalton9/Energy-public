"""
Probe: (1) Gaz-System MIR KspRealization rows for all zones/points (entry points Kondratki, Wysokoje ...), (2) PSE open data API filters/fields/history.
"""
import json
import sys
import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0"
GS = "https://swi.gaz-system.pl/mir/api/v1/en/KspRealization/data"


def mir(d0, d1):
    body = {"start": 0, "length": 2000, "globalFilter": {"value": ""}, "lang": "en",
            "sort": [{"column": "gasDay", "order": "asc"}, {"column": "zoneCode", "order": "asc"}],
            "customFilters": {"filters": [{"filterName": "gasDay$from", "values": d0}, {"filterName": "gasDay$to", "values": d1}]}}
    r = requests.post(GS, json=body, headers={"User-Agent": UA}, timeout=(15, 120))
    print("MIR", r.status_code, len(r.text), flush=True)
    return r.json()


try:
    j = mir("2022-03-01", "2022-03-03")
    print("keys", list(j.keys()), "total", j.get("total"))
    rows = j.get("data", [])
    print("row0", json.dumps(rows[0], ensure_ascii=False) if rows else None)
    seen = {}
    for x in rows:
        seen.setdefault((x.get("zoneCode"), x.get("zoneName") or x.get("zone")), []).append(x.get("realization"))
    for k, v in sorted(seen.items(), key=lambda t: str(t[0])):
        print(k, v[:3])
except Exception as e:  # noqa: BLE001
    print("MIR ERR", repr(e), flush=True)

for ep in ["https://swi.gaz-system.pl/mir/api/v1/en/KspRealization/filters", "https://swi.gaz-system.pl/mir/api/v1/en/PhysicalFlows/data",
           "https://swi.gaz-system.pl/mir/api/v1/en/Nominations/data", "https://swi.gaz-system.pl/mir/api/v1/en/Allocations/data",
           "https://swi.gaz-system.pl/mir/", "https://swi.gaz-system.pl/mir/api/v1/en/Zones/data"]:
    try:
        r = requests.post(ep, json={"start": 0, "length": 5, "globalFilter": {"value": ""}, "lang": "en"}, headers={"User-Agent": UA}, timeout=30)
        print("EP", r.status_code, ep, r.text[:300].replace("\n", " "), flush=True)
    except Exception as e:  # noqa: BLE001
        print("EP ERR", ep, repr(e)[:100], flush=True)

PSE = "https://api.raporty.pse.pl/api/"
for ep, flt in [("his-wlk-cal", "doba eq '2023-06-01'"), ("his-wlk-cal", "business_date eq '2023-06-01'"), ("his-wlk-cal", "udtczas ge '2023-06-01' and udtczas lt '2023-06-02'"),
                ("kse-load", "doba eq '2023-06-01'"), ("zap-kse", "doba eq '2023-06-01'"), ("krajowe-zapotrzebowanie", "doba eq '2023-06-01'"),
                ("his-gen-pal", "doba eq '2023-06-01'"), ("his-wlk-cal", None), ("kse-load", None)]:
    p = {"$filter": flt, "$first": 2} if flt else {"$first": 2}
    try:
        r = requests.get(PSE + ep, params=p, headers={"User-Agent": UA}, timeout=40)
        print("PSE", r.status_code, ep, flt, r.text[:700].replace("\n", " "), flush=True)
    except Exception as e:  # noqa: BLE001
        print("PSE ERR", ep, repr(e)[:100], flush=True)
