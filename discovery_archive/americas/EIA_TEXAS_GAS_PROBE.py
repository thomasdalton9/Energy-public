"""Probe EIA API v2 for Texas natural gas: consumption by sector (cons/sum), production (prod/sum), interstate
movement (move/state), pipeline exports by port (move/poe1, poe2), storage (stor/sum). Prints facets and recent rows."""
import json, os, requests
KEY = os.environ["EIA_API_KEY"]
B = "https://api.eia.gov/v2/natural-gas/"


def get(route, params=None):
    r = requests.get(B + route, params={"api_key": KEY, **(params or {})}, timeout=90)
    print(f"\n=== {route} {params} -> {r.status_code}")
    try:
        return r.json()
    except Exception:
        print(r.text[:500]); return {}


for route in ("cons/sum", "prod/sum", "move/state", "move/poe1", "move/poe2", "move/expc", "stor/sum"):
    j = get(route).get("response", {})
    print({k: j.get(k) for k in ("frequency", "startPeriod", "endPeriod")})
    print("data cols:", list(j.get("data", {}).keys()) if isinstance(j.get("data"), dict) else j.get("data"))
    for f in j.get("facets", []) if isinstance(j.get("facets"), list) else []:
        print("facet", f)
    for f in ("duoarea", "process", "product", "series"):
        jj = get(f"{route}/facet/{f}").get("response", {})
        fac = jj.get("facets", [])
        tx = [x for x in fac if "TX" in str(x.get("id")) or "Texas" in str(x.get("name")) or "TEXAS" in str(x.get("name")).upper()]
        print(f"  facet {f}: n={len(fac)} texas-ish={tx}")
        if f == "process" or len(fac) <= 40:
            print("   ", [(x.get("id"), x.get("name")) for x in fac][:80])


def rows(route, extra, label):
    p = {"frequency": "monthly", "data[0]": "value", "start": "2024-06", "length": 5000,
         "sort[0][column]": "period", "sort[0][direction]": "desc", **extra}
    j = get(route, p).get("response", {})
    d = j.get("data", [])
    print(label, "total", j.get("total"), "n", len(d))
    return d


# consumption, Texas
d = rows("cons/sum/data/", {"facets[duoarea][]": "STX"}, "cons STX")
procs = {}
for x in d:
    procs.setdefault((x["process"], x.get("process-name"), x.get("units")), []).append((x["period"], x["value"]))
for k, v in procs.items():
    print(k, sorted(v, reverse=True)[:5])
# production
d = rows("prod/sum/data/", {"facets[duoarea][]": "STX"}, "prod STX")
procs = {}
for x in d:
    procs.setdefault((x["process"], x.get("process-name"), x.get("units")), []).append((x["period"], x["value"]))
for k, v in procs.items():
    print(k, sorted(v, reverse=True)[:4])
# state to state movement involving Texas
d = rows("move/state/data/", {}, "move/state all (recent)")
print(json.dumps(d[:3], indent=1))
tx = [x for x in d if "TX" in str(x.get("duoarea")) or "Texas" in str(x.get("area-name"))]
print("TX rows", len(tx)); print(json.dumps(tx[:6], indent=1))
# poe1 / poe2
for route in ("move/poe1/data/", "move/poe2/data/"):
    d = rows(route, {}, route)
    print(json.dumps(d[:3], indent=1))
    ports = {}
    for x in d:
        ports.setdefault((x.get("duoarea"), x.get("area-name"), x.get("process"), x.get("process-name"), x.get("product"), x.get("units")), []).append((x["period"], x["value"]))
    for k, v in ports.items():
        if any(s in str(k) for s in ("TX", "Texas", "TEXAS")) or "Corpus" in str(k) or "Freeport" in str(k) or "Eagle" in str(k):
            print(k, sorted(v, reverse=True)[:3])
    print("all ports sample:", list(ports)[:60])
# storage Texas
d = rows("stor/sum/data/", {"facets[duoarea][]": "STX"}, "stor STX")
procs = {}
for x in d:
    procs.setdefault((x["process"], x.get("process-name"), x.get("units")), []).append((x["period"], x["value"]))
for k, v in procs.items():
    print(k, sorted(v, reverse=True)[:3])
