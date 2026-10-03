"""One-off probe: what Elexon BMRS (and NESO) offer for Great Britain power - fuel mix, demand, interconnectors,
embedded wind/solar - and how far back the history goes. Prints samples; no files written."""
import json
import sys
import requests

B = "https://data.elexon.co.uk/bmrs/api/v1"
S = requests.Session()
S.headers["Accept"] = "application/json"


def get(url, **params):
    try:
        r = S.get(url, params=params, timeout=90)
        print(f"\nGET {r.url}\n -> {r.status_code} {len(r.content)} bytes")
        return r
    except Exception as e:  # noqa: BLE001
        print(f"\nGET {url} {params}\n -> ERROR {type(e).__name__}: {e}")


def show(r, n=3):
    if r is None or r.status_code != 200:
        print((r.text[:400] if r is not None else ""))
        return None
    try:
        j = r.json()
    except Exception:  # noqa: BLE001
        print(r.text[:400])
        return None
    rows = j["data"] if isinstance(j, dict) and "data" in j else j
    print(f" rows: {len(rows) if hasattr(rows, '__len__') else '?'}")
    for x in (rows[:n] if isinstance(rows, list) else [rows]):
        print(" ", json.dumps(x)[:400])
    return rows


print("=== FUELHH (half-hourly generation by fuel incl. interconnectors) ===")
rows = show(get(f"{B}/datasets/FUELHH/stream", settlementDateFrom="2026-09-28", settlementDateTo="2026-09-28", format="json"), 4)
if isinstance(rows, list):
    print(" fuel types:", sorted({x.get("fuelType") for x in rows}))
for d in ("2021-01-01", "2019-01-01", "2016-01-01"):
    print("history check", d)
    show(get(f"{B}/datasets/FUELHH/stream", settlementDateFrom=d, settlementDateTo=d, format="json"), 1)

print("=== demand outturn (INDO / ITSDO) ===")
show(get(f"{B}/demand/outturn/stream", settlementDateFrom="2026-09-28", settlementDateTo="2026-09-28", format="json"), 3)
show(get(f"{B}/datasets/INDO/stream", publishDateTimeFrom="2026-09-28T00:00Z", publishDateTimeTo="2026-09-28T06:00Z", format="json"), 3)
show(get(f"{B}/demand/outturn/summary", format="json"), 2)

print("=== generation outturn ===")
show(get(f"{B}/generation/outturn/summary", format="json"), 2)
show(get(f"{B}/generation/actual/per-type/day-total", settlementDateFrom="2026-09-20", settlementDateTo="2026-09-28", format="json"), 3)

print("=== embedded wind and solar (forecast/outturn) ===")
show(get(f"{B}/forecast/generation/wind-and-solar/day-ahead", format="json"), 3)
show(get(f"{B}/datasets/DGWS/stream", publishDateTimeFrom="2026-09-27T00:00Z", publishDateTimeTo="2026-09-28T00:00Z", format="json"), 3)

print("=== NESO historic demand (embedded wind/solar, transmission system demand) ===")
r = get("https://api.neso.energy/api/3/action/package_show", id="historic-demand-data")
if r is not None and r.status_code == 200:
    for res in r.json()["result"]["resources"]:
        print(" ", res.get("name"), res.get("format"), res.get("id"), res.get("url"))
r = get("https://api.neso.energy/api/3/action/datastore_search", resource_id="b2bde559-3455-4021-b179-dfe60c0337b0", limit=2)
if r is not None and r.status_code == 200:
    print(json.dumps(r.json()["result"]["fields"])[:900])
    print(json.dumps(r.json()["result"]["records"][:2])[:600])
sys.exit(0)
