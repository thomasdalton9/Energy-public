"""One-off probe: where does complete Ireland generation come from?
 (1) ENTSO-E A75 actual generation per type for the SEM bidding zone vs the Irish control area (IE, 10YIE-1001A00010) vs
     Northern Ireland (10Y1001A1001A016) and the A65 load for each - annual-ish samples, TWh;
 (2) EirGrid Smart Grid Dashboard chartTypes/areas that might carry fuel mix or solar (no key);
 (3) SEMOpx / EirGrid public report endpoints.
Prints only; no files written."""
import json
import os
import sys
from datetime import datetime, timedelta, timezone

import requests

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "europe"))
import entsoe_common as C  # noqa: E402

AREAS = {"SEM (all-island bidding zone)": "10Y1001A1001A59C", "IE (control area)": "10YIE-1001A00010",
         "NI (SONI control area)": "10Y1001A1001A016"}
SAMPLES = [("2022-03-01", "2022-04-01"), ("2024-03-01", "2024-04-01"), ("2026-03-01", "2026-04-01")]


def sample(eic, doc, d0, d1):
    """A75 -> TWh by fuel group (generation direction only); A65 -> TWh load. None = no data."""
    t0 = datetime.fromisoformat(d0).replace(tzinfo=timezone.utc)
    t1 = datetime.fromisoformat(d1).replace(tzinfo=timezone.utc)
    try:
        if doc == "A75":
            res = C.fetch_split({"documentType": "A75", "processType": "A16", "in_Domain": eic}, t0, t1,
                                C.parse_generation, C.merge_generation)
        else:
            res = C.fetch_split({"documentType": "A65", "processType": "A16", "outBiddingZone_Domain": eic}, t0, t1,
                                C.parse_energy, C.merge_energy)
    except Exception as e:  # noqa: BLE001
        return f"ERR {type(e).__name__}: {str(e)[:150]}"
    if not res:
        return None
    if doc == "A65":
        return round(sum(res["e"].values()) / 1e6, 3)
    by = {}
    for (psr, direction), days in res["e"].items():
        key = f"{C.GEN_COLUMN.get(psr, psr)}{'' if direction == 'in' else '(out)'}"
        by[key] = round(by.get(key, 0) + sum(days.values()) / 1e6, 3)
    return by


print("=== (1) ENTSO-E ===", flush=True)
for name, eic in AREAS.items():
    for d0, d1 in SAMPLES:
        gen = sample(eic, "A75", d0, d1)
        load = sample(eic, "A65", d0, d1)
        tot = round(sum(v for k, v in gen.items() if "(out)" not in k), 3) if isinstance(gen, dict) else None
        print(f"{name} {d0}: generation TWh {tot} {json.dumps(gen)[:500] if isinstance(gen, dict) else gen} | load TWh {load}", flush=True)

print("\n=== (2) EirGrid Smart Grid Dashboard ===", flush=True)
H = {"User-Agent": "Mozilla/5.0", "Accept": "application/json", "Referer": "https://www.smartgriddashboard.com/"}
B = "https://www.smartgriddashboard.com/api/chart/"
for ct, areas in (("generation", "fuelmix"), ("generation", "fuelmixsolar"), ("solar", "solaractual"), ("interconnection", "interconnection"),
                  ("interconnection", "interconnectionactual"), ("generation", "generationactual"), ("snsp", "snsp"),
                  ("fuelmix", "fuelmix"), ("generation", "fuelmix,solaractual")):
    for rng in ("2026-09-20", ):
        try:
            r = requests.get(B, params={"region": "ALL", "chartType": ct, "dateRange": "day", "dateFrom": "20-Sep-2026",
                                        "dateTo": "21-Sep-2026", "areas": areas}, headers=H, timeout=60)
            txt = r.text[:600].replace("\n", " ")
            print(f"{ct}/{areas}: {r.status_code} {len(r.content)}B {txt}")
        except Exception as e:  # noqa: BLE001
            print(f"{ct}/{areas}: ERR {e}")

print("\n=== (3) public report endpoints ===", flush=True)
for u in ("https://www.eirgrid.ie/grid/system-and-renewable-data-reports", "https://www.semopx.com/market-data/",
          "https://reports.sem-o.com/", "https://www.soni.ltd.uk/customer-and-industry/general-data-library/"):
    try:
        r = requests.get(u, headers=H, timeout=60)
        links = [l for l in __import__("re").findall(r'href="([^"]+\.(?:xlsx|csv|xls|zip)[^"]*)"', r.text)][:12]
        print(f"{u}: {r.status_code} {len(r.content)}B links: {links}")
    except Exception as e:  # noqa: BLE001
        print(f"{u}: ERR {e}")
sys.exit(0)
