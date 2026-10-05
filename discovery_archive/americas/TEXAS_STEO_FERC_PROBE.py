"""
Probe (manual, Actions only): (1) EIA API v2 steo route - series ids for regional dry gas production (Permian, Eagle
Ford, Haynesville, ...) with their latest history and forecast values; (2) FERC LNG pages - links and any text on the
Texas projects (Rio Grande, Port Arthur, Golden Pass, Corpus Christi Stage 3) to source start dates.
Result recorded in americas/TEXAS_GAS_FORECAST.py docstring.
"""
import os
import re

import requests

KEY = os.environ.get("EIA_API_KEY", "")


def eia(route, params=None):
    r = requests.get("https://api.eia.gov/v2/" + route, params={"api_key": KEY, **(params or {})}, timeout=(10, 90))
    print(route, r.status_code)
    r.raise_for_status()
    return r.json()["response"]


print("=== STEO facets")
try:
    j = eia("steo/facet/seriesId")
    fac = j["facets"]
    print(len(fac), "series")
    for f in fac:
        t = (f["id"] + " " + str(f.get("name", ""))).lower()
        if re.search(r"permian|eagle|haynes|marcellus|bakken|anadarko|niobrara|appalachia|dry natural gas prod|marketed natural gas prod|natural gas.*(texas|gulf)|ngpr|ngmp", t):
            print("  ", f["id"], "|", f.get("name"))
except Exception as e:  # noqa: BLE001
    print("facet failed", e)

print("=== STEO data for candidate ids")
try:
    j = eia("steo/facet/seriesId")
    ids = [f["id"] for f in j["facets"] if re.search(r"permian|eagle|haynes|marcellus|bakken|anadarko|niobrara|dry natural gas prod|marketed natural gas", (f["id"] + str(f.get("name", ""))).lower())]
    print(ids)
    for i in ids[:40]:
        d = eia("steo/data/", {"frequency": "monthly", "data[0]": "value", "facets[seriesId][]": i, "start": "2025-09",
                                "sort[0][column]": "period", "sort[0][direction]": "asc", "length": 60})
        rows = d["data"]
        print(i, len(rows), rows[0]["period"] if rows else None, rows[-1]["period"] if rows else None,
              rows[0].get("seriesDescription"), rows[0].get("unit"), [r["value"] for r in rows[::6]])
except Exception as e:  # noqa: BLE001
    print("data failed", e)

print("=== STEO release info")
try:
    print(eia("steo/").get("description"))
    j = eia("steo/data/", {"frequency": "monthly", "data[0]": "value", "facets[seriesId][]": "NGPRPUS", "start": "2026-06",
                            "sort[0][column]": "period", "sort[0][direction]": "asc", "length": 60})
    print([(r["period"], r["value"]) for r in j["data"]])
except Exception as e:  # noqa: BLE001
    print("steo info failed", e)

print("=== FERC")
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120 Safari/537.36"}
pages = ["https://www.ferc.gov/natural-gas/lng", "https://www.ferc.gov/industries-data/natural-gas/overview/lng",
         "https://www.ferc.gov/industries-data/natural-gas/overview/liquefied-natural-gas-lng",
         "https://www.ferc.gov/industries-data/natural-gas/overview/liquefied-natural-gas-lng/lng-construction-status",
         "https://www.ferc.gov/industries-data/natural-gas/overview/liquefied-natural-gas-lng/north-american-lng-export-terminals",
         "https://www.ferc.gov/media/north-american-lng-export-terminals-existing-approved-not-yet-built-and-proposed",
         "https://www.ferc.gov/news-events/news/ferc-staff-lng-monthly-status"]
KW = re.compile(r"rio grande|port arthur|golden pass|corpus christi|texas lng|freeport", re.I)
for u in pages:
    try:
        r = requests.get(u, headers=H, timeout=40)
        print(u, r.status_code, len(r.text))
        if r.status_code != 200:
            continue
        for m in sorted(set(re.findall(r'href="([^"]*)"[^>]*>([^<]{3,120})<', r.text)), key=lambda x: x[0]):
            if re.search(r"lng|liquef|export-terminal|construction", m[0] + m[1], re.I):
                print("    link", m[0][:160], "|", m[1].strip()[:100])
        txt = re.sub(r"<[^>]+>", " ", r.text)
        for m in KW.finditer(txt):
            print("    text:", re.sub(r"\s+", " ", txt[max(0, m.start() - 100):m.end() + 200]))
            break
    except Exception as e:  # noqa: BLE001
        print(u, "failed", type(e).__name__, str(e)[:120])
