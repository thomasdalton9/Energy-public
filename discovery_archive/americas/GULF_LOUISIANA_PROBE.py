"""Probe (manual, Actions only) for americas/GULF_COAST_BALANCE.py: (1) EIA API v2 Louisiana series (consumption by sector,
production, LNG exports by port, state production TX/LA for the Haynesville split); (2) LNG project status pages that the Claude
sandbox cannot open (EIA, DOE, FERC, companies): link census + keyword snippets, and any EIA liquefaction-capacity workbook."""
import io
import os
import re
import sys

import requests

KEY = os.environ.get("EIA_API_KEY", "")
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120 Safari/537.36"}


def eia(route, params=None):
    r = requests.get("https://api.eia.gov/v2/" + route, params={"api_key": KEY, **(params or {})}, timeout=(10, 120))
    r.raise_for_status()
    return r.json()["response"]


def rows(route, facets, start="2020-01"):
    p = {"frequency": "monthly", "data[0]": "value", "start": start, "length": 5000,
         "sort[0][column]": "period", "sort[0][direction]": "asc"}
    for k, v in facets.items():
        for i, x in enumerate(v):
            p[f"facets[{k}][{i}]"] = x
    return eia(route, p)["data"]


print("=== A1 move/poe2 series with LA / Louisiana")
try:
    for f in eia("natural-gas/move/poe2/facet/series")["facets"]:
        n = str(f.get("name", ""))
        if re.search(r",\s*LA\b|Louisiana|Sabine|Cameron|Calcasieu|Plaquemines|Cove Point|Elba|Lake Charles", n, re.I):
            print("  ", f["id"], "|", n)
except Exception as e:  # noqa: BLE001
    print("fail", e)

print("=== A2 Louisiana cons/prod/stor latest rows")
for route in ("cons/sum", "prod/sum", "stor/sum"):
    try:
        d = rows(f"natural-gas/{route}/data/", {"duoarea": ["SLA"]}, "2025-06")
        seen = {}
        for r in d:
            seen.setdefault((r.get("process"), r.get("process-name") or r.get("series-description")), []).append((r["period"], r["value"], r.get("units")))
        print(route, len(d))
        for k, v in seen.items():
            print("  ", k, v[-3:])
    except Exception as e:  # noqa: BLE001
        print(route, "fail", e)

print("=== A3 state marketed/dry production TX vs LA vs STEO Haynesville (monthly, 2021-)")
for area in ("STX", "SLA"):
    for pr in ("VGM", "FPD"):
        try:
            d = rows("natural-gas/prod/sum/data/", {"duoarea": [area], "process": [pr]}, "2021-01")
            print(area, pr, len(d), [(r["period"], r["value"]) for r in d[-4:]], d[0].get("units"))
        except Exception as e:  # noqa: BLE001
            print(area, pr, "fail", e)
for sid in ("NGMPHA", "NGMPPM", "NGMPEF"):
    try:
        d = eia("steo/data/", {"frequency": "monthly", "data[0]": "value", "facets[seriesId][]": sid, "start": "2021-01",
                               "sort[0][column]": "period", "sort[0][direction]": "asc", "length": 200})["data"]
        print("STEO", sid, len(d), d[0]["period"], d[-1]["period"], [(r["period"], r["value"]) for r in d[:3]])
    except Exception as e:  # noqa: BLE001
        print("STEO", sid, "fail", e)

print("=== A4 STEO LNG / Louisiana-relevant ids")
try:
    for f in eia("steo/facet/seriesId")["facets"]:
        t = (f["id"] + " " + str(f.get("name", "")))
        if re.search(r"lng|liquef|export", t, re.I):
            print("  ", f["id"], "|", f.get("name"))
except Exception as e:  # noqa: BLE001
    print("fail", e)

print("=== B web pages")
KW = re.compile(r"first LNG|first cargo|start[- ]?up|commercial operations|in service|in-service|FID|final investment|Train \d|"
                r"Bcf/d|mtpa|MTPA|under construction|expected|2026|2027|2028|2029|2030", re.I)
PAGES = [
    "https://www.eia.gov/naturalgas/", "https://www.eia.gov/naturalgas/data.php", "https://www.eia.gov/todayinenergy/",
    "https://www.eia.gov/energyexplained/natural-gas/liquefied-natural-gas.php",
    "https://www.eia.gov/naturalgas/U.S.%20Liquefaction%20Capacity.xlsx",
    "https://www.eia.gov/outlooks/steo/report/natgas.php",
    "https://www.ferc.gov/industries-data/natural-gas/overview/liquefied-natural-gas-lng",
    "https://www.ferc.gov/media/north-american-lng-export-terminals-existing-approved-not-yet-built-and-proposed-8",
    "https://www.energy.gov/fecm/articles/lng-monthly-2026", "https://www.energy.gov/fecm/listings/lng-reports",
    "https://venturegloballng.com/our-projects/plaquemines-lng/", "https://venturegloballng.com/our-projects/cp2-lng/",
    "https://venturegloballng.com/investors/", "https://www.woodside.com/what-we-do/projects/louisiana-lng",
    "https://www.cheniere.com/our-business/sabine-pass-liquefaction", "https://www.cameronlng.com/", "https://www.commonwealthlng.com/",
    "https://www.energytransfer.com/lake-charles-lng", "https://delfinmidstream.com/", "https://www.glenfarnecompany.com/",
    "https://www.rystadenergy.com/", "https://www.cheniere.com/our-business/expansion",
]
for u in PAGES:
    try:
        r = requests.get(u, headers=H, timeout=45)
        ct = r.headers.get("content-type", "")
        print(f"\n## {u} -> {r.status_code} {ct[:40]} {len(r.content)}")
        if r.status_code != 200:
            continue
        if "sheet" in ct or u.endswith("xlsx"):
            import pandas as pd
            for n, d in pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None).items():
                print("  sheet", n, d.shape)
                for _, row in d.iterrows():
                    t = " | ".join(str(v) for v in row if str(v) != "nan")
                    if t:
                        print("   ", t[:300])
            continue
        if "pdf" in ct:
            from pypdf import PdfReader
            for pi, pg in enumerate(PdfReader(io.BytesIO(r.content)).pages[:40]):
                for ln in (pg.extract_text() or "").splitlines():
                    if re.search(r"Louisiana|, LA|Sabine|Cameron|Calcasieu|Plaquemines|CP2|Commonwealth|Delfin|Lake Charles", ln):
                        print(f"   p{pi + 1}: {ln[:250]}")
            continue
        for m in sorted(set(re.findall(r'href="([^"]*)"[^>]*>([^<]{3,120})<', r.text))):
            if re.search(r"liquef|\blng\b|export-terminal|louisiana|plaquemines|cp2|commonwealth|woodside|delfin|lake-charles|\.xlsx|\.pdf", m[0] + m[1], re.I):
                print("   link", m[0][:150], "|", m[1].strip()[:90])
        txt = re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", r.text, flags=re.S))
        n = 0
        for m in KW.finditer(txt):
            print("   txt:", txt[max(0, m.start() - 120):m.end() + 160])
            n += 1
            if n >= 14:
                break
    except Exception as e:  # noqa: BLE001
        print("##", u, "failed", type(e).__name__, str(e)[:120])
