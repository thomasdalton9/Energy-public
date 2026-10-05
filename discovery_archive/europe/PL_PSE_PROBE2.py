"""Probe 2: PSE open-data API: $metadata entity list, fields, history depth (business_date filters), generation by fuel and cross-border exchange entities."""
import re
import requests

UA = "Mozilla/5.0"
B = "https://api.raporty.pse.pl/api/"
try:
    r = requests.get(B + "$metadata", headers={"User-Agent": UA}, timeout=60)
    print("META", r.status_code, len(r.text), flush=True)
    names = sorted(set(re.findall(r'EntitySet Name="([^"]+)"', r.text)))
    print("ENTITIES", names, flush=True)
    for m in re.finditer(r'<EntityType Name="([^"]+)">(.*?)</EntityType>', r.text, re.S):
        props = re.findall(r'Property Name="([^"]+)"', m.group(2))
        print("TYPE", m.group(1), props, flush=True)
except Exception as e:  # noqa: BLE001
    print("META ERR", repr(e)[:200], flush=True)
for ep in ["his-wlk-cal", "kse-load", "his-gen-pal", "pk5l-wp", "gen-jw", "his-gen-jw", "gen-paliwo", "przeplywy-mocy", "his-gen-pal-5", "gen-moc-jw"]:
    for d in ["2021-06-01", "2022-06-01", "2023-06-01", "2024-06-01"]:
        try:
            r = requests.get(B + ep, params={"$filter": f"business_date eq '{d}'", "$first": 3}, headers={"User-Agent": UA}, timeout=40)
            print("H", ep, d, r.status_code, r.text[:260].replace("\n", " "), flush=True)
        except Exception as e:  # noqa: BLE001
            print("H ERR", ep, d, repr(e)[:80], flush=True)
        if r.status_code == 404:
            break
