"""Probe: the CDS catalogue entry, request form and constraints of sis-energy-derived-reanalysis, and licence status. No data download."""
import json, os, requests
B = "https://cds.climate.copernicus.eu/api"
ID = "sis-energy-derived-reanalysis"
for ep in ("catalogue/v1/collections/" + ID, "catalogue/v1/collections/" + ID + "/form.json",
           "catalogue/v1/collections/" + ID + "/constraints.json"):
    r = requests.get(f"{B}/{ep}", timeout=60)
    print("=====", ep, r.status_code, len(r.text))
    try:
        j = r.json()
    except Exception:
        print(r.text[:1500]); continue
    s = json.dumps(j, indent=1)
    print(s[:9000] if "form" in ep else s[:5000])
key = os.environ.get("CDS_API_KEY", "")
print("key present:", bool(key), len(key))
h = {"PRIVATE-TOKEN": key}
r = requests.get(f"{B}/profiles/v1/account/licences", headers=h, timeout=60)
print("licences", r.status_code, r.text[:3000])
r = requests.get(f"{B}/retrieve/v1/processes/{ID}", headers=h, timeout=60)
print("process", r.status_code, r.text[:500])
