"""Probe 6: Hydro-Quebec reservoir level series in the hydrometeorological dataset."""
import requests

H = {"User-Agent": "Mozilla/5.0"}
base = "https://donnees.hydroquebec.com/api/explore/v2.1/catalog/datasets/donnees-hydrometeorologiques/records"
r = requests.get(base, params={"group_by": "composition_depil_type_point_donnee", "limit": 100}, headers=H, timeout=60)
print(r.status_code, [x["composition_depil_type_point_donnee"] for x in r.json().get("results", [])])
r = requests.get(base, params={"where": "composition_depil_type_point_donnee like 'Niveau'", "limit": 5,
                               "order_by": "date desc"}, headers=H, timeout=60)
print(r.status_code, r.text[:2500])
r = requests.get(base, params={"where": "composition_depil_type_point_donnee like 'Niveau'", "group_by": "nom", "limit": 100}, headers=H, timeout=60)
print(r.status_code, [x["nom"] for x in r.json().get("results", [])])
r = requests.get(base, params={"where": "composition_depil_type_point_donnee like 'Niveau' and nom like 'Caniapiscau'",
                               "select": "min(date), max(date), count(*)"}, headers=H, timeout=60)
print(r.status_code, r.text[:500])
for u in ["https://www.hydro.mb.ca/hydrologicalData/static/", "https://www.hydro.mb.ca/hydrologicalData/static/inc/webpublic.js"]:
    r = requests.get(u, headers=H, timeout=60)
    import re
    print(u, r.status_code, sorted(set(re.findall(r'["\'](/?[\w./-]*(?:KiWIS|api|json|data)[\w./?=&-]*)["\']', r.text)))[:20])
