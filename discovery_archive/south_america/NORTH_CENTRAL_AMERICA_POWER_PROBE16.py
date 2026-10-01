"""
Round 16 (Oct-2026), El Salvador. SIGET's yearly 'Boletin de Estadisticas
Electricas' (PROBE15) has annual tables only. Its 'Visualizador dinamico de
Estadisticas Electricas' is a public Power BI report:
  https://app.powerbi.com/view?r=eyJrIjoiYmM0MDJhMTktNmQ2YS00ZDkxLTgyYWUtMDUxYjQ5MjQxYWY5IiwidCI6IjE1OTg0YzZmLTNiMmMtNDZmMi1iNmU0LTUzNDQ3MGY2MDVmNyJ9
Public reports answer the same anonymous 'modelsAndExploration' /
'conceptualschema' / 'querydata' calls the browser makes. Here: find the
cluster, list the model's tables and columns (is there a monthly generation /
injection by resource table?).

Usage: python3 NORTH_CENTRAL_AMERICA_POWER_PROBE16.py pbi
"""

print("STARTING", flush=True)

import base64
import json
import re
import sys

import requests

TOKEN = "eyJrIjoiYmM0MDJhMTktNmQ2YS00ZDkxLTgyYWUtMDUxYjQ5MjQxYWY5IiwidCI6IjE1OTg0YzZmLTNiMmMtNDZmMi1iNmU0LTUzNDQ3MGY2MDVmNyJ9"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def pbi():
    tok = json.loads(base64.b64decode(TOKEN + "=" * (-len(TOKEN) % 4)))
    key, tenant = tok["k"], tok["t"]
    print("key", key, "tenant", tenant, flush=True)
    hdr = dict(H, **{"X-PowerBI-ResourceKey": key, "Accept": "application/json"})
    view = requests.get("https://app.powerbi.com/view", params={"r": TOKEN}, headers=H, timeout=60)
    clusters = sorted(set(re.findall(r"https://[a-z0-9-]+\.analysis\.windows\.net", view.text)))
    print("clusters in view page:", clusters, flush=True)
    cands = clusters + ["https://wabi-us-north-central-api.analysis.windows.net",
                        "https://wabi-us-east2-api.analysis.windows.net",
                        "https://wabi-south-central-us-api.analysis.windows.net",
                        "https://wabi-west-us-api.analysis.windows.net",
                        "https://wabi-us-east-a-primary-api.analysis.windows.net"]
    r = requests.get("https://api.powerbi.com/public/routing/cluster/" + tenant, headers=hdr, timeout=30)
    print("routing:", r.status_code, r.text[:300], flush=True)
    try:
        fixed = r.json().get("FixedClusterUri")
        if fixed:
            cands.insert(0, fixed.rstrip("/").replace("-redirect.", "-api.").replace("redirect.", "api."))
    except ValueError:
        pass
    for base in cands:
        u = f"{base}/public/reports/{key}/modelsAndExploration?preferReadOnlySession=true"
        try:
            r = requests.get(u, headers=hdr, timeout=30)
        except requests.RequestException as e:
            print(" ", base, e, flush=True)
            continue
        print(f"  {base}: [{r.status_code}] {len(r.text)} chars", flush=True)
        if not r.ok:
            continue
        j = r.json()
        models = j.get("models", [])
        print("  models:", [(m.get("id"), m.get("dbName")) for m in models], flush=True)
        ex = j.get("exploration", {})
        for sec in ex.get("sections", [])[:20]:
            vis = []
            for vc in sec.get("visualContainers", [])[:30]:
                try:
                    cfg = json.loads(vc.get("config", "{}"))
                    sv = cfg.get("singleVisual", {})
                    proj = sv.get("projections", {})
                    vis.append((sv.get("visualType"), [p.get("queryRef") for v in proj.values() for p in v][:6]))
                except ValueError:
                    pass
            print(f"  PAGE {sec.get('displayName')}: {vis[:12]}", flush=True)
        if models:
            body = {"modelIds": [models[0]["id"]], "userPreferredLocale": "es-ES"}
            cs = requests.post(f"{base}/public/reports/conceptualschema", headers=dict(hdr, **{"Content-Type": "application/json"}),
                               data=json.dumps(body), timeout=60)
            print("  conceptualschema:", cs.status_code, flush=True)
            if cs.ok:
                for s in cs.json().get("schemas", []):
                    for ent in s.get("schema", {}).get("Entities", []):
                        props = [p.get("Name") for p in ent.get("Properties", [])]
                        print(f"   ENTITY {ent.get('Name')}: {props[:40]}", flush=True)
        break


if __name__ == "__main__":
    {"pbi": pbi}[sys.argv[1]]()
    print("DONE", flush=True)
