"""Probe 4 (gas residual, France): ODRE biomethane by receiving network operator (grx_demandeur), annual sums, to split transmission-connected from distribution-connected injection."""
import json, os
import requests
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results", "residual")
B = "https://odre.opendatasoft.com/api/explore/v2.1/catalog/datasets/odre-prod-grdgrt-operateur-def/records"
res = {}
for yr in (2022, 2023, 2024, 2025, 2026):
    r = requests.get(B, params={"select": "grx_demandeur, operateur_de_transport, sum(production_biomethane) as v, count(*) as n", "group_by": "grx_demandeur, operateur_de_transport",
                                "where": f"date>=date'{yr}-01-01' and date<=date'{yr}-12-31'", "limit": 100}, timeout=90)
    res[yr] = r.json().get("results", []) if r.ok else r.status_code
    print(yr, r.status_code, json.dumps(res[yr], ensure_ascii=False)[:1500], flush=True)
json.dump(res, open(os.path.join(OUT, "odre_biom_operators.json"), "w"), ensure_ascii=False, indent=1)
# monthly 2025 sums for the full network and per requester
r = requests.get(B, params={"select": "date_format(date,'YYYY-MM') as m, grx_demandeur, sum(production_biomethane) as v", "group_by": "m, grx_demandeur",
                            "where": "date>=date'2025-01-01'", "limit": 1000, "order_by": "m"}, timeout=90)
open(os.path.join(OUT, "odre_biom_monthly_operator.json"), "w").write(r.text)
print("monthly", r.status_code, len(r.text))
