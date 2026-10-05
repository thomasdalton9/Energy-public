"""List ENTSOG interconnection rows for given pointKeys and country pairs (labels, operators, far-side country/type). Prints only."""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, "europe"))
import ENTSOG_GAS_FLOWS_DAILY as G  # noqa: E402

KEYS = set(os.environ.get("KEYS", "").split(","))
PAIRS = {frozenset(x.split("-")) for x in os.environ.get("PAIRS", "").split(",") if x}
ics = G.get_json("interconnections", {"limit": -1}).get("interconnections", [])
print(len(ics), "rows")
for i in ics:
    fc, tc = i.get("fromCountryKey"), i.get("toCountryKey")
    if i["pointKey"] in KEYS or (frozenset([fc, tc]) in PAIRS):
        print(f"{i['pointKey']} | {i.get('pointLabel')} | from {i.get('fromOperatorKey')} {fc} {i.get('fromInfrastructureTypeLabel')} -> to {i.get('toOperatorKey')} {tc} {i.get('toInfrastructureTypeLabel')}")
