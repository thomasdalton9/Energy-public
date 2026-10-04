"""
South & Southeast Asia gas, round 11 (GitHub Actions): after the column-position fix (and case-insensitive month headings, e.g. "OcT.,") in PAKISTAN_PBS_GAS.parse_mbs
('-' = missing value, figures placed under their month heading, row-sum check), parse three MBS issues and show the
national / provincial rows and any rejected rows; then re-run all three pulls twice (round 6 test).
"""
import os
import runpy
import sys

import requests
import urllib3

urllib3.disable_warnings()
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, "asia"))
sys.path.insert(0, ROOT)
import PAKISTAN_PBS_GAS as P  # noqa: E402

for f in ("MBS_Sep_2024.pdf", "MBS_April_2023.pdf", "MBS_Apr_2023.pdf", "MBS_Apr_2022.pdf"):
    u = "https://www.pbs.gov.pk/wp-content/uploads/2020/07/" + f
    r = requests.get(u, headers=P.H, timeout=(20, 240), verify=False)
    if r.status_code != 200:
        print(f"{f}: {r.status_code}", flush=True)
        continue
    rej = []
    t, fl = P.parse_mbs(r.content, rej)
    print(f"\n### {f}: {len(t)} months, {sum(len(v) for v in fl.values())} field-month values, {len(rej)} rejected",
          flush=True)
    for d in sorted(t):
        print(f"  {d:%Y-%m} {t[d]}", flush=True)
    for x in rej[:15]:
        print(f"  REJECTED {x}", flush=True)

runpy.run_path(os.path.join(HERE, "GAS_SSEA_DISCOVERY6.py"), run_name="__main__")
