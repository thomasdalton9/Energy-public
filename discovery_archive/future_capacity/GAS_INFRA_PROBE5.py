"""Probe 5: ENTSOG TYNDP 2024 Annex A (list of projects) and Annex C1 (gas infrastructure capacities)."""
import io, os, re
import pandas as pd
import requests
OUT = "discovery_archive/results/future_capacity"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
for name, u in [("tyndp2024_annex_a.xlsx", "https://www.entsog.eu/sites/default/files/2024-07/TYNDP%202024%20Annex%20A%20-%20List%20of%20projects.xlsx"),
                ("tyndp2024_annex_c1.xlsx", "https://www.entsog.eu/sites/default/files/2024-12/TYNDP%202024_Annex%20C1_Natural%20Gas%20Infrastructure%20Capacities.xlsx")]:
    r = requests.get(u, headers=UA, timeout=(10, 180))
    print(r.status_code, len(r.content), r.headers.get("last-modified"), name)
    if r.status_code != 200: continue
    if len(r.content) < 9_000_000: open(os.path.join(OUT, name), "wb").write(r.content)
    x = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None)
    for sh, d in x.items():
        print(" SHEET", sh, d.shape)
        for i in range(min(6, len(d))): print("   ", [str(v)[:40] for v in d.iloc[i].tolist()[:30]])
