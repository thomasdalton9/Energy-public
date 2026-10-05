"""Sanity check: ERCOT gas burn estimate (ercot_gas_burn_daily.xlsx) vs EIA state-level Texas gas for electric power
(EIA API natural-gas/cons/sum, process VEU, duoarea STX, MMcf/month). Texas is larger than ERCOT, so expect Texas > ERCOT.
Run in GitHub Actions only (EIA_API_KEY secret)."""
import os

import pandas as pd
import requests

URL = "https://api.eia.gov/v2/natural-gas/cons/sum/data/"
key = os.environ["EIA_API_KEY"]
rows, offset = [], 0
while True:
    r = requests.get(URL, params={"api_key": key, "frequency": "monthly", "data[0]": "value", "facets[process][]": "VEU",
                                  "facets[duoarea][]": "STX", "start": "2021-01", "length": 5000, "offset": offset},
                     timeout=60)
    r.raise_for_status()
    j = r.json()["response"]
    rows += j["data"]
    offset += len(j["data"])
    if offset >= int(j["total"]) or not j["data"]:
        break
tx = pd.DataFrame(rows)
tx["month"] = pd.to_datetime(tx["period"] + "-01")
tx = tx.set_index("month")["value"].astype(float).sort_index()
tx_bcfd = tx / tx.index.days_in_month / 1e3
for sheet, label in (("ercot_gas_burn_daily.xlsx", "ERCOT"), ("miso_gas_burn_daily.xlsx", "MISO")):
    m = pd.read_excel(f"output/Data and Chart Outputs/{sheet}", sheet_name="Monthly")
    m["month"] = pd.to_datetime(m["month"])
    m = m.set_index("month")
    out = pd.DataFrame({f"{label}_Bcf_d": m["Gas_burn_Bcf_per_day"], "Texas_EIA_VEU_Bcf_d": tx_bcfd}).dropna(subset=[f"{label}_Bcf_d"])
    out["ratio_est_to_TX"] = out[f"{label}_Bcf_d"] / out["Texas_EIA_VEU_Bcf_d"]
    out["basis"] = m["Heat_rate_basis"].str[:25]
    print(out.round(3).to_string())
    print()
