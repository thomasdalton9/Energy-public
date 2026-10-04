"""
Pakistan power discovery, round 11: end-to-end test of asia/PAKISTAN_NEPRA.py with K-Electric added (temp path, run
twice), then: KE sheet as parsed, monthly GWh grid vs incl. KE, calendar-year TWh before / after KE, KE basis.
"""
import os
import subprocess
import sys
import tempfile
import time

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def main():
    d = tempfile.mkdtemp()
    out = os.path.join(d, "pk.xlsx")
    for run in (1, 2):
        t = time.time()
        print(f"\n######## run {run}", flush=True)
        subprocess.run([sys.executable, os.path.join(ROOT, "asia", "PAKISTAN_NEPRA.py"), "--out", out, "--no-hourly"],
                       check=False)
        print(f"run {run}: {time.time() - t:.0f}s", flush=True)
    pd.set_option("display.width", 300)
    pd.set_option("display.max_rows", 300)
    pd.set_option("display.max_columns", 40)
    x = pd.read_excel(out, sheet_name="Daily", index_col=0)
    ke = pd.read_excel(out, sheet_name="KE", index_col=0)
    f = pd.read_excel(out, sheet_name="Files", index_col=0)
    print("\nKE sheet (GWh as filed):\n" + ke.drop(columns=["KE_file"], errors="ignore").round(1).to_string())
    print("\nKE files:\n" + f[f["kind"].str.startswith(("KE", "NEPRA KE"))][["name", "months", "status"]].to_string())
    cols = ["Total_grid_MWh", "KE_own_MWh", "KE_purchases_nonCPPA_MWh", "KE_from_CPPA_MWh", "Total_MWh", "Gas_MWh",
            "Oil_MWh", "Coal_MWh", "Solar_MWh"]
    g = (x[[c for c in cols if c in x]] / 1000).round(0)
    g["KE_basis"] = x["KE_basis"].str[:60]
    print("\nMonthly GWh:\n" + g.to_string())
    y = x.groupby(x.index.year).agg(grid=("Total_grid_MWh", "sum"), total=("Total_MWh", "sum"),
                                    months=("Total_MWh", "count"))
    print("\nCalendar-year TWh (grid only -> incl. KE):\n" + (y.assign(grid=y.grid / 1e6, total=y.total / 1e6)).round(1)
          .to_string())


if __name__ == "__main__":
    main()
