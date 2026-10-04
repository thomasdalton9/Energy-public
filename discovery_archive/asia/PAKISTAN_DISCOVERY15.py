"""
Pakistan power discovery, round 15: end-to-end test of asia/PAKISTAN_NEPRA.py, K-Electric real figures only (SOIR tables) (temp path, run
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
    print("\nKE sheet (GWh as filed):\n" + ke.drop(columns=["KE_files"], errors="ignore").round(1).to_string())
    print("\nKE files:\n" + f[f["kind"].str.startswith(("KE", "NEPRA KE"))][["name", "months", "status"]].to_string())
    cols = ["Total_grid_MWh", "KE_own_MWh", "KE_purchases_nonCPPA_MWh", "KE_from_CPPA_MWh", "Total_MWh", "Gas_MWh",
            "Oil_MWh", "Coal_MWh", "Nuclear_MWh", "Solar_MWh"]
    g = (x[[c for c in cols if c in x]] / 1000).round(0)
    g["KE_basis"] = x["KE_basis"].str[:60]
    g["KE_incl"] = x.get("KE_included")
    print("\nMonthly GWh:\n" + g.to_string())
    y = x.groupby(x.index.year).agg(grid=("Total_grid_MWh", "sum"), total=("Total_MWh", "sum"),
                                    months=("Total_MWh", "count"),
                                    ke_months=("KE_included", lambda v: int(v.fillna(False).astype(bool).sum())))
    print("\nCalendar-year TWh (grid only -> incl. KE):\n" + (y.assign(grid=y.grid / 1e6, total=y.total / 1e6)).round(1)
          .to_string())


def _unused():
    pass


def soir_check(out):
    """Fiscal-year sums of the monthly KE columns vs NEPRA State of Industry Report 2025 (KE statistics table)."""
    x = pd.read_excel(out, sheet_name="Daily", index_col=0)
    fy = x.index.year + (x.index.month >= 7)
    ref = pd.DataFrame({"SOIR_own": [10185.60, 7889.96, 7093.34, 7471.36, 6432.61],
                        "SOIR_NTDC": [6118.04, 9036.54, 8960.81, 8538.07, 10234.27],
                        "SOIR_others": [3182.03, 2851.39, 2204.71, 1691.90, None]}, index=[2021, 2022, 2023, 2024, 2025])
    g = pd.DataFrame({"months_KE": x["KE_included"].fillna(False).astype(bool).groupby(fy).sum(),
                      "own": x["KE_own_MWh"].groupby(fy).sum() / 1000,
                      "from_CPPA": x["KE_from_CPPA_MWh"].groupby(fy).sum() / 1000,
                      "nonCPPA": x["KE_purchases_nonCPPA_MWh"].groupby(fy).sum() / 1000})
    if "KE_own_decision_MWh" in x:
        b = x[["KE_own_MWh", "KE_own_decision_MWh"]].dropna()
        b = b.assign(diff_pct=(b.KE_own_decision_MWh / b.KE_own_MWh - 1) * 100)
        print("\nKE own: SOIR vs NEPRA decision (GWh):\n" + (b / [1000, 1000, 1]).round(1).to_string())
        print(f"mean abs diff {b.diff_pct.abs().mean():.1f}%, max {b.diff_pct.abs().max():.1f}%")
    print("\nFiscal years (Jul-Jun, labelled by the June year), GWh, vs State of Industry Report 2025:\n" +
          g.join(ref, how="outer").round(1).to_string())


if __name__ == "__main__":
    main()
    import glob
    for f in glob.glob(tempfile.gettempdir() + "/tmp*/pk.xlsx"):
        soir_check(f)
