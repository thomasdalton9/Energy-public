"""
Pakistan power discovery, round 6: end-to-end test of asia/PAKISTAN_NEPRA.py into a temp path (twice, to check the
incremental second run reads nothing new), then a summary: monthly GWh by fuel, calendar-year TWh, sources, checks.
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
    out, hout = os.path.join(d, "pk.xlsx"), os.path.join(d, "pk_hourly.xlsx")
    for run in (1, 2):
        t = time.time()
        print(f"\n######## run {run}", flush=True)
        subprocess.run([sys.executable, os.path.join(ROOT, "asia", "PAKISTAN_NEPRA.py"), "--out", out,
                        "--hourly-out", hout], check=False)
        print(f"run {run}: {time.time() - t:.0f}s", flush=True)
    pd.set_option("display.width", 250)
    pd.set_option("display.max_rows", 200)
    x = pd.read_excel(out, sheet_name="Daily", index_col=0)
    m = pd.read_excel(out, sheet_name="Months", index_col=0)
    f = pd.read_excel(out, sheet_name="Files", index_col=0)
    g = (x / 1000).round(0)
    print("\nMonthly GWh:\n" + g[["Hydro_MWh", "Coal_MWh", "Coal_local_MWh", "Coal_imported_MWh", "Gas_local_MWh",
                                "RLNG_MWh", "RFO_MWh", "HSD_MWh", "Nuclear_MWh", "Wind_MWh", "Solar_MWh",
                                "Bioenergy_MWh", "Other_MWh", "Total_MWh", "Imports_MWh"]].to_string())
    y = x.groupby(x.index.year).agg(["sum", "count"])
    tot = pd.DataFrame({"TWh": (y[("Total_MWh", "sum")] / 1e6).round(1), "months": y[("Total_MWh", "count")]})
    print("\nCalendar-year totals:\n" + tot.to_string())
    print("\nMonths sheet:\n" + m[["Source", "Check_pct", "Filing_total_MWh"]].to_string())
    print("\nFiles sheet status counts:\n" + f["status"].value_counts().to_string())
    print(f[f["status"] != "ok"][["name", "status"]].to_string())
    if os.path.exists(hout):
        h = pd.read_excel(hout, sheet_name="Daily", index_col=0)
        print("\nHourly file, monthly GWh:\n" + (h.resample("MS").sum() / 1000).round(0).to_string())
        print(pd.ExcelFile(hout).sheet_names)


if __name__ == "__main__":
    main()
