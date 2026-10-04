"""
India daily renewable generation, round 9: end-to-end test of asia/INDIA_RE_DAILY.py (writes to /tmp, not the repo).
Runs the pull twice (second run = incremental merge over the saved sheets), then checks monthly wind / solar
against CEA's monthly RE generation report (Dec 2025: wind 5,009.06 MU, solar 14,946.83; Aug 2026: wind 17,815.20,
solar 18,629.70) and prints the yearly totals.
"""
import subprocess
import sys

import pandas as pd

OUT = "/tmp/india_power_generation_daily.xlsx"
for k in (1, 2):
    print(f"\n######## run {k}", flush=True)
    subprocess.run([sys.executable, "asia/INDIA_RE_DAILY.py", "--out", OUT], check=False, timeout=1400)

x = pd.ExcelFile(OUT)
print("\nsheets:", x.sheet_names)
print(pd.read_excel(OUT, sheet_name="Units").to_string()[:4000])
d = pd.read_excel(OUT, sheet_name="Daily", index_col=0, parse_dates=True)
m = d.resample("MS").sum() / 1000   # GWh = MU
print("\nmonthly MU, last 12 months:\n" + m.tail(12).round(0).to_string())
for mon, wind, solar in (("2025-12-01", 5009.06, 14946.83), ("2026-08-01", 17815.20, 18629.70)):
    if pd.Timestamp(mon) in m.index:
        r = m.loc[mon]
        print(f"{mon[:7]}: wind {r['Wind_MWh']:.0f} vs CEA monthly {wind} ({100 * (r['Wind_MWh'] / wind - 1):+.1f}%), "
              f"solar {r['Solar_MWh']:.0f} vs {solar} ({100 * (r['Solar_MWh'] / solar - 1):+.1f}%)")
dem = pd.read_excel(OUT, sheet_name="Demand", index_col=0, parse_dates=True)
print("\nDemand yearly: energy met TWh / max peak GW\n" + pd.DataFrame({
    "energy_TWh": dem["Energy_met_MWh"].groupby(dem.index.year).sum() / 1e6,
    "peak_GW": dem["Demand_peak_MW"].groupby(dem.index.year).max() / 1e3}).round(1).to_string())
y = d.groupby(d.index.year).sum() / 1e6
print("\nGeneration / energy met:\n" + (y["Total_MWh"] / (dem["Energy_met_MWh"].groupby(dem.index.year).sum() / 1e6))
      .round(3).dropna().to_string())
