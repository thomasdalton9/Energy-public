"""One-off probe 3: full structure of EirGrid's System-and-Renewable-Data-Summary-Report (Fuel Mix & CO2 sheet; row labels of
System Data Summary) and the column list of the quarter-hourly System Data file, plus any other report links on the page."""
import io
import re
import sys

import pandas as pd
import requests

H = {"User-Agent": "Mozilla/5.0", "Accept": "*/*"}
pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 60)
page = requests.get("https://www.eirgrid.ie/grid/system-and-renewable-data-reports", headers=H, timeout=90).text
print("all hrefs with 'data' or xlsx:", sorted(set(re.findall(r'href="([^"]*(?:ystem|enewable|ata)[^"]*)"', page)))[:60], flush=True)
base = "https://cms.eirgrid.ie/sites/default/files/publications/"
r = requests.get(base + "System-and-Renewable-Data-Summary-Report-V30.xlsx", headers=H, timeout=180)
x = pd.ExcelFile(io.BytesIO(r.content))
fm = x.parse("Fuel Mix & CO2", header=None)
print("\n=== Fuel Mix & CO2", fm.shape)
print(fm.to_string(max_colwidth=30)[:6000])
sd = x.parse("System Data Summary", header=None)
print("\n=== System Data Summary", sd.shape)
lab = sd.iloc[:, :3].ffill().dropna(how="all")
print(lab.to_string()[:5000])
print("header rows (first 3 cols of dates):", sd.iloc[0:2, 3:8].values.tolist(), "last cols:", sd.iloc[0:2, -4:].values.tolist())
r2 = requests.get(base + "System-Data-Qtr-Hourly-2026-V9.xlsx", headers=H, timeout=180)
q = pd.read_excel(io.BytesIO(r2.content), nrows=3)
print("\n=== Qtr-hourly columns:", list(q.columns))
for yr in (2025, 2024, 2023, 2022, 2021):
    for v in range(1, 16):
        u = base + f"System-Data-Qtr-Hourly-{yr}.xlsx" if v == 1 else base + f"System-Data-Qtr-Hourly-{yr}-V{v}.xlsx"
        try:
            h = requests.head(u, headers=H, timeout=30)
            if h.status_code == 200:
                print("EXISTS", u)
                break
        except Exception:  # noqa: BLE001
            pass
sys.exit(0)
