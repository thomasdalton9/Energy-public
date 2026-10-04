"""Probe 17: full run with the monthly-report (NMOR) backfill; summary as probe 14.
asia/NEPAL_NEA_LDC.py (text-first parsing, process pool) with output streamed, and summarise what it saved:
coverage, annual and monthly totals (compare Ember: ~6-11 TWh/yr generation, imports ~1-2 TWh), gaps, outliers."""
import io
import subprocess
import sys
import time
import pandas as pd
import pdfplumber
import requests

u0 = "https://td.neasite.dryicesolutions.net/uploads/shares/Daily_op_reports/NDOR%202080_06_10.pdf"
body = requests.get(u0, timeout=120, verify=False).content
for what in ("text", "tables"):
    t0 = time.time()
    with pdfplumber.open(io.BytesIO(body)) as pdf:
        p = pdf.pages[0]
        x = p.extract_text() if what == "text" else p.extract_tables()
    print(f"TIMING {what}: {time.time() - t0:.1f} s")
print(x if isinstance(x, str) else x)
with pdfplumber.open(io.BytesIO(body)) as pdf:
    print(pdf.pages[0].extract_text())

OUT = "/tmp/nepal_test.xlsx"
t0 = time.time()
for flags, lim in ((["--no-panel"], 1000), (["--no-pdfs"], 300)):
    t1 = time.time()
    try:
        subprocess.run([sys.executable, "asia/NEPAL_NEA_LDC.py", "--out", OUT] + flags, timeout=lim)
    except subprocess.TimeoutExpired:
        print("!! timed out:", flags)
    print(f"STEP {flags} took {time.time() - t1:.0f} s")
print(f"PULL took {time.time() - t0:.0f} s")
d = pd.read_excel(OUT, sheet_name="Daily", index_col=0, parse_dates=True)
dm = pd.read_excel(OUT, sheet_name="Demand", index_col=0, parse_dates=True)
u = pd.read_excel(OUT, sheet_name="Units")
print("Daily", d.shape, "Demand", dm.shape)
print(d.head(3).to_string())
print(d.tail(3).to_string())
print(dm.head(3).to_string())
print(dm.tail(3).to_string())
print(d["Source"].value_counts())
num = d.drop(columns="Source")
a = (num.resample("YS").sum(min_count=1) / 1000).round(0)
a["days"] = num["Hydro_MWh"].resample("YS").count()
print("\nANNUAL (GWh):\n", a[["Hydro_MWh", "NEA_MWh", "NEA_subsidiary_MWh", "IPP_MWh", "Imports_MWh", "Exports_MWh",
                           "Energy_met_MWh", "Interruption_MWh", "days"]].to_string())
m = (num.resample("MS").sum(min_count=1) / 1000).round(1)
m["days"] = num["Hydro_MWh"].resample("MS").count()
print("\nMONTHLY (GWh):\n", m[["Hydro_MWh", "Imports_MWh", "Exports_MWh", "Energy_met_MWh", "days"]].to_string())
print("\nPeak demand monthly max:", dm["Demand_peak_MW"].resample("MS").max().dropna().astype(int).to_dict())
pdf_d = d[d.Source.isin(["NDOR", "NMOR"])].index
full = pd.date_range(pdf_d.min(), pdf_d.max())
miss = full.difference(pdf_d)
print("by source and year:\n", d.groupby([d.index.year, "Source"]).size().to_string())
print("missing days inside the report span:", len(miss), [x.strftime("%Y-%m-%d") for x in miss])
med = num["Hydro_MWh"].rolling(15, center=True, min_periods=5).median()
print("generation outliers:", num["Hydro_MWh"][(num["Hydro_MWh"] < 0.6 * med) | (num["Hydro_MWh"] > 1.6 * med)].to_dict())
print("UNITS:\n" + "\n".join(u.iloc[:, 0].fillna("").astype(str).tolist()))
