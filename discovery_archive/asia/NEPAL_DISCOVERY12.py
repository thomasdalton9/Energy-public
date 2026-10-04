"""Probe 12: test asia/NEPAL_NEA_LDC.py end to end (full first run, written to a temp path), then summarise:
coverage, annual totals (compare Ember: generation ~6-11 TWh/yr, imports ~1-2 TWh), monthly table, gaps; plus
whether the main site hosts newer NDOR files under its own /uploads/shares path."""
import subprocess
import sys
import pandas as pd
import requests

OUT = "/tmp/nepal_test.xlsx"
r = subprocess.run([sys.executable, "asia/NEPAL_NEA_LDC.py", "--out", OUT], text=True, capture_output=True,
                   timeout=1300)
print(r.stdout[-25000:])
print("STDERR:", r.stderr[-4000:])
d = pd.read_excel(OUT, sheet_name="Daily", index_col=0, parse_dates=True)
dm = pd.read_excel(OUT, sheet_name="Demand", index_col=0, parse_dates=True)
u = pd.read_excel(OUT, sheet_name="Units")
print("SHEETS ok; Daily", d.shape, "Demand", dm.shape)
print(d.head(3).to_string())
print(d.tail(3).to_string())
print(dm.head(3).to_string())
print(d["Source"].value_counts())
num = d.drop(columns="Source")
print("\nANNUAL (GWh) + days:")
a = (num.resample("YS").sum(min_count=1) / 1000).round(0)
a["days"] = num["Hydro_MWh"].resample("YS").count()
print(a[["Hydro_MWh", "NEA_MWh", "NEA_subsidiary_MWh", "IPP_MWh", "Imports_MWh", "Exports_MWh", "Energy_met_MWh",
         "Interruption_MWh", "days"]].to_string())
print("\nMONTHLY (GWh):")
m = (num.resample("MS").sum(min_count=1) / 1000).round(1)
m["days"] = num["Hydro_MWh"].resample("MS").count()
print(m[["Hydro_MWh", "Imports_MWh", "Exports_MWh", "Energy_met_MWh", "days"]].to_string())
print("\nPeak demand monthly max:", dm["Demand_peak_MW"].resample("MS").max().dropna().astype(int).to_dict())
full = pd.date_range(d[d.Source == "NDOR"].index.min(), d[d.Source == "NDOR"].index.max())
miss = full.difference(d.index)
print("missing days inside the report span:", len(miss), [x.strftime("%Y-%m-%d") for x in miss][:40])
# sanity outliers
r7 = num["Hydro_MWh"].rolling(15, center=True, min_periods=5).median()
odd = num["Hydro_MWh"][(num["Hydro_MWh"] < 0.6 * r7) | (num["Hydro_MWh"] > 1.6 * r7)]
print("generation outliers vs 15-day median:", odd.to_dict())
print("UNITS:", "\n".join(u.iloc[:, 0].astype(str).tolist())[:3000])
