"""One-off probe: layout of EIA state-to-state capacity workbook sheets with Mexico rows. Log only."""
import io
import pandas as pd
import requests

UA = {"User-Agent": "Mozilla/5.0 (energy-data research)"}
r = requests.get("https://www.eia.gov/naturalgas/pipelines/EIA-StatetoStateCapacity_Jan2026.xlsx", headers=UA, timeout=90)
xl = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None)
print("SHEETS", {k: v.shape for k, v in xl.items()})
pd.set_option("display.width", 400); pd.set_option("display.max_columns", 60)
for sh, df in xl.items():
    if "H" not in sh and "Capacity" not in sh:
        continue
    print(f"-- {sh!r} {df.shape}")
    print(df.iloc[:8, :14].to_string(max_colwidth=30))
    print("header row candidates:", [i for i in range(12) if df.iloc[i].astype(str).str.contains(r"^(19|20)\d\d").sum() > 5])
    f = df.copy(); f[0] = f[0].ffill()
    m = f[(f[2].astype(str) == "Mexico") | (f[1].astype(str) == "Mexico")]
    print(m.to_string(max_colwidth=36)[:7000])
