"""Probe: EIA-923 download links and whether Page 1 carries a Balancing Authority Code."""
import re, io, zipfile, requests, pandas as pd
H = {"User-Agent": "Mozilla/5.0"}
r = requests.get("https://www.eia.gov/electricity/data/eia923/", headers=H, timeout=60)
print(r.status_code, len(r.text))
links = sorted(set(re.findall(r'href="([^"]*(?:xls|zip)[^"]*)"', r.text)))
for l in links[:60]: print(l)
for u in ["https://www.eia.gov/electricity/data/eia923/xls/f923_2024.zip"]:
    z = requests.get(u, headers=H, timeout=120); print(u, z.status_code, len(z.content))
    zf = zipfile.ZipFile(io.BytesIO(z.content)); print(zf.namelist())
    n = [x for x in zf.namelist() if "Schedules_2_3_4_5" in x][0]
    raw = pd.read_excel(zf.open(n), sheet_name=0, header=None, nrows=8)
    print(raw.to_string()[:3000])
    x = pd.ExcelFile(zf.open(n)); print(x.sheet_names)
    d = pd.read_excel(x, sheet_name=0, header=5); print(list(d.columns))
    for c in d.columns:
        if "alanc" in str(c): print(d[c].value_counts().head(8))
