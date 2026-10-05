"""Probe: EIA-923 download links and whether Page 1 carries a Balancing Authority Code."""
import re, io, zipfile, requests, pandas as pd
H = {"User-Agent": "Mozilla/5.0"}
r = requests.get("https://www.eia.gov/electricity/data/eia923/", headers=H, timeout=60)
print(r.status_code, len(r.text))
links = sorted(set(re.findall(r'href="([^"]*(?:xls|zip)[^"]*)"', r.text)))
for l in links[-4:]: print(l)
for u in ["https://www.eia.gov/electricity/data/eia923/archive/xls/f923_2022.zip","https://www.eia.gov/electricity/data/eia923/xls/f923_2026.zip"]:
    z = requests.get(u, headers=H, timeout=120); print(u, z.status_code, len(z.content))
    if z.content[:2] != b"PK":
        print(z.content[:200]); continue
    zf = zipfile.ZipFile(io.BytesIO(z.content)); print(zf.namelist())
    n = [x for x in zf.namelist() if "Schedules_2_3_4_5" in x][0]
    x = pd.ExcelFile(zf.open(n)); print(n, x.sheet_names)
    raw = pd.read_excel(x, sheet_name=0, header=None, nrows=8); print(raw.iloc[:7,:12].to_string())
    d = pd.read_excel(x, sheet_name=0, header=5); print(list(d.columns)); print(d.tail(2).T.to_string()[:3000])
    for c in d.columns:
        if "alanc" in str(c): print(d[c].value_counts().head(8))
