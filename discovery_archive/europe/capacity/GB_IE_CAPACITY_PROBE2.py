"""Probe 2 (12 Oct 2026): structure of ET 6.1 (quarterly renewable capacity), ET 6.3 (monthly solar PV), REPD,
EirGrid System and Renewable Data Summary Report, DUKES 5.7/5.12/5.16."""
import io, re, sys, requests, pandas as pd
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 energy-research"
def dl(u):
    r = S.get(u, timeout=180); print("\n##### ", u, r.status_code, len(r.content), r.headers.get("last-modified")); return r.content
def xl(u, rows=14, maxsheets=40):
    b = dl(u); x = pd.ExcelFile(io.BytesIO(b))
    for sh in x.sheet_names[:maxsheets]:
        d = x.parse(sh, header=None)
        print(f"--- sheet {sh!r} shape {d.shape}")
        with pd.option_context("display.width", 250, "display.max_columns", 16, "display.max_colwidth", 28):
            print(d.head(rows).to_string())
base = "https://assets.publishing.service.gov.uk/media/"
xl(base+"6aba724dfceb6fb3a65012d2/ET_6.1_SEP_26.xlsx", 12)
xl(base+"6aba729a734e2435f202f19e/ET_6.3_SOLAR_PV_SEP_26.xlsx", 12)
# REPD
b = dl(base+"6a6cbdc00c36759b5ccaa305/REPD_Publication_Q2_2026.csv")
d = pd.read_csv(io.BytesIO(b), encoding="cp1252", low_memory=False)
print(d.shape); print(list(d.columns))
for c in d.columns:
    if re.search("Technology|Status|Date|Capacity|Country|Storage|Development", c, re.I):
        print("COL", c, d[c].dropna().astype(str).head(4).tolist(), d[c].nunique())
print(d["Technology Type"].value_counts().to_string())
print(d["Development Status (short)"].value_counts().to_string())
# EirGrid
xl("https://cms.eirgrid.ie/sites/default/files/publications/System-and-Renewable-Data-Summary-Report-V30.xlsx", 10)
# DUKES
xl(base+"6a6a3622cceb23e8678976bf/DUKES_5.7.xlsx", 40)
xl(base+"6a6a3668862aaf18d9c629ed/DUKES_5.12.xlsx", 40)
xl(base+"6a6a36a0bf4a7aca8f8976c6/DUKES_5.16.xlsx", 30)
