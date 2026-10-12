"""Probe 3 (12 Oct 2026): compact structure dump of ET 6.1/6.3, REPD, EirGrid summary workbook, DUKES 5.7/5.16,
Elexon BMU registry and NESO registers (probe 2 showed they exist)."""
import io, re, sys, requests, pandas as pd
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 energy-research"
def dl(u):
    r = S.get(u, timeout=240); print("\n##### ", u, r.status_code, len(r.content), r.headers.get("last-modified")); return r.content
def cell(v, n=22):
    s = str(v).replace("\n", " ")
    return s[:n]
def dump(d, rows=None, cols=None, w=22):
    rows = range(len(d)) if rows is None else rows
    for i in rows:
        r = d.iloc[i]
        vals = [cell(r.iloc[j], w) for j in (cols if cols is not None else range(len(r)))]
        if any(v not in ("nan", "NaT") for v in vals): print(i, "|".join(vals))
base = "https://assets.publishing.service.gov.uk/media/"
# ET 6.1 quarterly
x = pd.ExcelFile(io.BytesIO(dl(base+"6aba724dfceb6fb3a65012d2/ET_6.1_SEP_26.xlsx")))
for sh in ["Quarter", "Northern Ireland - Qtr", "Main Table"]:
    d = x.parse(sh, header=None); print("--- ", sh, d.shape)
    last = list(range(len(d.columns)))[-3:]
    dump(d, cols=[0, 1] + last, w=48)
# ET 6.3 solar
x = pd.ExcelFile(io.BytesIO(dl(base+"6aba729a734e2435f202f19e/ET_6.3_SOLAR_PV_SEP_26.xlsx")))
for sh in ["Table_1_by_Capacity", "Table_4_by_Region", "Notes"]:
    d = x.parse(sh, header=None); print("--- ", sh, d.shape)
    last = list(range(len(d.columns)))[-3:]
    dump(d, cols=[0, 1] + last, w=60)
# REPD
d = pd.read_csv(io.BytesIO(dl(base+"6a6cbdc00c36759b5ccaa305/REPD_Publication_Q2_2026.csv")), encoding="cp1252", low_memory=False)
print(list(d.columns))
op = d[d["Development Status (short)"] == "Operational"]
print("operational rows", len(op))
for c in d.columns:
    if re.search("operat|date|decom|commission", c, re.I):
        print("COL", c, "nonnull op:", op[c].notna().sum(), op[c].dropna().astype(str).head(3).tolist(), "| decomm rows nonnull:", d[d["Development Status (short)"]=="Decommissioned"][c].notna().sum())
print(op.groupby("Technology Type")["Installed Capacity (MWelec)"].apply(lambda s: pd.to_numeric(s, errors="coerce").sum()).round(0).to_string())
print(op["Country"].value_counts().to_string())
print(d[d["Development Status (short)"]=="Decommissioned"].iloc[:, :12].head(8).to_string())
# EirGrid
x = pd.ExcelFile(io.BytesIO(dl("https://cms.eirgrid.ie/sites/default/files/publications/System-and-Renewable-Data-Summary-Report-V30.xlsx")))
d = x.parse("Wind Installed Capacities", header=None); print("--- wind inst", d.shape); dump(d, rows=list(range(0, 8))+list(range(len(d)-8, len(d))), w=14)
d = x.parse("IE Connected Renewables", header=None); print("--- IE conn", d.shape)
print([cell(v, 40) for v in d.iloc[3].tolist()])
hdr = d.iloc[3]; body = d.iloc[4:]
print(body.iloc[:, 1].value_counts().to_string()); print(body.iloc[:, 2].value_counts().to_string())
dump(d, rows=range(len(d)-12, len(d)), cols=list(range(0, 12)), w=18)
for sh in ["IE Connected Non-Renewables", "NI Connected Wind & Solar", "KPI"]:
    d = x.parse(sh, header=None); print("--- ", sh, d.shape); dump(d, rows=range(0, min(len(d), 100 if sh!="KPI" else 40)), cols=list(range(0, 8)), w=26)
# DUKES
for f, n in [("6a6a3622cceb23e8678976bf/DUKES_5.7.xlsx", "5.7"), ("6a6a36a0bf4a7aca8f8976c6/DUKES_5.16.xlsx", "5.16"), ("6a6a3668862aaf18d9c629ed/DUKES_5.12.xlsx", "5.12")]:
    d = pd.read_excel(io.BytesIO(dl(base+f)), sheet_name=n, header=None); print("--- DUKES", n, d.shape)
    dump(d, cols=[0, 1, 2, len(d.columns)-3, len(d.columns)-2, len(d.columns)-1], w=45)
# Elexon BMU
r = S.get("https://data.elexon.co.uk/bmrs/api/v1/reference/bmunits/all?format=json", timeout=120).json()
b = pd.DataFrame(r); b["gc"] = pd.to_numeric(b["generationCapacity"], errors="coerce")
print(b.columns.tolist()); print(b.groupby(b["fuelType"].fillna("NONE"))["gc"].agg(["count", "sum"]).round(0).to_string())
print(b["bmUnitType"].value_counts().to_string())
# NESO
for pkg in ["transmission-entry-capacity-tec-register", "embedded-register"]:
    j = S.get("https://api.neso.energy/api/3/action/package_show?id="+pkg, timeout=60).json()["result"]
    for rr in j["resources"]:
        print("RES", pkg, rr["name"], rr["format"], rr["url"], rr.get("last_modified"))
        try:
            t = pd.read_csv(rr["url"], low_memory=False, encoding="latin1"); print(t.shape, t.columns.tolist()); print(t.head(3).to_string()[:1500])
            for c in t.columns:
                if re.search("type|status|fuel", c, re.I): print(" VC", c, t[c].value_counts().head(25).to_dict())
        except Exception as e: print("  err", e)
