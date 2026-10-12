"""Probe 5 (12 Oct 2026): EirGrid 'IE Connected Renewables' right-hand blocks (distribution-connected, ESBN), KPI sheet
capacity rows, and the ESBN generator-statistics page."""
import io, re, requests, pandas as pd
S = requests.Session(); S.headers["User-Agent"] = "Mozilla/5.0 energy-research"
def cell(v, n=22): return str(v).replace("\n", " ")[:n]
u = "https://cms.eirgrid.ie/sites/default/files/publications/System-and-Renewable-Data-Summary-Report-V30.xlsx"
x = pd.ExcelFile(io.BytesIO(S.get(u, timeout=200).content))
d = x.parse("IE Connected Renewables", header=None)
print(d.shape)
# find columns that start a block: non-empty in header row 3
for j in range(d.shape[1]):
    col = d.iloc[:, j]
    nn = col.notna().sum()
    print(j, nn, "|", cell(d.iloc[0, j], 50), "|", cell(d.iloc[1, j], 40), "|", cell(d.iloc[2, j], 40), "|", cell(d.iloc[3, j], 28), "|", cell(d.iloc[4, j], 28), "|", cell(d.iloc[5, j], 28))
k = x.parse("KPI", header=None)
for i in range(len(k)):
    t = " ".join(cell(v, 60) for v in k.iloc[i].tolist() if str(v) != "nan")
    if re.search(r"capacity|solar|installed|battery|storage|connected", t, re.I): print("KPI", i, t[:200])
for url in ["https://www.esbnetworks.ie/services/get-connected/renewable-connection/generator-statistics"]:
    try:
        r = S.get(url, timeout=60); print("ESBN", r.status_code, len(r.content), r.headers.get("content-type"))
        for m in sorted(set(re.findall(r'href="([^"]+\.(?:xlsx|xls|csv|pdf))"', r.text, re.I))): print("   LINK", m)
        print(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))[:1500])
    except Exception as e: print("ESBN ERR", e)
