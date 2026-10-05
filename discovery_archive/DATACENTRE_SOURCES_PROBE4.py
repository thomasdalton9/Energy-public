"""One-off probe 4: ERCOT CDR load-forecast tab (data centres by year), ERCOT hearing slide text (queue GW), LBNL state section (Texas)."""
import io, re, requests, pandas as pd
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
from pypdf import PdfReader
def pages(u):
    r = requests.get(u, headers=H, timeout=120); print("   status", r.status_code, len(r.content))
    return [re.sub(r"\s+", " ", p.extract_text() or "") for p in PdfReader(io.BytesIO(r.content)).pages]
print("##### ERCOT hearing slides (full text of pages 1-4)")
P = pages("https://www.ercot.com/files/docs/2026/04/09/ERCOTLargeLoadUpdate-April9HouseStateAffairsHearing.pdf")
for i in range(0, 5): print(f"--- p{i+1}:", P[i][:1500])
print("##### ERCOT Aug 2026 data-centre panel pages 1-5")
P = pages("https://www.ercot.com/files/docs/2026/08/19/ERCOTPanel1DataCenters.pdf")
for i in range(0, min(6, len(P))): print(f"--- p{i+1}:", P[i][:1500])
print("##### LBNL state pages")
L = pages("https://escholarship.org/content/qt32d6m0d1/qt32d6m0d1.pdf")
for i in (37, 38, 39):
    print(f"--- LBNL p{i+1}:", L[i][:2200])
for i, t in enumerate(L):
    if re.search(r"Texas", t): print(f"   LBNL p{i+1} mentions Texas:", [s[:300] for s in re.split(r"(?<=[.])\s", t) if "Texas" in s][:3])
print("##### CDR Dec 2025 LoadResourceScenarios")
r = requests.get("https://www.ercot.com/files/docs/2025/12/19/CapacityDemandandReservesReport_December2025.xlsx", headers=H, timeout=120)
print(r.status_code, len(r.content))
x = pd.ExcelFile(io.BytesIO(r.content)); print(x.sheet_names)
for s in x.sheet_names:
    if re.search(r"LoadResource", s):
        d = x.parse(s, header=None)
        print("\n sheet", s, d.shape)
        for i, row in d.iterrows():
            vals = [str(v)[:26] for v in row.tolist() if str(v) != "nan"]
            if vals: print("  ", i, " | ".join(vals)[:260])
            if i > 60: break
