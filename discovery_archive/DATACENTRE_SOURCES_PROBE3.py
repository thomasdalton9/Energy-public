"""One-off probe 3: IEA Energy and AI annex tables (US data centres, load factor), ERCOT large-load hearing PDFs (queue GW, data-centre share),
EIA AEO2026 'Data Center Servers' series (US and Texas Reliability Entity)."""
import io, os, re, requests
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
KEY = os.environ.get("EIA_API_KEY", "")
from pypdf import PdfReader
def pages(u):
    r = requests.get(u, headers=H, timeout=120); print("   status", r.status_code, len(r.content))
    return [re.sub(r"\s+", " ", p.extract_text() or "") for p in PdfReader(io.BytesIO(r.content)).pages]
print("##### IEA annex")
I = pages("https://iea.blob.core.windows.net/assets/dd7c2387-2f60-4b60-8c5f-6563b6aa1e4c/EnergyandAI.pdf")
for pg in (258, 259, 260, 261, 262, 263):
    print(f"--- IEA p{pg}:", I[pg - 1][:3800])
for pat in (r"Lift-Off Case|Headwinds Case|High Efficiency Case",):
    k = 0
    for i, t in enumerate(I):
        for s in re.split(r"(?<=[.])\s", t):
            if re.search(pat, s) and re.search(r"TWh", s) and k < 14: print(f"   p{i+1}: {s[:380]}"); k += 1
print("##### ERCOT PDFs")
for u in ["https://www.ercot.com/files/docs/2026/04/01/ERCOT_LargeLoad_Update_April2026_B-C_-Hearing.pdf",
          "https://www.ercot.com/files/docs/2026/04/09/ERCOTLargeLoadUpdate-April9HouseStateAffairsHearing.pdf",
          "https://www.ercot.com/files/docs/2026/08/19/ERCOTPanel1DataCenters.pdf",
          "https://www.ercot.com/files/docs/2026/08/20/Batch-Zero-Verification-and-Audit-PUCT-Presentation.pdf"]:
    print("\n==", u)
    try:
        P = pages(u)
        for i, t in enumerate(P[:40]):
            for s in re.split(r"(?<=[.])\s|(?<=\d)\s(?=[A-Z])", t):
                if re.search(r"\bGW\b|\bMW\b|gigawatt", s) and re.search(r"large load|data ?cent|queue|request|forecast|crypto|officer|energi[sz]|approved|interconnect", s, re.I):
                    print(f"   p{i+1}: {s[:330]}")
    except Exception as e: print("   FAILED", e)
print("\n##### EIA AEO2026 data-center series")
def api(path, **p):
    p["api_key"] = KEY
    r = requests.get("https://api.eia.gov/v2/" + path, params=p, headers=H, timeout=90)
    return r.status_code, (r.json() if r.status_code == 200 else r.text[:300])
s, j = api("aeo/2026/facet/seriesId"); fs = j["response"]["facets"] if s == 200 else []
ids = [f["id"] for f in fs if f.get("id") and re.search(r"datactr", f["id"])]
print(len(ids), "datactr series"); 
for f in fs:
    if f.get("id") in ids: print("  ", f["id"], "|", f.get("name"))
ids += [f["id"] for f in fs if f.get("id") and re.match(r"cnsm_NA_comm_NA_elc_NA_(trel|ercot|usa)_", f["id"])][:6]
ids += [f["id"] for f in fs if f.get("id") and re.search(r"^cnsm_NA_(comm|resd|idal)_NA_elc_NA_(trel|NA|usa)_blnkwh", f["id"])][:8]
for sid in ids[:14]:
    for scen in ("ref2026", "lowogs"):
        pass
    s, j = api("aeo/2026/data", **{"frequency": "annual", "data[0]": "value", "facets[seriesId][]": sid, "length": "400"})
    if s != 200: print(sid, s, j); continue
    rows = j["response"]["data"]
    print("\n", sid, rows[0].get("seriesName"), rows[0].get("unit"), len(rows), "rows; regions", sorted({r["regionId"] for r in rows}))
    by = {}
    for r in rows: by.setdefault(r["scenario"], {})[r["period"]] = r["value"]
    for sc, d in by.items():
        print("   ", sc, {y: d.get(y) for y in ("2023", "2024", "2025", "2028", "2030", "2033", "2035", "2040", "2050") if y in d})
