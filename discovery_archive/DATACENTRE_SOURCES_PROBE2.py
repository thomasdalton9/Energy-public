"""One-off probe 2: LBNL 2024 report (escholarship), IEA Energy and AI PDF (iea.blob), ERCOT forecast / Batch Zero files, ERCOT large-load
queue pages, EIA AEO data-centre series. Prints figures and sentences only."""
import io, os, re, requests
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
KEY = os.environ.get("EIA_API_KEY", "")
def get(u, **k):
    return requests.get(u, headers=H, timeout=90, **k)
def pdf_text(u):
    from pypdf import PdfReader
    r = get(u); rd = PdfReader(io.BytesIO(r.content))
    return [re.sub(r"\s+", " ", p.extract_text() or "") for p in rd.pages]
def show(pages, pat, ctx=None, n=60, w=420):
    k = 0
    for i, t in enumerate(pages):
        for s in re.split(r"(?<=[.])\s", t):
            if re.search(pat, s, re.I) and (ctx is None or re.search(ctx, s, re.I)):
                print(f"   p{i+1}: {s[:w]}"); k += 1
                if k >= n: return
print("##### LBNL")
try:
    L = pdf_text("https://escholarship.org/content/qt32d6m0d1/qt32d6m0d1.pdf")
    print(len(L), "pages; title:", L[0][:200])
    show(L, r"Texas|ERCOT", n=25)
    show(L, r"load factor|utili[sz]ation|capacity factor", n=20)
    show(L, r"2028", r"TWh|GW", n=20)
    show(L, r"Virginia|state", r"TWh|%", n=15)
    print("--- LBNL p8:", L[7][:2500]); print("--- LBNL p53:", L[52][:2000])
    show(L, r"2023", r"TWh", n=10)
except Exception as e: print("LBNL FAILED", e)
print("##### IEA")
try:
    I = pdf_text("https://iea.blob.core.windows.net/assets/dd7c2387-2f60-4b60-8c5f-6563b6aa1e4c/EnergyandAI.pdf")
    print(len(I), "pages")
    show(I, r"United States", r"TWh|GW", n=45)
    show(I, r"Texas|ERCOT", n=15)
    for pg in (259, 260, 261, 262):
        print(f"--- IEA annex p{pg}:", I[pg-1][:3500])
    show(I, r"Lift-Off|High Efficiency|Headwinds", r"TWh", n=12)
    show(I, r"2035", r"United States", n=8)
    show(I, r"load factor|utili[sz]ation rate|capacity factor", r"data cent", n=10)
except Exception as e: print("IEA FAILED", e)
print("##### ERCOT files")
import pandas as pd
for u in ["https://www.ercot.com/files/docs/2025/04/08/2025-ERCOT-Monthly-Peak-Demand-and-Energy-Forecast.xlsx",
          "https://www.ercot.com/files/docs/2025/04/08/ERCOT-Peak-Demand-Scenarios.xlsx",
          "https://www.ercot.com/files/docs/2026/06/18/Batch-Zero-Load-Information-Form-06172026.xlsx"]:
    try:
        r = get(u); print("\n==", u, r.status_code, len(r.content))
        x = pd.ExcelFile(io.BytesIO(r.content))
        for s in x.sheet_names[:8]:
            d = x.parse(s, header=None)
            print(" sheet", s, d.shape)
            print(d.head(14).iloc[:, :12].to_string()[:1800])
    except Exception as e: print("FAILED", u, e)
print("##### ERCOT pages -> links with large/queue/forecast keywords")
PAGES = ["https://www.ercot.com/gridinfo/planning", "https://www.ercot.com/gridinfo/load/forecast", "https://www.ercot.com/services/rq/large-load-integration",
         "https://www.ercot.com/committees/board", "https://www.ercot.com/news/presentations", "https://www.ercot.com/services/rq", "https://www.ercot.com/gridinfo/resource",
         "https://www.ercot.com/mktrules/issues/NPRR1234", "https://www.ercot.com/about/ltlf", "https://www.ercot.com/gridinfo/planning/longtermloadforecast"]
for p in PAGES:
    try:
        r = get(p); print("\n==", p, r.status_code, len(r.content))
        if r.status_code != 200: continue
        for m in sorted(set(re.findall(r'href="([^"]+)"[^>]*>([^<]{0,120})', r.text))):
            if re.search(r"large|queue|ltlf|long.?term|status.?update|load.?forecast|data.?cent|batch|SB.?6|crypto|officer", m[0] + m[1], re.I): print("   ", m[0], "|", m[1].strip()[:100])
        t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
        for s in re.split(r"(?<=[.])\s", t):
            if re.search(r"\bGW\b|gigawatt", s) and re.search(r"large load|data cent|queue", s, re.I): print("   *", s[:350])
    except Exception as e: print("FAILED", p, e)
print("##### EIA AEO series")
def api(path, **p):
    p["api_key"] = KEY
    r = requests.get("https://api.eia.gov/v2/" + path, params=p, headers=H, timeout=90)
    return r.status_code, (r.json() if r.status_code == 200 else r.text[:200])
for yr in ("2026", "2025"):
    try:
        s, j = api(f"aeo/{yr}/facet/seriesId"); fs = j["response"]["facets"] if s == 200 else []
        print(yr, s, len(fs))
        for f in fs:
            nm = f.get("name") or ""
            if re.search(r"data ?cent|comput|server", nm, re.I) or re.search(r"datac|dtcnt", f["id"], re.I): print("  S", f["id"], "|", nm)
        for f in fs:
            nm = f.get("name") or ""
            if re.search(r"ercot|erct|texas", f["id"] + nm, re.I) and re.search(r"elc|electric", f["id"] + nm, re.I) and re.search(r"comm|demand|sales|retail", f["id"] + nm, re.I): print("  E", f["id"], "|", nm)
        s, j = api(f"aeo/{yr}/facet/scenario"); print(" scenarios", [x["id"] for x in j["response"]["facets"]][:30] if s == 200 else j)
        s, j = api(f"aeo/{yr}/facet/regionId"); print(" regions", [(x["id"], x.get("name")) for x in j["response"]["facets"]][:60] if s == 200 else j)
    except Exception as e: print("EIA FAILED", yr, type(e).__name__, e)
