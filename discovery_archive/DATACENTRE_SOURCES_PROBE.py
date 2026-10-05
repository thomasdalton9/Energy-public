"""One-off probe: reachable PUBLISHED sources for US / Texas data-centre electricity (LBNL 2024 report, IEA Energy and AI, EIA AEO via API v2,
ERCOT large-load / long-term load forecast). Prints status, size and sentences with TWh/GW near data centre / large load, plus links."""
import io, os, re, requests
H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
KEY = os.environ.get("EIA_API_KEY", "")
DOCS = [
 "https://eta-publications.lbl.gov/sites/default/files/2024-12/lbnl-2024-united-states-data-center-energy-usage-report.pdf",
 "https://eta-publications.lbl.gov/sites/default/files/lbnl-2024-united-states-data-center-energy-usage-report.pdf",
 "https://escholarship.org/content/qt32d6m0d1/qt32d6m0d1.pdf",
 "https://www.energy.gov/sites/default/files/2024-12/doe-data-center-report.pdf",
 "https://iea.blob.core.windows.net/assets/dd7c2387-2f60-4b60-8c5f-6563b6aa1e4c/EnergyandAI.pdf",
 "https://www.iea.org/reports/energy-and-ai/energy-demand-from-ai",
 "https://www.iea.org/data-and-statistics/data-product/energy-and-ai",
 "https://www.ercot.com/gridinfo/load/forecast",
 "https://www.ercot.com/services/rq/large-load-integration",
 "https://www.ercot.com/gridinfo/load",
 "https://www.ercot.com/news/mediakit/factsheets",
 "https://www.eia.gov/outlooks/aeo/",
 "https://www.eia.gov/outlooks/steo/report/elec_coal_renew.php",
 "https://www.puc.texas.gov/industry/electric/rules/SB6.aspx",
]
KW = re.compile(r"\b(TWh|GW|gigawatt|terawatt|MW)\b", re.I)
CTX = re.compile(r"data ?cent|large load|crypto|AI\b", re.I)
def text(r):
    ct = r.headers.get("content-type", "")
    if "pdf" in ct or r.content[:4] == b"%PDF":
        try:
            from pypdf import PdfReader
            rd = PdfReader(io.BytesIO(r.content))
            return " ".join((p.extract_text() or "") for p in rd.pages[:120])
        except Exception as e:
            return f"[pdf unreadable {e}]"
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", r.text))
for u in DOCS:
    try:
        r = requests.get(u, headers=H, timeout=60)
        print(f"\n== {u}\n   {r.status_code} {len(r.content)} B {r.headers.get('content-type')}", flush=True)
        if r.status_code != 200: continue
        t = re.sub(r"\s+", " ", text(r))
        hits = [s for s in re.split(r"(?<=[.])\s", t) if KW.search(s) and CTX.search(s)]
        for s in hits[:40]: print("   *", s[:380])
        if "html" in r.headers.get("content-type", ""):
            for m in sorted(set(re.findall(r'href="([^"]+\.(?:pdf|xlsx|xls|csv|zip))"', r.text, re.I)))[:60]: print("   link", m)
    except Exception as e:
        print(f"\n== {u}\n   FAILED {type(e).__name__}: {str(e)[:150]}")
print("\n##### EIA API v2 aeo")
def api(path, **p):
    p["api_key"] = KEY
    r = requests.get("https://api.eia.gov/v2/" + path, params=p, headers=H, timeout=60)
    return r.status_code, r.json() if r.status_code == 200 else r.text[:200]
try:
    s, j = api("aeo"); print(s, str(j)[:600])
    resp = j["response"] if s == 200 else {}
    for rt in resp.get("routes", []): print(" route", rt)
    for yr in ("2025", "2023"):
        s, j = api(f"aeo/{yr}"); print(yr, s, str(j)[:500])
        if s != 200: continue
        for rt in j["response"].get("routes", []): print("  ", rt["id"], rt.get("name"))
        s, j = api(f"aeo/{yr}/data"); print(" data meta", s, str(j)[:800])
        s, j = api(f"aeo/{yr}/facet/seriesId"); 
        if s == 200:
            fs = j["response"]["facets"]; print(" nseries", len(fs))
            for f in fs:
                if re.search(r"comput|data ?cent|ELEC.*COMM|commercial.*electric|CNSM.*COMM", f["id"] + f.get("name", ""), re.I): print("  S", f["id"], f.get("name"))
        s, j = api(f"aeo/{yr}/facet/scenario"); print(" scenarios", str(j)[:700])
except Exception as e:
    print("EIA FAILED", type(e).__name__, e)
