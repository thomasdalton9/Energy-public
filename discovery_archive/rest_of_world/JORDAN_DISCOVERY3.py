"""Jordan discovery 3: NEPCO Stats/MaxLoadArchive/AnnualReports pages, CKAN POST, DOS pxweb, MEMR lists, NEPCO PDF text."""
import re, io, json, sys, requests
requests.packages.urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}
def show(label, r, n=1500):
    print(f"\n=== {label}\n  {r.status_code} {len(r.content)}B {r.headers.get('content-type')} {r.url}", flush=True)
def page(url, n=2500):
    try:
        r = requests.get(url, headers=H, timeout=(10, 40), verify=False)
    except Exception as e:
        print("\n===", url, "ERR", str(e)[:120]); return None
    show(url, r)
    t = r.text
    body = re.sub(r'<script.*?</script>|<style.*?</style>', ' ', t, flags=re.S)
    txt = re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', body))
    print("  TEXT:", txt[:n])
    for l in sorted(set(re.findall(r'(?:href|src|data-url|action)=["\']([^"\']+)["\']', t))):
        if re.search(r'\.(xlsx?|csv|pdf|json|asmx|ashx)|api|handler|chart|stat|load|generat|download|resource', l, re.I) and not re.search(r'\.(png|jpg|css|js)\b', l):
            print("   L", l[:200])
    for l in re.findall(r'(?:url|ajax)\s*[:=]\s*["\']([^"\']+)["\']', t)[:20]: print("   JS", l[:200])
    return r
for u in ["https://www.nepco.com.jo/en/Stats.aspx", "https://www.nepco.com.jo/en/MaxLoadArchive.aspx", "https://www.nepco.com.jo/en/AnnualReports.aspx",
          "https://www.nepco.com.jo/Stats.aspx", "https://www.memr.gov.jo/En/List/Open_Data", "https://www.memr.gov.jo/En/List/Studies_and_Statistics",
          "https://dosweb.dos.gov.jo/industry/industry-and-energy/", "https://jorinfo.dos.gov.jo/Databank/pxweb/en/", "https://opendata.gov.jo/en/dataset/?tags=ELECTRICITY",
          "https://opendata.gov.jo/en/api-guide/"]:
    page(u)
for q in ["electricity", "nepco", "natural gas", "energy"]:
    try:
        r = requests.post("https://opendata.gov.jo/api/3/action/package_search", json={"q": q, "rows": 30}, headers=H, timeout=30, verify=False)
        show("CKAN POST " + q, r)
        j = r.json()
        print("  count", j.get("result", {}).get("count"))
        for p in j.get("result", {}).get("results", []):
            print("  PKG", p["name"], "|", p.get("organization", {}).get("name"), "|", [(x["format"], x["url"][-60:]) for x in p["resources"]][:4])
    except Exception as e: print("CKAN ERR", q, str(e)[:150])
try:
    r = requests.post("https://opendata.gov.jo/api/3/action/package_show", json={"id": "nepco-2170-2023"}, headers=H, timeout=30, verify=False)
    show("package_show", r); j = r.json()
    for x in j["result"]["resources"]: print("  RES", x["name"], x["format"], x["url"])
except Exception as e: print("ERR", e)
# NEPCO PDF text
try:
    import subprocess
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "pdfplumber"], check=False)
    import pdfplumber
    r = requests.get("https://www.nepco.com.jo/store/DOCS/web/2023_en.pdf", headers=H, timeout=120, verify=False)
    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
        print("\n=== PDF pages", len(pdf.pages))
        for i, p in enumerate(pdf.pages):
            t = p.extract_text() or ""
            if re.search(r'natural gas|gas|import|export|generation by|fuel', t, re.I) and re.search(r'GWh|ktoe|MMBTU|BCM|MMSCF', t):
                print(f"--- page {i+1}\n", t[:1200])
            if i > 120: break
except Exception as e: print("PDF ERR", e)
