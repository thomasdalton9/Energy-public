"""Discovery 2: follow links on NUPRC, NNPC, NBS, JODI, NMDPRA, NLNG pages."""
import re, sys, io
import requests
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
def get(u, **k):
    try:
        return requests.get(u, headers=UA, timeout=40, **k)
    except Exception as e:
        print("   ERR", u, type(e).__name__, str(e)[:80]); return None
def hrefs(t):
    return sorted(set(re.findall(r'href=["\']([^"\']+)', t)))
def text(t, n=1500):
    t = re.sub(r"<(script|style).*?</\1>", " ", t, flags=re.S|re.I)
    t = re.sub(r"<[^>]+>", " ", t)
    return re.sub(r"\s+", " ", t)[:n]
def show(u, pat=None, n=60, txt=0):
    print(f"\n== {u}")
    r = get(u)
    if r is None: return None
    print("  ", r.status_code, r.headers.get("content-type"), len(r.content))
    if "html" in r.headers.get("content-type", ""):
        hs = hrefs(r.text)
        if pat: hs = [h for h in hs if re.search(pat, h, re.I)]
        for h in hs[:n]: print("   href:", h)
        if txt: print("   text:", text(r.text, txt))
    sys.stdout.flush()
    return r
B = "https://www.nuprc.gov.ng"
show(B + "/development-production", n=80, txt=800)
show(B + "/reports", n=100, txt=800)
for p in ["/gas", "/gas-flare-data", "/data", "/statistics", "/resources", "/publications", "/annual-statistical-bulletin", "/monthly-production", "/production-data", "/downloads", "/gas-development", "/domestic-gas"]:
    show(B + p, pat=r"\.(pdf|xlsx?|csv)|prod|gas|bulletin|stat", n=20)
show("https://www.nnpcgroup.com/insights/nnpc-limited-monthly-report-summary-august-2026", pat=r"pdf|xls|report|download|insights", n=60, txt=3000)
r = show("https://www.nnpcgroup.com/insights", pat=r"report|monthly", n=80)
r = show("https://www.nnpcgroup.com/reports", pat=r"report|monthly|pdf", n=80)
r = show("https://www.nnpcgroup.com/insights/reports", pat=r"report|monthly|pdf", n=80)
for u in ["https://www.nlng.com/", "https://nlng.com/", "https://www.nlng.com/Pages/default.aspx", "https://www.nlng.com/investors", "https://www.nigerialng.com.ng/", "https://nigerialng.com/"]:
    show(u, n=30, txt=300)
show("https://nmdpra.gov.ng/", n=80, txt=1200)
for p in ["gas", "domestic-gas-supply-obligation", "dsno", "downloads", "publications", "reports", "gas-data"]:
    show("https://nmdpra.gov.ng/" + p, pat=r"pdf|xls|gas|dsno|report", n=20)
show("https://www.nigerianstat.gov.ng/elibrary", pat=r"oil|gas|petrol|crude|energy|download|read", n=120)
show("https://nigerianstat.gov.ng/elibrary?search=oil", n=60)
show("https://www.jodidata.org/gas/database/data-downloads.aspx", pat=r"zip|csv|xls|download|rest|api", n=60)
show("https://www.jodidata.org/gas/database/data-downloads.aspx", n=0, txt=2500)
