"""Nigeria discovery round 3 (compact output): nsong.org portal, NERC statistics categories, other NISO hosts."""
import re, socket, urllib3, requests
urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=(8, 30), verify=False, **kw)
        print(f"GET {url} -> {r.status_code} {len(r.content)}B {(r.headers.get('content-type') or '')[:30]} final={r.url}", flush=True)
        return r
    except Exception as e:
        print(f"GET {url} -> ERR {type(e).__name__}: {str(e)[:100]}", flush=True)


def clean(t):
    return re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>|<[^>]+>", " ", t, flags=re.S))


print("## DNS")
for h in ["niso.ng", "www.niso.ng", "neso.ng", "nsong.org", "www.nsong.org", "tcn.org.ng", "www.tcn.org.ng", "nigerianiso.com",
          "niso.gov.ng", "nisong.org", "gencos.org.ng", "nbet.com.ng", "www.nbet.com.ng", "data.nerc.gov.ng", "opendata.nerc.gov.ng"]:
    try:
        print(h, socket.gethostbyname(h))
    except Exception as e:
        print(h, "no DNS")

print("\n## nsong.org")
for u in ["https://nsong.org/", "https://nsong.org/Dashboard.aspx"]:
    r = get(u)
    if r is not None and r.status_code == 200:
        print("TEXT:", clean(r.text)[:1200])
        print("LINKS:", re.findall(r'href=["\']([^"\']+)["\']', r.text)[:30])
        print("SRC/ACTION:", re.findall(r'(?:src|action)=["\']([^"\']+)["\']', r.text)[:30])
for p in ["Pages/Default.aspx", "Pages/DailyReports/", "DailyReports", "Reports", "Pages/Reports.aspx", "Login.aspx", "Pages/Login.aspx",
          "Default.aspx", "Pages/Generation.aspx", "api/", "Dashboard/Reports", "Home/Index"]:
    get("https://nsong.org/" + p)

print("\n## NERC listings")
for cat in ["statistics-and-performance-data", "operational-performance-factsheet", "commercial-performance-factsheet"]:
    r = get(f"https://nerc.gov.ng/resource-category/{cat}/")
    if r is not None and r.status_code == 200:
        fl = re.findall(r'href=["\']([^"\']+\.(?:xlsx?|csv|pdf)[^"\']*)', r.text, re.I)
        print(len(fl), "files; first 12:", sorted(set(fl))[:12])
        print("pages:", sorted(set(re.findall(r'/resource-category/' + cat + r'/page/(\d+)', r.text))))
r = get("https://nerc.gov.ng/industry-statistics/")
if r is not None:
    print(clean(r.text)[:800])
for u in ["https://nerc.gov.ng/wp-json/wp/v2/types", "https://nerc.gov.ng/wp-json/wp/v2/resources?per_page=5&_fields=link,title,date"]:
    r = get(u)
    if r is not None and r.status_code == 200:
        print(r.text[:800])
