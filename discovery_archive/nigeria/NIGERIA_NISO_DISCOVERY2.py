"""Nigeria discovery round 2: nsong.org Dashboard portal, NERC industry statistics / wp-json."""
import re, json, urllib3, requests
urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}


def get(url, n=1500, **kw):
    try:
        r = requests.get(url, headers=H, timeout=(8, 30), verify=False, **kw)
        print(f"\nGET {url} -> {r.status_code} {len(r.content)}B {r.headers.get('content-type')} final={r.url}")
        ct = r.headers.get("content-type") or ""
        if n and ("text" in ct or "json" in ct):
            print(r.text[:n])
        return r
    except Exception as e:
        print(f"\nGET {url} -> ERR {type(e).__name__}: {str(e)[:120]}")


def clean(t):
    return re.sub(r"\s+", " ", re.sub(r"<script.*?</script>|<style.*?</style>", "", t, flags=re.S))


def alllinks(r):
    for m in re.finditer(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', r.text, re.I | re.S):
        print("   LINK", m.group(1), "|", re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()[:70])


for u in ["https://nsong.org/", "https://nsong.org/Dashboard.aspx"]:
    r = get(u, n=0)
    if r is not None and r.status_code == 200:
        print(clean(r.text)[:2500])
        alllinks(r)
        print("ids:", re.findall(r'id="([^"]+)"', r.text)[:40])
for p in ["Pages/Home.aspx", "Reports.aspx", "DailyReport.aspx", "api", "swagger", "robots.txt", "sitemap.xml"]:
    get("https://nsong.org/" + p, n=300)

print("\n=========== NERC ===========")
r = get("https://nerc.gov.ng/industry-statistics/", n=0)
if r is not None and r.status_code == 200:
    print(clean(r.text)[:2500])
    for m in re.finditer(r'href=["\']([^"\']+\.(?:xlsx?|csv|pdf|json)[^"\']*)', r.text, re.I):
        print("   FILE", m.group(1))
r = get("https://nerc.gov.ng/resource-category/nerc-reports", n=0)
if r is not None and r.status_code == 200:
    alllinks(r)
for q in ["generation", "daily", "quarterly", "monthly+report", "operational"]:
    get(f"https://nerc.gov.ng/wp-json/wp/v2/media?per_page=30&search={q}&_fields=source_url,title,date", n=2500)

print("\n=========== Electricity Hub ===========")
r = get("https://nigeriaelectricityhub.com/data-reports/", n=0)
if r is not None and r.status_code == 200:
    alllinks(r)
for q in ["daily+generation", "NISO"]:
    get(f"https://nigeriaelectricityhub.com/wp-json/wp/v2/search?search={q}&per_page=15", n=1500)
