"""Nigeria raw grid data discovery (NISO / NERC). Prints what is reachable, links and file formats.
Usage: python3 NIGERIA_NISO_DISCOVERY.py
"""
import re
import sys
import urllib3
import requests

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
KEY = re.compile(r"(daily|report|generation|operational|energy|power|data|download|archive|\.xlsx?|\.csv|\.pdf|api|dashboard|market)", re.I)


def get(url, **kw):
    try:
        r = requests.get(url, headers=H, timeout=(10, 45), verify=False, **kw)
        print(f"GET {url} -> {r.status_code} {len(r.content)}B {r.headers.get('content-type')} final={r.url}")
        return r
    except Exception as e:
        print(f"GET {url} -> ERR {type(e).__name__}: {str(e)[:150]}")
        return None


def links(r):
    out = []
    for m in re.finditer(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', r.text, re.I | re.S):
        t = re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()[:80]
        out.append((m.group(1), t))
    return out


seen = set()
for base in ["https://niso.ng/", "https://www.niso.ng/", "https://nsong.org/", "https://www.nsong.org/",
             "https://nerc.gov.ng/", "https://www.nerc.gov.ng/", "https://www.gencosnigeria.com/",
             "https://nigeriaelectricityhub.com/", "https://www.nigeriaelectricityhub.com/",
             "https://niso.ng/services/", "https://niso.ng/services/daily-reports/",
             "https://niso.ng/services/daily-operational-reports/", "https://nsong.org/Pages/DailyReports.aspx"]:
    r = get(base)
    if r is None or r.status_code != 200 or "html" not in (r.headers.get("content-type") or ""):
        continue
    ls = links(r)
    print(f"  {len(ls)} links; wp-json? {'wp-json' in r.text}")
    for h, t in ls:
        if h in seen:
            continue
        seen.add(h)
        if KEY.search(h) or KEY.search(t):
            print(f"   - {h}  [{t}]")
print("\n--- WordPress REST probe ---")
for b in ["https://niso.ng", "https://nsong.org"]:
    r = get(b + "/wp-json/wp/v2/media?per_page=20&search=daily")
    if r is not None and r.status_code == 200:
        print(r.text[:1500])
    r = get(b + "/wp-json/wp/v2/pages?per_page=100&_fields=link,title")
    if r is not None and r.status_code == 200:
        print(r.text[:3000])
