"""
Probe 7 (manual workflow china_sources_probe7.yml): open the reachable pages found by CHINA_SOURCES_PROBE.py and print
their text, to see whether the monthly tables can be parsed:
  - NDRC 'national natural gas operation bulletin' (全国天然气运行快报): list pages + latest bulletins
  - NEA monthly power-industry statistics (installed capacity) and electricity consumption releases
  - GACC English statistics pages (imports of crude oil / natural gas / coal)
"""
import re
import signal
import time

import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"}
S = requests.Session()
S.headers.update(H)


class Hard(Exception):
    pass


def _alarm(*_):
    raise Hard("hard 45 s limit")


signal.signal(signal.SIGALRM, _alarm)


def get(url):
    signal.alarm(45)   # requests' timeouts do not bound a slowly trickling body
    try:
        r = S.get(url, timeout=(10, 30))
        signal.alarm(0)
    except (requests.RequestException, Hard) as e:
        signal.alarm(0)
        print(f"  ERROR {url}: {type(e).__name__}: {str(e)[:120]}", flush=True)
        return None
    if r.encoding in (None, "ISO-8859-1"):
        r.encoding = r.apparent_encoding
    print(f"GET {url} -> {r.status_code} {len(r.content)} bytes", flush=True)
    time.sleep(1)
    return r


def text(html):
    html = re.sub(r"(?s)<(script|style).*?</\1>", " ", html)
    html = re.sub(r"</(p|div|tr|li|h\d)>", "\n", html)
    html = re.sub(r"</t[dh]>", " | ", html)
    t = re.sub(r"<[^>]+>", "", html)
    t = re.sub(r"&nbsp;|　", " ", t)
    return re.sub(r"\n\s*\n+", "\n", re.sub(r"[ \t]+", " ", t))


def links(html, pat):
    return [(h, t.strip()) for h, t in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>([^<]{3,90})</a>', html) if re.search(pat, t)]


def show(url, n=3500, start=None):
    r = get(url)
    if r is None or r.status_code != 200:
        return None
    t = text(r.text)
    if start and start in t:
        t = t[t.index(start):]
    print(t[:n], "\n-----", flush=True)
    return r



CE = "http://english.customs.gov.cn"
print("=== GACC table (14), Aug 2026: header and energy rows, cell by cell ===")
r = get(CE + "/Statics/4b78370a-cd59-4e25-b483-5ef203143664.html")
if r is not None and r.status_code == 200:
    rows = re.findall(r"(?s)<tr[^>]*>(.*?)</tr>", r.text)
    print("rows:", len(rows))
    shown = 0
    for i, row in enumerate(rows):
        cells = [re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", c)).strip() for c in re.findall(r"(?s)<t[dh][^>]*>(.*?)</t[dh]>", row)]
        key = cells[0] if cells else ""
        if i < 8 or re.match(r"(?i)(coal|crude|natural gas|refined|liquefied|petroleum|naphtha)", key):
            print(f"  row {i}:", cells[:14])
            shown += 1
        if shown > 24:
            break
print("=== GACC: list of all monthly bulletin pages: (14) titles, paging until empty ===")
months = {}
last_pg = 0
for pg in range(1, 160):
    rr = get(f"{CE}/Statistics/Statistics?ColumnId=2&page={pg}")
    if rr is None or rr.status_code != 200:
        break
    items = re.findall(r'<a[^>]+href="([^"]*Statics[^"]*)"[^>]*>([^<]+)</a>', rr.text)
    if not items:
        print("  empty at page", pg)
        break
    last_pg = pg
    for h, t in items:
        if re.search(r"Major Import Commodities", t):
            months[t.strip()] = h
print("last page with items:", last_pg, "; (14) titles:", len(months))
for t in months:
    print("   ", t)
