"""
Probe 6 (manual workflow china_sources_probe6.yml): open the reachable pages found by CHINA_SOURCES_PROBE.py and print
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
print("=== GACC: major imports table (6), raw text around energy rows ===")
t6 = get(CE + "/Statics/1c29d265-76bf-4a44-9867-224c6ba7ec52.html")
if t6 is not None and t6.status_code == 200:
    t = text(t6.text)
    print(t[:900].replace("\n", " / "))
    for k in ("Coal and lignite", "Crude petroleum", "Natural gases", "Refined petroleum"):
        i = t.find(k)
        print("  ...", t[max(0, i - 20):i + 160].replace("\n", " / "))
print("=== GACC: (14) Major Import Commodities table ===")
t14 = get(CE + "/Statics/4b78370a-cd59-4e25-b483-5ef203143664.html")
if t14 is not None and t14.status_code == 200:
    t = text(t14.text)
    print(t[:700].replace("\n", " / "))
    for k in ("Coal", "Crude petroleum", "Natural gas", "Refined petroleum", "Liquefied"):
        for m in re.finditer(k, t):
            print("  ...", t[max(0, m.start() - 10):m.start() + 170].replace("\n", " / "))
            break
print("=== GACC: ColumnId=2 depth: titles of table (14) by page ===")
months = {}
for pg in range(1, 40):
    r = get(f"{CE}/Statistics/Statistics?ColumnId=2&page={pg}")
    if r is None or r.status_code != 200:
        break
    items = re.findall(r'<a[^>]+href="([^"]*Statics[^"]*)"[^>]*>([^<]+)</a>', r.text)
    if not items:
        print("  page", pg, "empty")
        break
    for h, t in items:
        if re.search(r"Major Import Commodities", t):
            months[t.strip()] = h
    last = items[-1][1].strip()[:70]
    if pg % 5 == 0 or pg < 4:
        print(f"  page {pg}: {len(items)} items, last: {last}", flush=True)
print("table (14) titles found:", len(months))
for t, h in list(months.items())[:6] + list(months.items())[-6:]:
    print("   ", t, "->", h[-50:])
