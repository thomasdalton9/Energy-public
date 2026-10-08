"""
Probe 3 (manual workflow china_sources_probe3.yml): open the reachable pages found by CHINA_SOURCES_PROBE.py and print
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



print("=== NDRC: all gas bulletins, dates and first lines ===")
base = "https://www.ndrc.gov.cn/fggz/jjyxtj/"
items = []
for page in ["", "index_1.html", "index_2.html"]:
    r = get(base + page)
    if r is not None and r.status_code == 200:
        items += [(h, t) for h, t in links(r.text, "天然气运行") ]
seen = set()
for h, t in items:
    if h in seen:
        continue
    seen.add(h)
    r = get(base + h.lstrip("./") if h.startswith("./") else h)
    if r is not None and r.status_code == 200:
        body = text(r.text)
        i = body.find("据快报")
        print("  ", t, "|", body[i:i + 330].replace("\n", " "), flush=True)

print("=== NEA: list pages ===")
for u in ["https://www.nea.gov.cn/xwfb/index.htm", "https://www.nea.gov.cn/xwfb/index_2.htm", "https://www.nea.gov.cn/xwfb/index_10.htm"]:
    r = get(u)
    if r is not None and r.status_code == 200:
        for h, t in links(r.text, "用电量|电力工业|装机|统计数据")[:20]:
            print("   ", t[:70], "->", h[:100])
for u in ["https://www.nea.gov.cn/sjzz/dls/index.htm", "https://www.nea.gov.cn/sjzz/zhs/index.htm", "https://www.nea.gov.cn/sjzz/index.htm"]:
    r = get(u)
    if r is not None and r.status_code == 200:
        print(text(r.text)[:600])
        for h, t in links(r.text, ".")[:25]:
            print("   ", t[:70], "->", h[:100])

print("=== GACC English: preliminary release major imports ===")
r = get("http://english.customs.gov.cn/Statistics/Statistics?ColumnId=1")
if r is not None and r.status_code == 200:
    ls = links(r.text, "(?i)major imports|major exports")
    for h, t in ls:
        print("   ", t[:90], "->", h[:100])
    for h, t in ls:
        if "Imports" in t and "USD" in t:
            show(h if h.startswith("http") else "http://english.customs.gov.cn" + h, 4500)
            break
    for pg in (2, 3):
        r2 = get(f"http://english.customs.gov.cn/Statistics/Statistics?ColumnId=1&pageIndex={pg}")
        if r2 is not None and r2.status_code == 200:
            print("  page", pg, [t[:60] for h, t in links(r2.text, "(?i)major imports")][:4])
