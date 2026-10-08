"""
Probe 4 (manual workflow china_sources_probe4.yml): open the reachable pages found by CHINA_SOURCES_PROBE.py and print
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



print("=== NEA: press list pages (electricity consumption, power statistics) ===")
hits = {}
for n in ["index.htm"] + [f"index_{i}.htm" for i in range(1, 31)]:
    r = get("https://www.nea.gov.cn/xwfb/" + n)
    if r is None or r.status_code != 200:
        continue
    got = re.findall(r'<a[^>]+href="([^"]+)"[^>]*>([^<]{3,90})</a>', r.text)
    sel = [(h, t.strip()) for h, t in got if re.search("用电量|电力工业统计|装机", t)]
    print(f"  {n}: {len(got)} links, {len(sel)} selected, e.g. {sel[:2]}", flush=True)
    for h, t in sel:
        hits[h] = t
print("distinct selected:", len(hits))
for h, t in list(hits.items())[:60]:
    print("   ", t[:60], "->", h[:90])
pw = [(h, t) for h, t in hits.items() if "电力工业统计" in t]
if pw:
    show(pw[0][0] if pw[0][0].startswith("http") else "https://www.nea.gov.cn/" + pw[0][0].lstrip("/"), 2500)
uc = [(h, t) for h, t in hits.items() if "用电量" in t]
if uc:
    show(uc[0][0] if uc[0][0].startswith("http") else "https://www.nea.gov.cn/" + uc[0][0].lstrip("/"), 1800)

print("=== NEA: current electricity statistics article (Aug 2026 consumption) and capacity ===")
r = get("https://www.nea.gov.cn/")
if r is not None:
    for h, t in links(r.text, "电力工业|装机|发电量|统计数据"):
        print("   ", t[:70], "->", h[:100])
show("https://www.nea.gov.cn/20260930/cc8c34e7fc454d51aee608d9e75b0d9a/c.html", 2500)

print("=== GACC English: pager mechanics and all table links ===")
r = get("http://english.customs.gov.cn/Statistics/Statistics?ColumnId=1")
if r is not None and r.status_code == 200:
    for m in re.finditer(r"(?i)(pageindex|pagesize|page=|currentpage|\.page\(|pager)", r.text):
        print("   pager hint:", r.text[max(0, m.start() - 80):m.start() + 160].replace("\n", " ")[:240])
        break
    allt = re.findall(r'<a[^>]+href="([^"]*Statics[^"]*)"[^>]*>([^<]+)</a>', r.text)
    print("   table links on page:", len(allt))
    for h, t in allt:
        print("     ", t.strip()[:100], "->", h[-60:])
for u in ["http://english.customs.gov.cn/Statistics/Statistics?ColumnId=1&page=2",
          "http://english.customs.gov.cn/Statistics/Statistics?ColumnId=1&pageNo=2",
          "http://english.customs.gov.cn/Statistics/Statistics?ColumnId=1&p=2",
          "http://english.customs.gov.cn/Statistics/Statistics?ColumnId=1&PageIndex=2"]:
    r = get(u)
    if r is not None and r.status_code == 200:
        t2 = re.findall(r'<a[^>]+href="[^"]*Statics[^"]*"[^>]*>([^<]+)</a>', r.text)
        print("    ", u[-25:], len(r.content), [x.strip()[:50] for x in t2[:2]])
for cid in (2, 3, 4, 5, 7, 8):
    r = get(f"http://english.customs.gov.cn/Statistics/Statistics?ColumnId={cid}")
    if r is not None and r.status_code == 200:
        t2 = re.findall(r'<a[^>]+href="[^"]*Statics[^"]*"[^>]*>([^<]+)</a>', r.text)
        print("    ColumnId", cid, len(t2), [x.strip()[:70] for x in t2[:4]])
