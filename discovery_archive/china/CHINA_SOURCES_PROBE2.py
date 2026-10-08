"""
Probe 2 (manual workflow china_sources_probe2.yml): open the reachable pages found by CHINA_SOURCES_PROBE.py and print
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


print("=== NDRC gas bulletin list ===")
base = "https://www.ndrc.gov.cn/fggz/jjyxtj/"
items = []
for page in ["", "index_1.html", "index_2.html", "index_3.html", "index_4.html"]:
    r = get(base + page)
    if r is not None and r.status_code == 200:
        got = links(r.text, "天然气运行")
        print(f"  page {page or 'index'}: {len(got)} gas bulletins, first {got[:1]}, last {got[-1:]}")
        items += got
print("total gas bulletins listed:", len(items))
show("https://www.ndrc.gov.cn/fggz/jjyxtj/202609/t20260902_1407386.html", 3000)
show("https://www.ndrc.gov.cn/fggz/jjyxtj/202603/t20260330_1404420.html", 2500)

print("=== NEA ===")
r = get("https://www.nea.gov.cn/")
if r is not None:
    for h, t in links(r.text, "统计|用电量|装机|电力工业|发布|数据")[:40]:
        print("   ", t[:60], "->", h[:100])
    for h in re.findall(r'href="([^"]*(?:sjzz|tjsj|xxgk|statistic|zxfb)[^"]*)"', r.text)[:30]:
        print("   section link:", h)
show("https://www.nea.gov.cn/20260920/bd277235282140f4b43e9e0f5f0af3d9/c.html", 3000)
for u in ["https://www.nea.gov.cn/sjzz/tjsj/", "https://www.nea.gov.cn/sjzz/tjsj/index.htm", "https://www.nea.gov.cn/xxgk/",
          "https://www.nea.gov.cn/sjzz/qgdlgyjtj/", "https://www.nea.gov.cn/dl/"]:
    r = get(u)
    if r is not None and r.status_code == 200:
        for h, t in links(r.text, "统计|用电量|装机|电力工业")[:15]:
            print("   ", t[:60], "->", h[:100])
r = get("https://www.nea.gov.cn/search/?q=%E5%85%A8%E5%9B%BD%E7%94%B5%E5%8A%9B%E5%B7%A5%E4%B8%9A%E7%BB%9F%E8%AE%A1%E6%95%B0%E6%8D%AE")
if r is not None:
    print(text(r.text)[:1500])

print("=== GACC English ===")
for u in ["http://english.customs.gov.cn/Statistics/Statistics?ColumnId=6", "http://english.customs.gov.cn/statistics/statistics?ColumnId=6",
          "http://english.customs.gov.cn/Statistics/Statistics?ColumnId=1"]:
    r = get(u)
    if r is not None and r.status_code == 200:
        print(text(r.text)[:1500])
        for h, t in links(r.text, "(?i)crude|natural gas|coal|import|export|major|commodit|table|202")[:40]:
            print("   ", t[:70], "->", h[:110])
