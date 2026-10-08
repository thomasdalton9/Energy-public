"""
Probe 5 (manual workflow china_sources_probe5.yml): open the reachable pages found by CHINA_SOURCES_PROBE.py and print
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



def keep_lines(url, pat, n=40):
    r = get(url)
    if r is None or r.status_code != 200:
        return None
    t = text(r.text)
    ls = [l.strip() for l in t.split("\n") if re.search(pat, l)]
    print("   ", len(t), "chars;", len(ls), "matching lines")
    for l in ls[:n]:
        print("     |", l[:200])
    return r


print("=== NEA press list pages: concise ===")
hits = {}
for n in ["index.htm"] + [f"index_{i}.htm" for i in range(1, 16)]:
    r = get("https://www.nea.gov.cn/xwfb/" + n)
    if r is None or r.status_code != 200:
        print("  ", n, "not 200")
        continue
    got = re.findall(r'<a[^>]+href="([^"]+)"[^>]*>([^<]{3,90})</a>', r.text)
    sel = [(h, t.strip()) for h, t in got if re.search("用电量|电力工业统计|装机", t)]
    dates = sorted(set(re.findall(r"(20\d\d-\d\d)/\d\d/c_|/(20\d{6})/", " ".join(h for h, _ in got))))
    print(f"  {n}: {len(got)} links, {len(sel)} selected; date span {dates[:1]}..{dates[-1:]}", flush=True)
    for h, t in sel:
        hits[h] = t
print("distinct selected:", len(hits))
for h, t in hits.items():
    print("   ", t[:50], "->", h[-60:])

print("=== NEA: capacity article (newest 电力工业统计数据 found) ===")
pw = [h for h, t in hits.items() if "电力工业统计" in t]
if pw:
    keep_lines(pw[0] if pw[0].startswith("http") else "https://www.nea.gov.cn/" + pw[0].lstrip("/"), "装机|太阳能|风电|水电|火电|核电|万千瓦|发布", 40)

print("=== NEA site search for current-month 电力工业统计数据 ===")
for q in ["https://www.nea.gov.cn/sjzz/zhs/gzjz.htm", "https://www.nea.gov.cn/sjzz/zhs/", "https://www.nea.gov.cn/sjzz/zhs/index_1.htm"]:
    r = get(q)
    if r is not None and r.status_code == 200:
        for h, t in links(r.text, "电力工业|用电量|统计")[:15]:
            print("   ", t[:60], "->", h[:100])

print("=== GACC English: Major Imports table (page 2) ===")
r = get("http://english.customs.gov.cn/Statistics/Statistics?ColumnId=1&page=2")
if r is not None and r.status_code == 200:
    for h, t in re.findall(r'<a[^>]+href="([^"]*Statics[^"]*)"[^>]*>([^<]+)</a>', r.text):
        print("   ", t.strip()[:90], "->", h[-50:])
        if "Imports" in t and "USD" in t:
            keep_lines(h if h.startswith("http") else "http://english.customs.gov.cn" + h, "(?i)crude|natural gas|coal|liquefied|petroleum|gas|Commodity|Aug", 30)
            break

print("=== GACC English: ColumnId=2 / 3 (monthly bulletin / tables) full lists ===")
for cid in (2,):
    for pg in (1, 2, 3):
        r = get(f"http://english.customs.gov.cn/Statistics/Statistics?ColumnId={cid}&page={pg}")
        if r is not None and r.status_code == 200:
            for h, t in re.findall(r'<a[^>]+href="([^"]*Statics[^"]*)"[^>]*>([^<]+)</a>', r.text):
                print("   ", t.strip()[:100], "->", h[-45:])
