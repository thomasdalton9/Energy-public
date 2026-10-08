"""
NEA hunt 4 (manual workflow nea_hunt4.yml): (1) all press-list titles since Dec 2025 (where are the 2026 capacity releases?),
(2) every other NEA column page reachable from the home page: its ds_*.json size, date range, power-statistics hits,
(3) sample articles of both release types (capacity, consumption), URLs resolved against the list page.
"""
import re, signal, time, json
from urllib.parse import urljoin
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"}
S = requests.Session(); S.headers.update(H)
class Hard(Exception): pass
def _alarm(*_): raise Hard("hard limit")
signal.signal(signal.SIGALRM, _alarm)

def get(url, quiet=False):
    signal.alarm(40)
    try:
        r = S.get(url, timeout=(10, 25)); signal.alarm(0)
    except (requests.RequestException, Hard) as e:
        signal.alarm(0); print(f"  ERROR {url}: {type(e).__name__}: {str(e)[:100]}", flush=True); return None
    if r.encoding in (None, "ISO-8859-1"): r.encoding = r.apparent_encoding
    if not quiet: print(f"GET {url} -> {r.status_code} {len(r.content)} bytes", flush=True)
    time.sleep(0.4)
    return r

def text(html):
    html = re.sub(r"(?s)<(script|style).*?</\1>", " ", html)
    html = re.sub(r"\s+", " ", html)
    html = re.sub(r"</(p|div|tr|li|h\d)>", "\n", html)
    html = re.sub(r"</t[dh]>", " | ", html)
    t = re.sub(r"<[^>]+>", "", html)
    t = re.sub(r"&nbsp;|　", " ", t)
    return re.sub(r"\n\s*\n+", "\n", re.sub(r"[ \t]+", " ", t))

def clean(s): return re.sub("<[^>]+>", "", s or "")
PAT = re.compile(r"用电量|电力工业统计|发电装机|装机容量|电力统计")

LIST = "https://www.nea.gov.cn/xwfb/"
r = get(LIST + "ds_4f3484af7ea244e7ab18d094856c82a6.json", quiet=True)
items = r.json().get("datasource") or []
print("=== 1. press titles since 2025-12 ===")
for i in sorted(items, key=lambda x: x.get("publishTime")):
    if i.get("publishTime", "") >= "2025-12-01":
        print(" ", i["publishTime"][:10], clean(i.get("title"))[:70])

print("=== 2. other columns from the home page ===")
home = get("https://www.nea.gov.cn/").text
secs = sorted(set(re.findall(r'href="(?:https?://www\.nea\.gov\.cn)?(/[a-z0-9_/]+/)(?:index\.htm)?"', home)))
print("sections linked on home:", secs)
done = set()
for sec in secs[:60]:
    r = get("https://www.nea.gov.cn" + sec + "index.htm", quiet=True)
    if r is None or r.status_code != 200: continue
    for ds in set(re.findall(r'datasource:([0-9a-f]{32})', r.text)):
        if ds in done: continue
        done.add(ds)
        rj = get(f"https://www.nea.gov.cn{sec}ds_{ds}.json", quiet=True)
        if rj is None or rj.status_code != 200: continue
        try: its = rj.json().get("datasource") or []
        except Exception: continue
        if not its: continue
        ds_dates = sorted(i.get("publishTime", "") for i in its)
        hit = [i for i in its if PAT.search(clean(i.get("title")))]
        print(f"  {sec} ds_{ds[:8]} n={len(its)} {ds_dates[0][:10]}..{ds_dates[-1][:10]} hits={len(hit)}")
        for i in hit:
            if i["publishTime"] >= "2025-12-01" or sec != "/xwfb/":
                print("     ", i["publishTime"][:10], clean(i.get("title"))[:60], i.get("publishUrl"))

print("=== 3. sample articles ===")
def pick(sub, date):
    for i in items:
        if sub in clean(i.get("title")) and i["publishTime"].startswith(date): return i
for sub, date in [("电力工业统计", "2025-12-26"), ("全社会用电量", "2025-12-24"), ("电力工业统计", "2024-08-23")]:
    i = pick(sub, date)
    if not i: print("not found", sub, date); continue
    u = urljoin(LIST, i["publishUrl"])
    r2 = get(u)
    if r2 is not None and r2.status_code == 200: print(text(r2.text)[:3000], "\n-----")
