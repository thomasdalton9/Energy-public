"""
NEA hunt 2 (manual workflow nea_hunt2.yml): (a) full raw HTML of the NEA list page and the Xhwpage JS to find the ds_*.json
data source of the script-rendered list; (b) the gov.cn site-search JSON API for the monthly NEA releases (titles, dates, URLs).
"""
import re, signal, time, json
import requests

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
     "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"}
S = requests.Session(); S.headers.update(H)
class Hard(Exception): pass
def _alarm(*_): raise Hard("hard limit")
signal.signal(signal.SIGALRM, _alarm)

def get(url):
    signal.alarm(40)
    try:
        r = S.get(url, timeout=(10, 25)); signal.alarm(0)
    except (requests.RequestException, Hard) as e:
        signal.alarm(0); print(f"  ERROR {url}: {type(e).__name__}: {str(e)[:100]}", flush=True); return None
    if r.encoding in (None, "ISO-8859-1"): r.encoding = r.apparent_encoding
    print(f"GET {url} -> {r.status_code} {len(r.content)} bytes", flush=True)
    time.sleep(0.5)
    return r

print("=== A. raw list page ===")
r = get("https://www.nea.gov.cn/xwfb/index.htm")
if r is not None: print(r.text)
print("=== B. Xhwpage JS: lines mentioning json / ds_ / datasource ===")
r = get("http://www.nea.gov.cn/2015nyj/xinban/js/nyjcb_dw_es5_2023.js")
if r is not None:
    js = r.text
    for m in re.finditer(r'.{0,200}(ds_|\.json|datasource|dataSource).{0,250}', js):
        print("  ...", m.group(0).replace("\n", " "))
        
print("=== C. guesses for the list JSON ===")
for u in ["https://www.nea.gov.cn/xwfb/ds_32cdf56e9d0b4369af12f31cfc7b30ed.json", "https://www.nea.gov.cn/xwfb/index.json",
          "https://www.nea.gov.cn/xwfb/ds.json"]:
    r = get(u)
    if r is not None and r.status_code == 200: print(r.text[:500])

print("=== D. gov.cn search API ===")
Q = "https://sousuo.www.gov.cn/search-gov/data?t=zhengce&q={q}&timetype=timeqb&mintime=&maxtime=&sort=pubtime&sortType=1&searchfield=title&pcodeJiguan=&childtype=&subchildtype=&tsbq=&pubtimeyear=&puborg=&pcodeYear=&pcodeNum=&filetype=&p={p}&n=20&inpro=&bmfl=&dup=&orpro=&type=gwyzcwjk"
from urllib.parse import quote
seen = set()
for q in ["全社会用电量", "全国电力工业统计数据", "全国累计发电装机容量", "电力工业统计数据", "全社会用电量同比增长"]:
    for p in range(1, 8):
        r = get(Q.format(q=quote(q), p=p))
        if r is None or r.status_code != 200: break
        try: j = r.json()
        except Exception: print(r.text[:300]); break
        cm = (j.get("searchVO") or {}).get("catMap") or {}
        n = 0
        for cat, v in cm.items():
            for it in (v.get("listVO") or []):
                n += 1
                key = it.get("url")
                if key in seen: continue
                seen.add(key)
                print(f"  [{cat}] {it.get('pubtimeStr')} | {re.sub('<[^>]+>', '', it.get('title') or '')} | {it.get('url')} | {it.get('source')}")
        if n == 0: break
print("distinct results:", len(seen))
