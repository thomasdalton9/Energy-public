"""
NEA hunt 7 (manual workflow nea_hunt7.yml): union of the ds_*.json lists of the NEA press and department columns, classified into
capacity releases (全国电力(工业)统计数据) and consumption releases (全社会用电量), per release month, to find which months have no release.
Writes discovery_archive/results/nea/release_index.json.
"""
import re, json, signal, time
from urllib.parse import urljoin
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
    time.sleep(0.3); return r
def clean(s): return re.sub("<[^>]+>", "", s or "")

SECS = ["/xwfb/", "/sjzz/ghs/", "/sjzz/zhs/", "/sjzz/dls/", "/sjzz/xny/", "/sjzz/mts/", "/sjzz/fgs/", "/sjzz/gjs/", "/sjzz/jgs/", "/sjzz/hds/", "/xwzx/"]
allitems = {}
for sec in SECS:
    r = get("https://www.nea.gov.cn" + sec + "index.htm")
    if r is None or r.status_code != 200: continue
    for ds in sorted(set(re.findall(r'datasource:([0-9a-f]{32})', r.text))):
        rj = get(f"https://www.nea.gov.cn{sec}ds_{ds}.json")
        if rj is None or rj.status_code != 200: continue
        try: its = rj.json().get("datasource") or []
        except Exception: continue
        base = f"https://www.nea.gov.cn{sec}"
        for i in its:
            t = clean(i.get("title"))
            kind = "cap" if re.search(r"全国电力(工业)?统计数据", t) else ("cons" if re.search(r"全社会用电量", t) else None)
            if not kind: continue
            u = urljoin(base, i["publishUrl"])
            if u.startswith("http://www.nea.gov.cn/20"): u = u.replace("http://", "https://")
            allitems.setdefault((kind, i["publishTime"][:10], t), u)
rows = sorted([{"kind": k, "date": d, "title": t, "url": u} for (k, d, t), u in allitems.items()], key=lambda x: (x["kind"], x["date"]))
json.dump(rows, open("discovery_archive/results/nea/release_index.json", "w"), ensure_ascii=False, indent=0)
print("releases:", len(rows), "cap", sum(r["kind"] == "cap" for r in rows), "cons", sum(r["kind"] == "cons" for r in rows))
for kind in ("cap", "cons"):
    print("====", kind)
    have = {}
    for r in rows:
        if r["kind"] == kind: have.setdefault(r["date"][:7], []).append(r["date"][8:] + " " + r["title"][:26])
    y, m = 2021, 1
    while (y, m) <= (2026, 9):
        k = f"{y}-{m:02d}"
        print(k, "; ".join(have.get(k, [])) or "-- none --")
        m += 1
        if m == 13: y, m = y + 1, 1
