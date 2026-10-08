"""
NEA hunt 9 (manual workflow nea_hunt9.yml): save the plain text (table rows kept, cells separated by ' | ') and image URLs of every
capacity / consumption release listed in release_index.json to discovery_archive/results/nea/articles.json, as the test input
for the parsers in asia/CHINA_NEA_CAPACITY.py and asia/CHINA_NEA_CONSUMPTION.py.
"""
import re, json, signal, time
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
    r.encoding = "utf-8"
    time.sleep(0.3); return r
def text(html):
    html = re.sub(r"(?s)<(script|style).*?</\1>", " ", html)
    html = re.sub(r"\s+", " ", html)
    html = re.sub(r"</(p|div|tr|li|h\d)>", "\n", html)
    html = re.sub(r"</t[dh]>", " | ", html)
    t = re.sub(r"<[^>]+>", "", html)
    t = re.sub(r"&nbsp;|&emsp;|　", " ", t)
    t = t.replace("&gt;", ">").replace("&lt;", "<").replace("&amp;", "&")
    return re.sub(r"\n\s*\n+", "\n", re.sub(r"[ \t]+", " ", t))
rows = json.load(open("discovery_archive/results/nea/release_index.json"))
out = []
for r in rows:
    p = get(r["url"])
    if p is None or p.status_code != 200:
        print("ERR", r["url"], p.status_code if p is not None else ""); continue
    ims = [x for x in re.findall(r'<img[^>]+src="([^"]+)"', p.text) if not re.search(r"logo|icon|ewm|qr|1\.gif", x, re.I)]
    body = re.search(r'(?s)(<div[^>]+id="detailContent".*?)(<div[^>]+class="(?:fenxiang|share)|$)', p.text)
    out.append({**r, "text": text(body.group(1) if body else p.text), "imgs": ims})
json.dump(out, open("discovery_archive/results/nea/articles.json", "w"), ensure_ascii=False)
print("saved", len(out), "of", len(rows))
