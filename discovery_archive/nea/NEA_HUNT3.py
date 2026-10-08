"""
NEA hunt 3 (manual workflow nea_hunt3.yml): fetch the JSON data source behind the script-rendered NEA press list
(https://www.nea.gov.cn/xwfb/ds_4f3484af7ea244e7ab18d094856c82a6.json, found in the list page's data="datasource:..." attribute),
report its size and date range, list the monthly power-statistics releases in it, and print two sample articles.
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

def text(html):
    html = re.sub(r"(?s)<(script|style).*?</\1>", " ", html)
    html = re.sub(r"</(p|div|tr|li|h\d)>", "\n", html)
    html = re.sub(r"</t[dh]>", " | ", html)
    t = re.sub(r"<[^>]+>", "", html)
    t = re.sub(r"&nbsp;|　", " ", t)
    return re.sub(r"\n\s*\n+", "\n", re.sub(r"[ \t]+", " ", t))

r = get("https://www.nea.gov.cn/xwfb/ds_4f3484af7ea244e7ab18d094856c82a6.json")
items = []
if r is not None and r.status_code == 200:
    try:
        j = r.json()
        items = j.get("datasource") or []
        print("keys:", list(j.keys()), "items:", len(items))
        if items: print("sample item:", json.dumps(items[0], ensure_ascii=False)[:600])
        dates = sorted(str(i.get("publishTime")) for i in items)
        print("publishTime range:", dates[:1], dates[-1:])
    except Exception as e:
        print("not JSON:", e, r.text[:400])
pat = re.compile(r"用电量|电力工业统计|发电装机|装机容量")
hits = [i for i in items if pat.search(re.sub("<[^>]+>", "", i.get("showTitle") or i.get("title") or ""))]
print("matching titles:", len(hits))
for i in sorted(hits, key=lambda x: str(x.get("publishTime"))):
    print(" ", i.get("publishTime"), "|", re.sub("<[^>]+>", "", i.get("showTitle") or i.get("title") or ""), "|", i.get("publishUrl"))
shown = {"用电量": 0, "电力工业": 0}
for i in sorted(hits, key=lambda x: str(x.get("publishTime")), reverse=True):
    t = re.sub("<[^>]+>", "", i.get("showTitle") or "")
    k = "用电量" if "用电量" in t else ("电力工业" if "电力工业" in t else None)
    if k and shown[k] < 1:
        shown[k] += 1
        r2 = get(i["publishUrl"])
        if r2 is not None and r2.status_code == 200:
            print(text(r2.text)[:3500])
